#!/usr/bin/env python3
"""
validate_slides_json.py — slides.json 跨环境一致性验证脚本

验证 slides.json 是否符合 R1-R7 硬性规则和 C1-C8 校验清单。
在 gen_pptx.py 之前运行，确保不同 AI 环境生成的 slides.json 结构一致。

用法:
    python validate_slides_json.py <slides.json> [--manifest <image_manifest_filtered.json>] [--json]

退出码:
    0 = 全部通过
    1 = 有未通过项（errors）
"""
import argparse
import json
import sys
from pathlib import Path


# ═══════════════════════════════════════════════════════════
#  规则定义
# ═══════════════════════════════════════════════════════════

EXPECTED_PAGE_COUNT = 17

# 文献类型 → 板块单一数据源（与 gen_pptx.py 保持一致）
# 不同文献类型对应不同的汇报板块（目录 sections）与导航短标签。
DOC_TYPE_TEMPLATES = {
    "paper": ["研究背景与科学问题", "材料与方法", "主要结果", "讨论与结论"],
    "review": ["研究背景", "研究进展", "主要结论与展望"],
    "thesis": ["研究背景与意义", "研究内容与方法", "研究结果", "总结与展望"],
}

# R6: bullets 数量和长度限制
BULLET_LIMITS = {
    "content": {"min": 3, "max": 4, "max_len": 40},
    "figure": {"min": 2, "max": 3, "max_len": 40},
    "conclusion": {"min": 3, "max": 5, "max_len": 35},
}

# 有效的 page_type
VALID_PAGE_TYPES = {"cover", "toc", "section", "content", "figure", "model", "conclusion", "qa"}


# ═══════════════════════════════════════════════════════════
#  验证核心
# ═══════════════════════════════════════════════════════════

def validate(slides, manifest_filenames=None):
    """验证 slides.json，返回 (errors, warnings)。

    errors:   违反硬性规则（R1-R7），必须修复
    warnings: 建议改进项，不阻断生成
    """
    errors = []
    warnings = []

    if not isinstance(slides, list):
        errors.append("slides.json 必须是数组 (list of slide objects)")
        return errors, warnings

    # ── C1/R1: 总页数（兼容论文/综述/学位不同结构，仅警告）──
    if len(slides) < 8 or len(slides) > 30:
        errors.append(f"[C1/R1] 总页数异常（应为 8-30 页），实际 {len(slides)} 页")
    elif len(slides) < 12 or len(slides) > 22:
        warnings.append(f"[R1] 建议页数 12-22 页，实际 {len(slides)} 页")

    if len(slides) == 0:
        errors.append("slides.json 为空数组")
        return errors, warnings

    # ── C2/R2: 封面含 cn_title + en_title ──
    cover = slides[0]
    if cover.get("page_type") == "cover":
        if "cn_title" not in cover:
            errors.append("[C2/R2] 封面页缺少 cn_title 字段（禁止仅使用 title）")
        if "en_title" not in cover:
            errors.append("[C2/R2] 封面页缺少 en_title 字段")
        if "subtitle" not in cover:
            warnings.append("[R2] 封面页建议包含 subtitle 字段（作者, 期刊, 年份）")
    else:
        errors.append(f"[C2] 第1页应为 cover 类型，实际为 '{cover.get('page_type')}'")

    # ── C3/R5: 第2页 toc sections（兼容论文4/综述3/学位4）──
    if len(slides) > 1:
        toc = slides[1]
        if toc.get("page_type") == "toc":
            sections = toc.get("sections", [])
            if not isinstance(sections, list) or len(sections) == 0:
                errors.append("[C3/R5] 目录页 sections 必须是非空数组")
            elif not (3 <= len(sections) <= 6):
                warnings.append(
                    f"[R5] 目录页 sections 建议 3-6 段，实际 {len(sections)} 段"
                    f"（论文4/综述3/学位4）"
                )
            # doc_type 一致性：若显式声明 doc_type，则 sections 必须匹配模板
            doc_type = toc.get("doc_type")
            if doc_type in DOC_TYPE_TEMPLATES and sections != DOC_TYPE_TEMPLATES[doc_type]:
                errors.append(
                    f"[R5] doc_type='{doc_type}' 的 sections 应为 "
                    f"{DOC_TYPE_TEMPLATES[doc_type]}，实际 {sections}"
                )
            # nav_labels 数量应与 sections 一致
            nav_labels = toc.get("nav_labels")
            if nav_labels and len(nav_labels) != len(sections):
                warnings.append(
                    f"[R5] nav_labels 数量({len(nav_labels)}) 应与 "
                    f"sections 数量({len(sections)}) 一致"
                )
        else:
            errors.append(f"[C3] 第2页应为 toc 类型，实际为 '{toc.get('page_type')}'")

    # ── C4/R1: page_type 序列（结构性，兼容不同 doc_type）──
    actual_types = [s.get("page_type", "unknown") for s in slides]
    if actual_types and actual_types[0] != "cover":
        errors.append(f"[C4/R1] 第1页应为 cover，实际 '{actual_types[0]}'")
    if len(actual_types) >= 2 and actual_types[1] != "toc":
        errors.append(f"[C4/R1] 第2页应为 toc，实际 '{actual_types[1]}'")
    if actual_types and actual_types[-1] != "qa":
        warnings.append(f"[R1] 末页应为 qa（致谢/问答），实际 '{actual_types[-1]}'")
    if "section" not in actual_types:
        warnings.append("[R1] 至少应有 1 个 section 分隔页")
    try:
        first_sec = actual_types.index("section")
    except ValueError:
        first_sec = None
    if first_sec is not None:
        for j in range(first_sec):
            if actual_types[j] in ("content", "figure", "model"):
                errors.append(
                    f"[C4/R1] 第{j+1}页在首个 section 之前出现 "
                    f"'{actual_types[j]}'，内容页须在 section 之后"
                )
    # section 分隔页数应与 toc sections 数量一致
    toc_sec = slides[1] if len(slides) > 1 and slides[1].get("page_type") == "toc" else None
    if toc_sec is not None:
        _n = len(toc_sec.get("sections") or [])
        _d = actual_types.count("section")
        if _n and _d != _n:
            warnings.append(
                f"[R1] section 分隔页数({_d}) 应与 toc sections 数({_n}) 一致"
            )

    # ── 逐页检查 C5-C8, R3-R7 ──
    for i, slide in enumerate(slides):
        page_num = i + 1
        pt = slide.get("page_type", "unknown")

        # page_num 连续性
        actual_pn = slide.get("page_num")
        if actual_pn != page_num:
            errors.append(f"第{page_num}页 page_num 应为 {page_num}，实际 {actual_pn}")

        # page_type 有效性
        if pt not in VALID_PAGE_TYPES:
            errors.append(f"第{page_num}页 page_type '{pt}' 不是有效类型")
            continue

        # C5/C6/C8/R3/R4: figure 页检查
        if pt == "figure":
            ip = slide.get("image_path")
            if not ip:
                errors.append(f"[C5] 第{page_num}页 (figure) 缺少 image_path")
            elif isinstance(ip, list):
                errors.append(f"[C8/R4] 第{page_num}页 (figure) image_path 是数组，违反一图一页规则")

            # C6/R3: image_path 与 manifest 一致
            if manifest_filenames and isinstance(ip, str) and ip:
                if ip not in manifest_filenames:
                    errors.append(f"[C6/R3] 第{page_num}页 image_path '{ip}' 不在 manifest 文件名列表中")

            # R7: figure 页必须有 image_caption 和 sub_title
            if not slide.get("image_caption"):
                warnings.append(f"[R7] 第{page_num}页 (figure) 缺少 image_caption")
            if not slide.get("sub_title") and not slide.get("analysis_title"):
                warnings.append(f"[R7] 第{page_num}页 (figure) 缺少 sub_title")

        # C7: content/figure 页有 sub_title
        if pt in ("content", "figure"):
            if not slide.get("sub_title") and not slide.get("analysis_title"):
                warnings.append(f"[C7] 第{page_num}页 ({pt}) 缺少 sub_title")

        # R7: content 页必须有 conclusion
        if pt == "content":
            if not slide.get("conclusion"):
                warnings.append(f"[R7] 第{page_num}页 (content) 缺少 conclusion 字段")

        # R6: bullets 数量和长度
        if pt in BULLET_LIMITS:
            limits = BULLET_LIMITS[pt]
            bullets = slide.get("bullets", [])
            if bullets:
                if len(bullets) < limits["min"] or len(bullets) > limits["max"]:
                    warnings.append(
                        f"[R6] 第{page_num}页 ({pt}) bullets 数量应为 "
                        f"{limits['min']}-{limits['max']} 条，实际 {len(bullets)} 条"
                    )
                for j, b in enumerate(bullets):
                    if isinstance(b, str) and len(b) > limits["max_len"]:
                        warnings.append(
                            f"[R6] 第{page_num}页 bullets[{j}] 长度 {len(b)} "
                            f"超过 {limits['max_len']} 字符"
                        )

    return errors, warnings


# ═══════════════════════════════════════════════════════════
#  主函数
# ═══════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="validate_slides_json —— slides.json 跨环境一致性验证"
    )
    parser.add_argument("input", help="slides.json 路径")
    parser.add_argument(
        "--manifest", "-m",
        help="image_manifest_filtered.json 路径（用于 C6 检查）",
    )
    parser.add_argument(
        "--json", action="store_true",
        help="输出 JSON 格式报告",
    )
    args = parser.parse_args()

    # 读取 slides.json
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误: 文件不存在: {args.input}")
        sys.exit(1)

    try:
        with open(input_path, "r", encoding="utf-8-sig") as f:
            slides = json.load(f)
    except json.JSONDecodeError as e:
        print(f"错误: JSON 格式不合法: {e}")
        sys.exit(1)

    # 读取 manifest（如果提供）
    manifest_filenames = None
    if args.manifest:
        manifest_path = Path(args.manifest)
        if manifest_path.exists():
            try:
                with open(manifest_path, "r", encoding="utf-8-sig") as f:
                    manifest = json.load(f)
                if isinstance(manifest, list):
                    manifest_filenames = [
                        item.get("filename", "")
                        for item in manifest
                        if isinstance(item, dict)
                    ]
                elif isinstance(manifest, dict):
                    items = manifest.get("images", manifest.get("items", []))
                    manifest_filenames = [
                        item.get("filename", "")
                        for item in items
                        if isinstance(item, dict)
                    ]
            except (json.JSONDecodeError, KeyError):
                pass

    # 验证
    errors, warnings = validate(slides, manifest_filenames)

    # 输出报告
    if args.json:
        report = {
            "passed": len(errors) == 0,
            "errors": errors,
            "warnings": warnings,
            "summary": {
                "total_pages": len(slides),
                "errors_count": len(errors),
                "warnings_count": len(warnings),
            },
        }
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print("=" * 60)
        print("  slides.json 跨环境一致性验证报告")
        print("=" * 60)
        print(f"\n  总页数: {len(slides)}")

        if errors:
            print(f"\n  ❌ 错误 ({len(errors)} 项):")
            for err in errors:
                print(f"    • {err}")

        if warnings:
            print(f"\n  ⚠️  警告 ({len(warnings)} 项):")
            for warn in warnings:
                print(f"    • {warn}")

        if not errors and not warnings:
            print("\n  ✅ 全部通过！符合 R1-R7 规则和 C1-C8 校验清单。")
        elif not errors:
            print(f"\n  ✅ 硬性规则全部通过，{len(warnings)} 项建议改进。")
        else:
            print(f"\n  ❌ {len(errors)} 项错误必须修复，{len(warnings)} 项警告建议改进。")

        print("\n" + "=" * 60)

    sys.exit(0 if not errors else 1)


if __name__ == "__main__":
    main()
