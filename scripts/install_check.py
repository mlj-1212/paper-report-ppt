#!/usr/bin/env python3
"""
Environment check and dependency resolver for paper-report-ppt.

ppt-master is BUNDLED in vendor/ppt-master/ — users do NOT need to install it
separately. This script:
  1. Checks vendor/ppt-master/ (bundled, highest priority)
  2. Falls back to TRAE built-in or other skills directories
  3. Reports the resolved PPT_MASTER_DIR and checks Python/Node runtimes

Exit codes:
  0 = ready, PPT_MASTER_DIR resolved
  2 = needs user action (ppt-master not found anywhere)
  3 = unsupported environment (no filesystem / pure chatbot)

Usage:
  python install_check.py
  python install_check.py --json   # machine-readable output
"""

import os
import sys
import shutil
import json
import platform
from pathlib import Path


# Exit codes
EXIT_READY = 0
EXIT_NEEDS_ACTION = 2
EXIT_UNSUPPORTED = 3

# Skill name to look for
PPT_MASTER_SKILL = "ppt-master"
PPT_MASTER_ENTRY = "SKILL.md"


def find_skills_dirs():
    """
    Return a list of candidate skills directories, in priority order.
    Includes both existing dirs and common default locations.
    """
    home = Path.home()
    candidates = []

    # Unix-style home locations
    candidates.append(home / ".trae-cn" / "skills")
    candidates.append(home / ".trae" / "skills")
    candidates.append(home / ".claude" / "skills")
    candidates.append(home / ".cursor" / "skills")
    candidates.append(home / ".codex" / "skills")

    # Windows AppData locations
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA")
        if appdata:
            candidates.append(Path(appdata) / "TRAE SOLO CN" / "skills")
            candidates.append(Path(appdata) / "TRAE" / "skills")
        localappdata = os.environ.get("LOCALAPPDATA")
        if localappdata:
            candidates.append(Path(localappdata) / "TRAE" / "skills")
            candidates.append(Path(localappdata) / "TRAE SOLO CN" / "skills")

    # Environment variable override
    env_skills = os.environ.get("SKILLS_DIR")
    if env_skills:
        candidates.insert(0, Path(env_skills))

    # De-duplicate while preserving order
    seen = set()
    unique = []
    for c in candidates:
        try:
            resolved = c.resolve()
        except Exception:
            resolved = c
        if resolved not in seen:
            seen.add(resolved)
            unique.append(c)
    return unique


def find_ppt_master():
    """
    Search for ppt-master in priority order:
    1. Environment variable PPT_MASTER_DIR
    2. vendor/ppt-master (bundled with this skill)
    3. TRAE / Claude / Cursor / Codex skills directories
    Returns (path, source) if found, else (None, None).
    """
    # 1. Explicit env var wins
    env_path = os.environ.get("PPT_MASTER_DIR")
    if env_path:
        p = Path(env_path)
        if (p / PPT_MASTER_ENTRY).exists():
            return p, "env-var"

    # 2. Bundled vendor version (highest priority for portability)
    script_dir = Path(__file__).resolve().parent
    skill_root = script_dir.parent  # paper-report-ppt/
    vendor_path = skill_root / "vendor" / PPT_MASTER_SKILL
    if (vendor_path / PPT_MASTER_ENTRY).exists():
        return vendor_path, "vendor"

    # 3. Scan candidate skills dirs (TRAE built-in, Claude, Cursor, etc.)
    for skills_dir in find_skills_dirs():
        candidate = skills_dir / PPT_MASTER_SKILL
        if (candidate / PPT_MASTER_ENTRY).exists():
            return candidate, str(skills_dir)

    return None, None


def detect_environment():
    """Best-effort detection of the running AI environment.

    检测优先级：
    1. 环境变量 AI_ENV（用户可显式指定）
    2. WorkBuddy（检测 .workbuddy 目录或 WORKBUDDY 环境变量）
    3. TRAE（检测 .trae-cn 目录，但需排除 WorkBuddy 嵌套 TRAE 的情况）
    4. Claude Code / Cursor / Codex
    5. unknown
    """
    home = Path.home()

    # 0. 显式环境变量优先
    env_explicit = os.environ.get("AI_ENV", "").strip().lower()
    if env_explicit:
        return env_explicit

    # 1. WorkBuddy 检测（优先于 TRAE，因为 WorkBuddy 可能运行在装了 TRAE 的机器上）
    if (home / ".workbuddy").exists() or os.environ.get("WORKBUDDY_HOME"):
        return "workbuddy"

    # 2. TRAE 检测
    if (home / ".trae-cn").exists() or (home / ".trae").exists():
        return "trae"
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA", "")
        if "TRAE" in appdata and Path(appdata).exists():
            if (Path(appdata) / "TRAE SOLO CN").exists():
                return "trae"

    if (home / ".claude").exists():
        return "claude-code"
    if (home / ".cursor").exists():
        return "cursor"
    if (home / ".codex").exists():
        return "codex"

    return "unknown"


def get_install_guidance(env):
    """
    Return environment-specific guidance for obtaining ppt-master.
    ppt-master is a TRAE built-in skill — it is NOT on public GitHub.
    """
    guidance = {
        "trae": {
            "title": "TRAE 环境",
            "steps": [
                "ppt-master 是 TRAE 内置 skill，通常已预装。",
                "若未检测到，请在 TRAE 的 Skill 管理面板检查是否启用 ppt-master。",
                "或重新安装/更新 TRAE 以恢复内置 skill。",
                "安装后重新运行本脚本确认。",
            ],
        },
        "claude-code": {
            "title": "Claude Code 环境",
            "steps": [
                "ppt-master 是 TRAE 内置 skill，不在公开 GitHub 仓库。",
                "获取方式一：在安装了 TRAE 的机器上，从 ~/.trae-cn/skills/ppt-master 复制整个目录到 ~/.claude/skills/ppt-master",
                "获取方式二：从 TRAE 社区或 Skill 市场下载 ppt-master 的打包版本",
                "复制时保留完整目录结构（含 scripts/ references/ 等子目录）",
                "安装后重新运行本脚本确认。",
            ],
        },
        "cursor": {
            "title": "Cursor 环境",
            "steps": [
                "ppt-master 是 TRAE 内置 skill，不在公开 GitHub 仓库。",
                "获取方式一：在安装了 TRAE 的机器上，从 ~/.trae-cn/skills/ppt-master 复制整个目录到 ~/.cursor/skills/ppt-master",
                "获取方式二：从 TRAE 社区或 Skill 市场下载 ppt-master 的打包版本",
                "复制时保留完整目录结构（含 scripts/ references/ 等子目录）",
                "安装后重新运行本脚本确认。",
            ],
        },
        "codex": {
            "title": "Codex 环境",
            "steps": [
                "ppt-master 是 TRAE 内置 skill，不在公开 GitHub 仓库。",
                "获取方式一：在安装了 TRAE 的机器上，从 ~/.trae-cn/skills/ppt-master 复制整个目录到 ~/.codex/skills/ppt-master",
                "获取方式二：从 TRAE 社区或 Skill 市场下载 ppt-master 的打包版本",
                "复制时保留完整目录结构（含 scripts/ references/ 等子目录）",
                "安装后重新运行本脚本确认。",
            ],
        },
        "unknown": {
            "title": "未知环境",
            "steps": [
                "ppt-master 是 TRAE 内置 skill，不在公开 GitHub 仓库。",
                "推荐在 TRAE 环境中使用本 skill（ppt-master 已预装）。",
                "其他环境：从安装了 TRAE 的机器复制 ~/.trae-cn/skills/ppt-master 整个目录到你的 skills 目录。",
                "skills 目录常见位置：~/.trae/skills/、~/.claude/skills/、~/.cursor/skills/",
                "复制时保留完整目录结构（含 scripts/ references/ 等子目录）",
                "安装后重新运行本脚本确认。",
            ],
        },
    }
    return guidance.get(env, guidance["unknown"])


def check_python_runtime():
    """Check Python version and key packages."""
    info = {"version": sys.version.split()[0], "packages": {}}
    packages = [
        ("matplotlib", "公式渲染（路径A可选）"),
        ("pptx", "PPTX直接生成（路径B必需）"),
        ("fitz", "PDF解析（PyMuPDF，路径B必需）"),
    ]
    for pkg, desc in packages:
        try:
            mod = __import__(pkg)
            ver = getattr(mod, "__version__", "ok")
            info["packages"][pkg] = {"status": "ok", "version": ver, "desc": desc}
        except ImportError:
            info["packages"][pkg] = {"status": "missing", "desc": desc}
    return info


def check_node_runtime():
    """Check Node.js availability for DOCX speech generation."""
    info = {"available": bool(shutil.which("node")), "npm": bool(shutil.which("npm"))}
    if info["available"]:
        try:
            import subprocess
            r = subprocess.run(["node", "--version"], capture_output=True, text=True, timeout=5)
            info["version"] = r.stdout.strip()
        except Exception:
            info["version"] = "unknown"
    return info


def main():
    use_json = "--json" in sys.argv

    # Detect environment
    env = detect_environment()

    # Scan for ppt-master
    ppt_master_path, source = find_ppt_master()

    # Build result
    # Determine recommended generation path
    is_trae = env == "trae"
    has_pptx = False
    has_fitz = False
    # We need to check packages before building result, so do it inline
    try:
        __import__("pptx")
        has_pptx = True
    except ImportError:
        pass
    try:
        __import__("fitz")
        has_fitz = True
    except ImportError:
        pass

    if is_trae:
        recommended_path = "A (SVG pipeline)"
    elif has_pptx and has_fitz:
        recommended_path = "B (direct generation)"
    else:
        recommended_path = "B (needs pip install python-pptx PyMuPDF)"

    # 非 TRAE 环境一律推荐路径 B，即使 ppt-master 可用
    if env != "trae":
        if has_pptx and has_fitz:
            recommended_path = "B (direct generation)"
        else:
            recommended_path = "B (needs pip install python-pptx PyMuPDF)"

    result = {
        "environment": env,
        "ppt_master": {
            "found": ppt_master_path is not None,
            "path": str(ppt_master_path) if ppt_master_path else None,
            "source": source if source else None,
        },
        "python": check_python_runtime(),
        "node": check_node_runtime(),
        "platform": platform.platform(),
        "recommended_path": recommended_path,
    }

    if use_json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return EXIT_READY if ppt_master_path else EXIT_NEEDS_ACTION

    # Human-readable output
    print("=" * 60)
    print("paper-report-ppt 环境自检")
    print("=" * 60)
    print(f"\n运行环境：{env}")
    print(f"平台：{platform.platform()}")

    print("\n--- ppt-master 检测 ---")
    if ppt_master_path:
        source_label = {
            "vendor": "内置 (vendor/ppt-master/)",
            "env-var": "环境变量 PPT_MASTER_DIR",
        }.get(source, source)
        print(f"  ✅ 已找到 -> {ppt_master_path}")
        print(f"  来源：{source_label}")
        print(f"  建议：将 PPT_MASTER_DIR 设为 {ppt_master_path}")
    else:
        print(f"  ❌ 未找到 ppt-master")
        print(f"\n  vendor/ppt-master/ 目录不存在，可能克隆不完整。")
        print(f"  请重新克隆仓库：")
        print(f"    git clone https://github.com/mlj-1212/paper-report-ppt.git")
        print(f"\n  或检查环境变量 PPT_MASTER_DIR 是否指向有效的 ppt-master 目录。")

    print("\n--- Python 运行时 ---")
    py = result["python"]
    print(f"  版本：{py['version']}")
    for pkg, info in py["packages"].items():
        status = "✅" if info["status"] == "ok" else "⚠️（可选）"
        ver = info.get("version", "")
        print(f"  {pkg} ({info['desc']}): {status} {ver}")

    print("\n--- Node.js 运行时（演讲稿生成）---")
    nd = result["node"]
    if nd["available"]:
        print(f"  node ✅ {nd.get('version', '')}")
        print(f"  npm  {'✅' if nd['npm'] else '❌'}")
    else:
        print(f"  node ❌ 未安装（演讲稿 DOCX 生成需要 Node.js + docx 包）")
        print(f"  安装：https://nodejs.org/")

    print("\n" + "=" * 60)
    if ppt_master_path:
        print("✅ 核心依赖就绪，可以开始使用 paper-report-ppt")
        print(f"   PPT_MASTER_DIR = {ppt_master_path}")
    else:
        print("❌ ppt-master 未就绪，请重新克隆仓库或检查环境变量")
        print("=" * 60)
        return EXIT_NEEDS_ACTION

    # Generation path recommendation
    print(f"\n--- 推荐生成路径 ---")
    if is_trae:
        print(f"  路径 A（SVG 管线）✅ 推荐 — TRAE 环境深度集成 ppt-master，质量最高")
    elif has_pptx and has_fitz:
        print(f"  路径 B（直接生成）✅ 推荐 — gen_pptx.py 就绪，所有非 TRAE 环境通用")
    else:
        missing = []
        if not has_pptx:
            missing.append("python-pptx")
        if not has_fitz:
            missing.append("PyMuPDF")
        print(f"  路径 B（直接生成）⚠️ 需要安装: pip install {' '.join(missing)}")

    print("=" * 60)


if __name__ == "__main__":
    sys.exit(main())
