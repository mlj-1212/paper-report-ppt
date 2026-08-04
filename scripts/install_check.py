#!/usr/bin/env python3
"""
Environment check and dependency resolver for paper-report-ppt v4.0.

v4.0 is fully self-contained — no ppt-master, no Node.js required.
Only 5 pip packages are needed:
  python-pptx  — PPTX generation (gen_pptx.py)
  PyMuPDF      — PDF parsing (parse_pdf.py)
  python-docx  — speech script DOCX generation (gen_speech_docx.py)
  Pillow       — image handling (gen_pptx.py)
  matplotlib   — formula rendering (render_formula.py, optional)

Exit codes:
  0 = ready
  2 = needs user action (missing dependencies)

Usage:
  python install_check.py
  python install_check.py --json   # machine-readable output
"""

import sys
import json
import platform

EXIT_READY = 0
EXIT_NEEDS_ACTION = 2

# Required packages: (import_name, pip_name, description, required)
PACKAGES = [
    ("pptx", "python-pptx", "PPTX generation", True),
    ("fitz", "PyMuPDF", "PDF parsing", True),
    ("docx", "python-docx", "Speech script DOCX generation", True),
    ("PIL", "Pillow", "Image handling", True),
    ("matplotlib", "matplotlib", "Formula rendering (optional)", False),
]


def check_python_version():
    """Check Python >= 3.8."""
    major, minor = sys.version_info[:2]
    ok = (major, minor) >= (3, 8)
    return {
        "version": f"{major}.{minor}.{sys.version_info[2]}",
        "ok": ok,
        "message": "" if ok else "Python 3.8+ required",
    }


def check_packages():
    """Check all required pip packages."""
    results = {}
    for import_name, pip_name, desc, required in PACKAGES:
        try:
            mod = __import__(import_name)
            ver = getattr(mod, "__version__", "ok")
            results[import_name] = {
                "status": "ok",
                "version": ver,
                "pip_name": pip_name,
                "desc": desc,
                "required": required,
            }
        except ImportError:
            results[import_name] = {
                "status": "missing",
                "pip_name": pip_name,
                "desc": desc,
                "required": required,
            }
    return results


def main():
    use_json = "--json" in sys.argv

    py_info = check_python_version()
    pkg_info = check_packages()

    # Determine readiness
    required_missing = [
        name for name, info in pkg_info.items()
        if info["status"] == "missing" and info["required"]
    ]
    ready = py_info["ok"] and len(required_missing) == 0

    result = {
        "python": py_info,
        "packages": pkg_info,
        "platform": platform.platform(),
        "ready": ready,
    }

    if use_json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return EXIT_READY if ready else EXIT_NEEDS_ACTION

    # Human-readable output
    print("=" * 60)
    print("paper-report-ppt v4.0 环境自检")
    print("=" * 60)
    print(f"\n平台：{platform.platform()}")

    print(f"\n--- Python ---")
    py_status = "✅" if py_info["ok"] else "❌"
    print(f"  {py_status} Python {py_info['version']}")
    if not py_info["ok"]:
        print(f"     {py_info['message']}")

    print(f"\n--- Python 依赖包 ---")
    for name, info in pkg_info.items():
        tag = "必需" if info["required"] else "可选"
        if info["status"] == "ok":
            print(f"  ✅ {info['pip_name']} ({info['desc']}, {tag}) — {info['version']}")
        else:
            mark = "❌" if info["required"] else "⚠️"
            print(f"  {mark} {info['pip_name']} ({info['desc']}, {tag}) — 未安装")

    print(f"\n{'=' * 60}")
    if ready:
        print("✅ 核心依赖就绪，可以开始使用 paper-report-ppt")
        install_cmd = "pip install python-pptx PyMuPDF python-docx Pillow matplotlib"
        print(f"   完整安装命令：{install_cmd}")
    else:
        missing_pip = [pkg_info[n]["pip_name"] for n in required_missing]
        print(f"❌ 缺少必需依赖：{', '.join(missing_pip)}")
        print(f"   请运行：pip install {' '.join(missing_pip)}")
        if not py_info["ok"]:
            print(f"   并升级 Python 至 3.8+")
    print("=" * 60)

    return EXIT_READY if ready else EXIT_NEEDS_ACTION


if __name__ == "__main__":
    sys.exit(main())
