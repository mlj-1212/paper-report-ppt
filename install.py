#!/usr/bin/env python3
"""
paper-report-ppt 一键安装脚本

自动完成：依赖安装 + 环境自检

用法：
  python install.py
"""

import sys
import subprocess

# (import_name, pip_name, required)
PACKAGES = [
    ("pptx", "python-pptx", True),
    ("fitz", "PyMuPDF", True),
    ("docx", "python-docx", True),
    ("PIL", "Pillow", True),
    ("matplotlib", "matplotlib", False),
]

PIPS = [p[1] for p in PACKAGES if p[2]] + [PACKAGES[-1][1]]  # all including optional


def check_python():
    major, minor = sys.version_info[:2]
    if (major, minor) < (3, 8):
        print(f"❌ Python {major}.{minor} 版本过低，需要 3.8+")
        sys.exit(1)
    print(f"✅ Python {major}.{minor}.{sys.version_info[2]}")


def check_missing():
    """Return list of (pip_name, required) for missing packages."""
    missing = []
    for import_name, pip_name, required in PACKAGES:
        try:
            __import__(import_name)
        except ImportError:
            missing.append((pip_name, required))
    return missing


def install_all():
    """pip install all packages."""
    print("\n📦 正在安装依赖包...")
    cmd = [sys.executable, "-m", "pip", "install"] + PIPS
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print("❌ 安装失败，请手动运行：")
        print(f"   pip install {' '.join(PIPS)}")
        print(f"\n错误信息：\n{result.stderr[-500:]}")
        sys.exit(1)
    print("✅ 依赖包安装完成")


def verify():
    """Run install_check.py for final verification."""
    print("\n🔍 运行环境自检...")
    try:
        result = subprocess.run(
            [sys.executable, "scripts/install_check.py"],
            capture_output=True, text=True
        )
        print(result.stdout)
        return result.returncode == 0
    except Exception:
        # Fallback: inline check
        missing = check_missing()
        required_missing = [p for p, r in missing if r]
        if required_missing:
            print(f"❌ 仍有缺失：{', '.join(required_missing)}")
            return False
        print("✅ 核心依赖就绪")
        return True


def main():
    print("=" * 50)
    print("  paper-report-ppt 一键安装")
    print("=" * 50)

    # Step 1: Python version
    print("\n🐍 检查 Python 版本...")
    check_python()

    # Step 2: Check what's missing
    missing = check_missing()
    if not missing:
        print("\n✅ 所有依赖已安装，无需操作")
    else:
        missing_names = [p for p, _ in missing]
        print(f"\n📋 缺少：{', '.join(missing_names)}")
        install_all()

    # Step 3: Verify
    ok = verify()
    if ok:
        print("\n" + "=" * 50)
        print("🎉 安装完成！")
        print("=" * 50)
        print("\n现在可以使用了：")
        print("  上传文献 PDF，告诉 AI「生成组会汇报 PPT」")
    else:
        print("\n❌ 安装未完成，请查看上方错误信息")
        sys.exit(1)


if __name__ == "__main__":
    main()
