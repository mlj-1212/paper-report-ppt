#!/usr/bin/env python3
"""
Environment check and dependency resolver for paper-report-ppt.

IMPORTANT: `ppt-master` is a TRAE built-in skill, NOT a public GitHub repo.
This script does NOT attempt to clone a non-existent repository. Instead it:
  1. Detects the running environment (TRAE / Claude Code / Cursor / Codex)
  2. Scans common skills directories for `ppt-master`
  3. If found: reports the resolved path and exits 0 (ready)
  4. If missing: prints environment-specific guidance and exits 2 (needs user action)
  5. If no filesystem: exits 3 (unsupported environment)

Exit codes:
  0 = ready, PPT_MASTER_DIR resolved
  2 = needs user action (ppt-master not found, guidance printed)
  3 = unsupported environment (no filesystem / pure chatbot)

Usage:
  python install_check.py
  python install_check.py --skills-dir /custom/skills/path
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
    Search all candidate skills directories for ppt-master.
    Returns (path, skills_dir) if found, else (None, None).
    """
    # 1. Explicit env var wins
    env_path = os.environ.get("PPT_MASTER_DIR")
    if env_path:
        p = Path(env_path)
        if (p / PPT_MASTER_ENTRY).exists():
            return p, p.parent

    # 2. Scan candidate skills dirs
    for skills_dir in find_skills_dirs():
        candidate = skills_dir / PPT_MASTER_SKILL
        if (candidate / PPT_MASTER_ENTRY).exists():
            return candidate, skills_dir

    return None, None


def detect_environment():
    """Best-effort detection of the running AI environment."""
    # Check for TRAE-specific paths
    home = Path.home()
    if (home / ".trae-cn").exists() or (home / ".trae").exists():
        return "trae"
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA", "")
        if "TRAE" in appdata and Path(appdata).exists():
            # Check if TRAE SOLO CN dir exists
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
        ("matplotlib", "公式渲染（可选）"),
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
    ppt_master_path, skills_dir = find_ppt_master()

    # Build result
    result = {
        "environment": env,
        "ppt_master": {
            "found": ppt_master_path is not None,
            "path": str(ppt_master_path) if ppt_master_path else None,
            "skills_dir": str(skills_dir) if skills_dir else None,
        },
        "python": check_python_runtime(),
        "node": check_node_runtime(),
        "platform": platform.platform(),
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
        print(f"  ✅ 已找到 -> {ppt_master_path}")
        print(f"  Skills 目录：{skills_dir}")
        print(f"  建议：将 PPT_MASTER_DIR 设为 {ppt_master_path}")
    else:
        print(f"  ❌ 未找到 ppt-master skill")
        scanned = find_skills_dirs()
        print(f"\n  已扫描以下目录：")
        for d in scanned:
            exists = "存在" if d.exists() else "不存在"
            print(f"    [{exists}] {d}")

        guidance = get_install_guidance(env)
        print(f"\n--- 安装指引（{guidance['title']}）---")
        for i, step in enumerate(guidance["steps"], 1):
            print(f"  {i}. {step}")

        print(f"\n  ⚠️ 注意：ppt-master 是 TRAE 内置 skill，不是公开 GitHub 仓库。")
        print(f"     不要尝试 git clone https://github.com/trae-ai/ppt-master.git")
        print(f"     （该地址不存在，克隆会失败）")

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
        print("=" * 60)
        return EXIT_READY
    else:
        print("❌ ppt-master 未就绪，请按上方指引安装后重新运行本脚本")
        print("=" * 60)
        return EXIT_NEEDS_ACTION


if __name__ == "__main__":
    sys.exit(main())
