#!/bin/bash
# ──────────────────────────────────────────────────────────
# 法律评估意见书生成系统 — Mac 启动脚本
# 双击即可启动，无需手动输入命令
# ──────────────────────────────────────────────────────────

set -e

# 切换到脚本所在目录（项目根目录）
cd "$(dirname "$0")"
PROJECT_DIR="$(pwd)"

echo "═══════════════════════════════════════════════════════════"
echo "  法律评估意见书生成系统"
echo "  Law Firm Legal Assessment Report Generator"
echo "═══════════════════════════════════════════════════════════"
echo ""

# ── 检测 Python3 ────────────────────────────────────────
PYTHON=""
for cmd in python3 python; do
    if command -v "$cmd" &>/dev/null; then
        ver=$("$cmd" --version 2>&1)
        echo "✓ 找到 $ver"
        PYTHON="$cmd"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    echo "✗ 未找到 Python，请先安装 Python 3.9+"
    echo "  下载地址：https://www.python.org/downloads/"
    echo ""
    read -p "按回车键退出..."
    exit 1
fi

# ── 创建/修复虚拟环境 ──────────────────────────────────
VENV_DIR="$PROJECT_DIR/venv"
VENV_PYTHON="$VENV_DIR/bin/python"
NEED_CREATE=false

if [ ! -d "$VENV_DIR" ] || [ ! -f "$VENV_PYTHON" ]; then
    NEED_CREATE=true
else
    # 检查 venv 中的 Python 是否能正常运行（换电脑后路径可能失效）
    if ! "$VENV_PYTHON" --version &>/dev/null; then
        echo "⚠ 虚拟环境已损坏（路径失效），正在重建..."
        rm -rf "$VENV_DIR"
        NEED_CREATE=true
    fi
fi

if [ "$NEED_CREATE" = true ]; then
    echo ""
    echo "⏳ 正在创建 Python 虚拟环境..."
    "$PYTHON" -m venv "$VENV_DIR"
    echo "✓ 虚拟环境创建完成"
fi

echo "✓ 使用虚拟环境：$VENV_PYTHON"

# ── 检查 Xcode 命令行工具（lxml 编译需要）────────────
if ! xcode-select -p &>/dev/null; then
    echo ""
    echo "⚠ 未检测到 Xcode 命令行工具（lxml 依赖编译需要）"
    echo "  正在尝试安装（可能需要几分钟）..."
    xcode-select --install 2>/dev/null || true
    echo "  安装完成后请重新双击启动脚本"
    read -p "按回车键退出..."
    exit 1
fi

# ── 安装依赖 ──────────────────────────────────────────
if [ -f "$PROJECT_DIR/webapp/requirements.txt" ]; then
    echo ""
    echo "⏳ 安装依赖（首次可能需要几分钟编译 lxml）..."
    set +e
    PIP_OUTPUT=$("$VENV_PYTHON" -m pip install -r "$PROJECT_DIR/webapp/requirements.txt" 2>&1)
    PIP_EXIT=$?
    set -e
    if [ $PIP_EXIT -ne 0 ]; then
        echo "✗ 依赖安装失败："
        echo "$PIP_OUTPUT" | tail -20
        echo ""
        echo "常见原因："
        echo "  1. 网络问题 → 检查网络连接后重试"
        echo "  2. Xcode 命令行工具未完整安装 → 运行 xcode-select --install"
        echo "  3. Python 版本过低 → 需要 Python 3.9+"
        read -p "按回车键退出..."
        exit 1
    fi
    echo "✓ 依赖已就绪"
fi

# ── 启动服务 ──────────────────────────────────────────
echo ""
echo "═══════════════════════════════════════════════════════════"
echo "  正在启动服务器..."
echo "  浏览器将自动打开 http://localhost:8000"
echo "  关闭此窗口即可停止服务器"
echo "═══════════════════════════════════════════════════════════"
echo ""

# 打开浏览器
sleep 1
open "http://localhost:8000" 2>/dev/null || true

# 启动 FastAPI（在 webapp 目录下运行）
cd "$PROJECT_DIR/webapp"
"$VENV_PYTHON" main.py
