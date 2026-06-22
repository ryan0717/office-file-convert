"""
One-click installer for all dependencies required by batch_convert_v2.py.
一键安装 batch_convert_v2.py 所需的全部依赖。

Usage:
    python install_deps.py
"""

import subprocess
import sys

# Packages to install / 需要安装的第三方包
# - markitdown: Core conversion library (auto-installs mammoth, pdfplumber, etc.)
#              核心转换库（自动安装 mammoth, pdfplumber 等子依赖）
# - openpyxl:   Custom XLSX conversion (merged cells + number formats)
#              XLSX 自定义转换（合并单元格 + 数字格式）
# - pywin32:    Windows COM interface for legacy .doc/.xls/.ppt conversion
#              Windows COM 接口，用于旧格式转换
PACKAGES = [
    "markitdown",
    "openpyxl",
    "pywin32",
]


def main():
    print("=" * 50)
    print("  Dependency Installer / 依赖安装脚本")
    print("=" * 50)
    print(f"\nPython version / 版本: {sys.version}")
    print(f"Python path / 路径   : {sys.executable}\n")

    # Check pip availability / 检查 pip 是否可用
    try:
        subprocess.run(
            [sys.executable, "-m", "pip", "--version"],
            check=True,
            capture_output=True,
        )
    except Exception:
        print("[Error/错误] pip is not available / pip 不可用。")
        print("  Please install Python and ensure pip is configured.")
        print("  请先安装 Python 并确保 pip 已配置。")
        input("Press Enter to exit / 按回车键退出...")
        return

    # Install packages one by one / 逐个安装
    success_count = 0
    fail_count = 0

    for pkg in PACKAGES:
        print(f"Installing / 正在安装: {pkg} ...")
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", pkg],
            capture_output=True,
            text=True,
        )
        if result.returncode == 0:
            print(f"  ✓ {pkg} installed / 安装成功\n")
            success_count += 1
        else:
            print(f"  ✗ {pkg} failed / 安装失败")
            print(f"    {result.stderr.strip()}\n")
            fail_count += 1

    # pywin32 post-install configuration / pywin32 安装后配置
    print("Configuring pywin32 / 正在配置 pywin32 ...")
    try:
        import pythoncom
        print("  ✓ pywin32 is ready / pywin32 已就绪\n")
    except ImportError:
        # Try running pywin32_post_install / 尝试执行 post-install 脚本
        print("  Running pywin32_post_install ...")
        try:
            subprocess.run(
                [sys.executable, "-m", "pywin32_postinstall", "-install"],
                capture_output=True,
            )
            print("  ✓ pywin32 configured / pywin32 配置完成\n")
        except Exception:
            print("  ! pywin32 may need manual setup / 可能需要手动执行:\n")
            print("    python -m pywin32_postinstall -install\n")

    # Verify installation / 验证安装
    print("=" * 50)
    print("  Verification / 安装验证")
    print("=" * 50 + "\n")

    checks = [
        ("markitdown", "from markitdown import MarkItDown"),
        ("openpyxl", "from openpyxl import load_workbook"),
        ("pywin32", "import win32com.client"),
        ("mammoth", "import mammoth"),
        ("pdfplumber", "import pdfplumber"),
        ("pandas", "import pandas"),
    ]

    for name, stmt in checks:
        try:
            exec(stmt)
            print(f"  ✓ {name}")
        except ImportError:
            print(f"  ✗ {name} not ready / 未就绪")

    print(f"\nDone / 安装完成: {success_count} succeeded / 成功, {fail_count} failed / 失败")
    if fail_count > 0:
        print("Please check errors above and install failed packages manually.")
        print("请检查上方错误信息，手动安装失败的包。")

    input("\nPress Enter to exit / 按回车键退出...")


if __name__ == "__main__":
    main()
"""
一键安装 batch_convert_v2.py 所需的全部依赖

使用方法：
    python install_deps.py
"""

import subprocess
import sys

# 需要安装的第三方包
# - markitdown: 核心转换库（会自动安装 mammoth, pdfplumber, pdfminer, pandas 等子依赖）
# - openpyxl:   XLSX 自定义转换（合并单元格 + 数字格式）
# - pywin32:    Windows COM 接口，用于旧格式 .doc/.xls/.ppt 转换
PACKAGES = [
    "markitdown",
    "openpyxl",
    "pywin32",
]


def main():
    print("=" * 50)
    print("  依赖安装脚本")
    print("=" * 50)
    print(f"\nPython 版本: {sys.version}")
    print(f"Python 路径: {sys.executable}\n")

    # 检查 pip 是否可用
    try:
        subprocess.run(
            [sys.executable, "-m", "pip", "--version"],
            check=True,
            capture_output=True,
        )
    except Exception:
        print("[错误] pip 不可用，请先安装 Python 并确保 pip 已配置。")
        input("按回车键退出...")
        return

    # 逐个安装
    success_count = 0
    fail_count = 0

    for pkg in PACKAGES:
        print(f"正在安装: {pkg} ...")
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", pkg],
            capture_output=True,
            text=True,
        )
        if result.returncode == 0:
            print(f"  ✓ {pkg} 安装成功\n")
            success_count += 1
        else:
            print(f"  ✗ {pkg} 安装失败")
            print(f"    {result.stderr.strip()}\n")
            fail_count += 1

    # pywin32 安装后需要执行 post-install 脚本
    print("正在配置 pywin32 ...")
    try:
        import pythoncom
        print("  ✓ pywin32 已就绪\n")
    except ImportError:
        # 尝试执行 pywin32_post_install
        print("  执行 pywin32_post_install ...")
        try:
            subprocess.run(
                [sys.executable, "-m", "pywin32_postinstall", "-install"],
                capture_output=True,
            )
            print("  ✓ pywin32 配置完成\n")
        except Exception as e:
            print(f"  ! pywin32 配置可能需要手动执行: python -m pywin32_postinstall -install\n")

    # 验证安装
    print("=" * 50)
    print("  安装验证")
    print("=" * 50 + "\n")

    checks = [
        ("markitdown", "from markitdown import MarkItDown"),
        ("openpyxl", "from openpyxl import load_workbook"),
        ("pywin32", "import win32com.client"),
        ("mammoth", "import mammoth"),
        ("pdfplumber", "import pdfplumber"),
        ("pandas", "import pandas"),
    ]

    for name, stmt in checks:
        try:
            exec(stmt)
            print(f"  ✓ {name}")
        except ImportError:
            print(f"  ✗ {name} 未就绪")

    print(f"\n安装完成: 成功 {success_count} 个, 失败 {fail_count} 个")
    if fail_count > 0:
        print("请检查上方错误信息，手动安装失败的包。")

    input("\n按回车键退出...")


if __name__ == "__main__":
    main()
