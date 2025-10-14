"""
使用 PyInstaller 将项目打包为 Windows 可执行文件（exe）。

用法：
  1) 安装依赖：python -m pip install pyinstaller
  2) 运行：python build_exe.py
  3) 输出：dist/pdf_svg_gui.exe

可选：将 --icon 设置为你的 .ico 图标路径。
"""

import sys
from pathlib import Path

try:
    import PyInstaller.__main__ as pyim
except Exception as e:
    print("未安装 PyInstaller，请先运行: python -m pip install pyinstaller")
    sys.exit(1)


def main():
    root = Path(__file__).resolve().parent
    entry = str(root / "pdf_svg_gui.py")

    opts = [
        entry,
        "--name",
        "pdf_svg_gui",
        "--onefile",
        "--windowed",
        "--noconfirm",
        "--clean",
        # 解决常见隐式导入（图像格式与 GUI）
        "--hidden-import=fitz",
        "--hidden-import=PIL.Image",
        "--hidden-import=PIL.PngImagePlugin",
        "--hidden-import=PIL.WebPImagePlugin",
        "--hidden-import=PIL.JpegImagePlugin",
        "--hidden-import=tkinter",
    ]

    # 如存在 CairoSVG，则加入隐式导入以保证打包时可用
    try:
        import cairosvg  # noqa: F401
    except Exception:
        pass
    else:
        opts.append("--hidden-import=cairosvg")

    # 可选：设置图标（将路径替换为你的 .ico 文件）
    icon_path = root / "app.ico"
    if icon_path.exists():
        opts.extend(["--icon", str(icon_path)])

    print("运行 PyInstaller：", " ".join(opts))
    pyim.run(opts)


if __name__ == "__main__":
    main()