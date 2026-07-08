"""
打包脚本：将 Seewo Teacher Helper 打包为单个 exe 文件
使用方法：python build_exe.py
"""
import os
import sys
import subprocess


def build():
    # 确保在项目根目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    # 主程序路径
    main_script = os.path.join(script_dir, "app", "main.py")

    if not os.path.exists(main_script):
        print(f"错误：找不到主程序文件 {main_script}")
        sys.exit(1)

    # 输出目录
    output_dir = os.path.join(script_dir, "dist")

    # PyInstaller 命令
    cmd = [
        "pyinstaller",
        "--onefile",                # 打包为单个 exe
        "--windowed",               # 无控制台窗口（GUI模式）
        "--name", "SeewoTeacherHelper",
        "--distpath", output_dir,
        "--workpath", os.path.join(script_dir, "build"),
        "--specpath", script_dir,
        "--add-data", f"app{os.pathsep}app",  # 添加 app 目录
        "--hidden-import", "PyQt5.sip",
        "--clean",
        "--noconfirm",
        main_script
    ]

    print("=" * 60)
    print("  开始打包 Seewo Teacher Helper...")
    print("=" * 60)
    print(f"  输出目录: {output_dir}")
    print("=" * 60)
    print()

    try:
        subprocess.run(cmd, check=True)
        print()
        print("=" * 60)
        print("  ✅ 打包成功！")
        print(f"  📁 exe 文件位于: {output_dir}")
        print("=" * 60)
    except subprocess.CalledProcessError as e:
        print()
        print("=" * 60)
        print(f"  ❌ 打包失败：{e}")
        print("=" * 60)
        sys.exit(1)
    except FileNotFoundError:
        print()
        print("=" * 60)
        print("  ❌ 未找到 pyinstaller，请先安装：")
        print("     pip install pyinstaller")
        print("=" * 60)
        sys.exit(1)


if __name__ == "__main__":
    build()
