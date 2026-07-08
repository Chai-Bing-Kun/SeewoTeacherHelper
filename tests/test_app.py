"""
测试用程序：创建一个置顶窗口，运行期间始终保持在最顶层
用于测试 Seevo Teacher Helper 的外置程序检测和重启功能
"""
import sys
import tkinter as tk


def main():
    root = tk.Tk()
    root.title("测试程序 - 置顶窗口")
    root.geometry("400x300+200+200")

    # 设置窗口置顶
    root.attributes("-topmost", True)

    # 窗口内容
    label = tk.Label(root, text="测试程序运行中", font=("微软雅黑", 20, "bold"), fg="#2c3e50")
    label.pack(expand=True)

    info = tk.Label(root, text="此窗口始终置顶\n关闭此窗口即停止程序", font=("微软雅黑", 12), fg="#7f8c8d")
    info.pack()

    status = tk.Label(root, text="进程名: test_app.exe", font=("微软雅黑", 10), fg="#95a5a6")
    status.pack(side="bottom", pady=10)

    # 点击关闭时真正退出
    def on_close():
        root.destroy()
        sys.exit(0)

    root.protocol("WM_DELETE_WINDOW", on_close)

    root.mainloop()


if __name__ == "__main__":
    main()
