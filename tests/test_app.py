"""测试用程序：置顶常驻弹窗，提示正在运行"""
import tkinter as tk
import sys
import os

def main():
    app_name = os.path.splitext(os.path.basename(sys.argv[0]))[0]

    root = tk.Tk()
    root.title(app_name)
    root.attributes("-topmost", True)  # 保持在屏幕最上方
    root.resizable(False, False)

    # 窗口大小和位置
    win_width, win_height = 360, 180
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    x = (screen_w - win_width) // 2
    y = (screen_h - win_height) // 2
    root.geometry(f"{win_width}x{win_height}+{x}+{y}")

    # 主框架
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)

    # 图标/标题
    tk.Label(frame, text="🟢", font=("Segoe UI", 36)).pack(pady=(0, 5))

    tk.Label(
        frame,
        text=f"「{app_name}」正在运行",
        font=("Segoe UI", 14, "bold"),
        fg="#2e7d32"
    ).pack()

    tk.Label(
        frame,
        text=f"PID: {os.getpid()}",
        font=("Segoe UI", 10),
        fg="#888"
    ).pack(pady=(5, 0))

    tk.Label(
        frame,
        text="关闭此窗口即可停止程序",
        font=("Segoe UI", 9),
        fg="#aaa"
    ).pack(pady=(10, 0))

    # 关闭按钮
    close_btn = tk.Button(
        frame,
        text="关闭程序",
        font=("Segoe UI", 11),
        bg="#e74c3c",
        fg="white",
        relief=tk.FLAT,
        padx=20,
        pady=4,
        command=root.destroy
    )
    close_btn.pack(pady=(10, 0))

    root.mainloop()

if __name__ == "__main__":
    main()
