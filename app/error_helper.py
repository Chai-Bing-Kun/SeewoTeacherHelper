"""
错误处理工具模块
提供友好的错误弹窗，附带通俗易懂的解决建议
"""

import sys
import os
import traceback
from PyQt5.QtWidgets import QMessageBox


# ============================================================
# 常见错误类型 → 解决建议 映射表
# ============================================================
ERROR_SUGGESTIONS = {
    # ---------- 文件/路径相关 ----------
    "FileNotFoundError": {
        "title": "文件未找到",
        "message": "程序找不到需要的文件。",
        "suggestion": (
            "解决方法：\n"
            "1. 检查文件是否被移动或删除了\n"
            "2. 重新选择正确的文件路径\n"
            "3. 如果问题持续，请重新安装本程序"
        )
    },
    "PermissionError": {
        "title": "没有权限",
        "message": "程序没有权限访问该文件或文件夹。",
        "suggestion": (
            "解决方法：\n"
            "1. 右键点击程序 → 「以管理员身份运行」\n"
            "2. 检查文件是否被其他程序占用\n"
            "3. 重启电脑后再试一次"
        )
    },
    "IsADirectoryError": {
        "title": "路径错误",
        "message": "选择的路径是一个文件夹，不是文件。",
        "suggestion": "请重新选择正确的文件，而不是文件夹。"
    },
    "NotADirectoryError": {
        "title": "路径错误",
        "message": "选择的路径不是一个文件夹。",
        "suggestion": "请重新选择正确的文件夹路径。"
    },

    # ---------- 系统/调用相关 ----------
    "OSError": {
        "title": "系统错误",
        "message": "操作系统返回了一个错误。",
        "suggestion": (
            "解决方法：\n"
            "1. 检查文件是否被其他程序占用\n"
            "2. 重启电脑后再试\n"
            "3. 如果问题持续，请重新安装本程序"
        )
    },
    "subprocess.CalledProcessError": {
        "title": "程序调用失败",
        "message": "运行外部程序时出错。",
        "suggestion": (
            "解决方法：\n"
            "1. 检查外部程序的路径是否正确\n"
            "2. 确认该程序没有被杀毒软件拦截\n"
            "3. 尝试手动打开该程序看是否正常"
        )
    },
    "TimeoutError": {
        "title": "操作超时",
        "message": "程序响应时间过长。",
        "suggestion": (
            "解决方法：\n"
            "1. 检查电脑是否运行缓慢\n"
            "2. 关闭一些不需要的程序再试\n"
            "3. 如果搜索全盘文件，请耐心等待"
        )
    },

    # ---------- 配置/JSON 相关 ----------
    "json.JSONDecodeError": {
        "title": "配置文件损坏",
        "message": "程序的配置文件格式错误，可能已损坏。",
        "suggestion": (
            "解决方法：\n"
            "1. 关闭本程序\n"
            "2. 找到程序所在文件夹中的 config.json 文件\n"
            "3. 删除它（程序会自动重新创建）\n"
            "4. 重新打开程序"
        )
    },
    "json.JSONEncoderError": {
        "title": "配置保存失败",
        "message": "保存设置时出现问题。",
        "suggestion": (
            "解决方法：\n"
            "1. 检查程序所在文件夹是否被设置为「只读」\n"
            "2. 右键程序 → 「以管理员身份运行」\n"
            "3. 如果问题持续，重新安装本程序"
        )
    },

    # ---------- PyQt5 相关 ----------
    "ImportError": {
        "title": "组件缺失",
        "message": "程序缺少必要的组件。",
        "suggestion": (
            "解决方法：\n"
            "1. 重新安装本程序\n"
            "2. 如果是从源代码运行，请运行：pip install PyQt5\n"
            "3. 如果问题持续，请联系开发者"
        )
    },
    "ModuleNotFoundError": {
        "title": "模块缺失",
        "message": "程序缺少必要的模块。",
        "suggestion": (
            "解决方法：\n"
            "1. 重新安装本程序\n"
            "2. 如果是从源代码运行，请安装缺少的模块\n"
            "3. 如果问题持续，请联系开发者"
        )
    },

    # ---------- 通用 ----------
    "MemoryError": {
        "title": "内存不足",
        "message": "电脑可用内存不足，程序无法继续运行。",
        "suggestion": (
            "解决方法：\n"
            "1. 关闭一些不需要的程序\n"
            "2. 重启电脑释放内存\n"
            "3. 如果经常出现，请考虑增加电脑内存"
        )
    },
    "KeyboardInterrupt": {
        "title": "操作已取消",
        "message": "用户取消了当前操作。",
        "suggestion": "如果您不是有意取消，请重试。"
    },
    "ConnectionError": {
        "title": "网络连接失败",
        "message": "程序无法连接到网络。",
        "suggestion": (
            "解决方法：\n"
            "1. 检查网络是否正常连接\n"
            "2. 检查防火墙是否阻止了本程序\n"
            "3. 联系网络管理员"
        )
    },
}


def get_suggestion_for_exception(exc: Exception):
    """根据异常类型返回友好的错误信息和解决建议"""
    exc_type = type(exc).__name__

    # 精确匹配
    if exc_type in ERROR_SUGGESTIONS:
        info = ERROR_SUGGESTIONS[exc_type]
        return info["title"], info["message"], info["suggestion"]

    # 模糊匹配：检查异常类型名称是否包含关键词
    for key, info in ERROR_SUGGESTIONS.items():
        if key.lower() in exc_type.lower():
            return info["title"], info["message"], info["suggestion"]

    # 通用兜底
    return (
        "程序遇到意外错误",
        f"发生了未预期的错误：{str(exc)}",
        (
            "建议尝试以下方法：\n"
            "1. 重启本程序\n"
            "2. 重启电脑\n"
            "3. 如果问题持续，请联系开发者并提供错误信息"
        )
    )


def show_error(parent, exc: Exception, extra_context: str = ""):
    """
    显示友好的错误弹窗，包含错误原因和解决建议。

    参数：
        parent: 父窗口（QWidget）
        exc: 捕获到的异常对象
        extra_context: 额外的上下文说明（可选）
    """
    title, message, suggestion = get_suggestion_for_exception(exc)

    # 构建完整错误信息
    full_message = f"❌ {message}\n"
    if extra_context:
        full_message += f"\n📌 位置：{extra_context}\n"
    full_message += f"\n📋 错误详情：{type(exc).__name__}: {str(exc)}"
    full_message += f"\n\n💡 {suggestion}"

    QMessageBox.critical(parent, f"错误 - {title}", full_message)


def show_warning(parent, title: str, message: str, suggestion: str = ""):
    """显示警告弹窗（带建议）"""
    full = f"{message}"
    if suggestion:
        full += f"\n\n💡 {suggestion}"
    QMessageBox.warning(parent, title, full)


def show_info(parent, title: str, message: str):
    """显示信息弹窗"""
    QMessageBox.information(parent, title, message)
