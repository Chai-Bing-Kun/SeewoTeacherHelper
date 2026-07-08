import sys
import os
import json
import time
import subprocess
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QListWidgetItem, QFileDialog,
    QMessageBox, QLineEdit, QCheckBox, QGroupBox,
    QTextEdit, QSystemTrayIcon, QMenu, QAction,
    QDialog, QFormLayout, QDialogButtonBox,
    QTabWidget
)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QThread
from PyQt5.QtGui import QFont, QIcon, QPixmap, QPainter, QColor, QPen, QBrush
from PyQt5.QtWidgets import QStyle

from error_helper import show_error, show_warning, show_info


# ============================================================
# 开机自启动管理
# ============================================================
def get_startup_shortcut_path():
    """获取开机自启动快捷方式路径"""
    startup_folder = os.path.join(os.environ.get("APPDATA", ""),
                                  r"Microsoft\Windows\Start Menu\Programs\Startup")
    return os.path.join(startup_folder, "SeewoTeacherHelper.lnk")


def set_auto_startup(enable: bool):
    """设置或取消开机自启动"""
    shortcut_path = get_startup_shortcut_path()
    if enable:
        try:
            if getattr(sys, 'frozen', False):
                # 打包为 exe 时，直接使用 exe 路径
                target = sys.executable
                working_dir = os.path.dirname(sys.executable)
                args = ""
            else:
                # 源码运行时，使用 python.exe + main.py 路径
                target = sys.executable
                script_path = os.path.abspath(__file__)
                working_dir = os.path.dirname(script_path)
                args = script_path
            # 使用 PowerShell 创建快捷方式（通过标准输入传递命令，避免引号转义问题）
            ps_script = (
                f'$ws = New-Object -ComObject WScript.Shell\n'
                f'$s = $ws.CreateShortcut("{shortcut_path}")\n'
                f'$s.TargetPath = "{target}"\n'
                f'$s.Arguments = "{args}"\n'
                f'$s.WorkingDirectory = "{working_dir}"\n'
                f'$s.Description = "Seewo Teacher Helper"\n'
                f'$s.Save()\n'
            )
            subprocess.run(
                ["powershell", "-NoProfile", "-Command", "-"],
                input=ps_script, capture_output=True, text=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            return True, "已设置开机自启动"
        except Exception as e:
            return False, str(e)
    else:
        try:
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
            return True, "已取消开机自启动"
        except Exception as e:
            return False, str(e)


def is_auto_startup_enabled() -> bool:
    """检查是否已设置开机自启动"""
    return os.path.exists(get_startup_shortcut_path())


# ============================================================
# PPT 进程监控线程
# ============================================================
class PptMonitorThread(QThread):
    """监控 PowerPoint/WPS 进程数量变化"""
    ppt_opened = pyqtSignal()   # 检测到新的PPT打开时触发
    ppt_closed = pyqtSignal()   # PPT全部关闭时触发

    def __init__(self):
        super().__init__()
        self._stopped = False
        self._last_count = 0

    def stop(self):
        self._stopped = True

    def _count_ppt_processes(self, stdout_lower):
        """统计 PPT 相关进程的数量"""
        ppt_processes = ['powerpnt.exe', 'wpp.exe', 'et.exe', 'wps.exe']
        total = 0
        for proc in ppt_processes:
            # CSV 中每行一个进程: "powerpnt.exe","1234","Console","1","xxx"
            # 统计该进程名出现的行数
            proc_quoted = f'"{proc}"'
            count = stdout_lower.count(proc_quoted)
            if count > 0:
                total += count
            else:
                # 兜底：直接统计包含进程名的行
                for line in stdout_lower.splitlines():
                    if proc in line:
                        total += 1
        return total

    def run(self):
        while not self._stopped:
            try:
                result = subprocess.run(
                    ["tasklist", "/NH", "/FO", "CSV"],
                    capture_output=True, text=True, timeout=5,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                stdout_lower = result.stdout.lower()

                current_count = self._count_ppt_processes(stdout_lower)

                if current_count > self._last_count:
                    # PPT 进程数量增加了 → 有新的 PPT 被打开
                    self._last_count = current_count
                    # 延迟一小段时间再触发，确保 PPT 进程已完全就绪
                    self.msleep(800)
                    self.ppt_opened.emit()
                elif current_count < self._last_count:
                    self._last_count = current_count
                    if current_count == 0:
                        self.ppt_closed.emit()
            except Exception:
                pass
            self.msleep(2000)  # 每2秒检测一次


# ============================================================
# PPT 检测到后的外置程序检测/重启线程
# ============================================================
class PptCheckThread(QThread):
    """在 PPT 打开后，在工作线程中执行外置程序检测和重启"""
    log = pyqtSignal(str)  # 日志消息

    def __init__(self, apps, monitor_ppt_enabled=True, check_before_ppt_enabled=True):
        super().__init__()
        self.apps = apps
        # 以设置中的总开关为准，覆盖单个程序的设置
        self.monitor_ppt_enabled = monitor_ppt_enabled
        self.check_before_ppt_enabled = check_before_ppt_enabled

    def run(self):
        # 先等待一下，确保 PPT 进程完全就绪
        self.msleep(500)

        for app in self.apps:
            # 先检查设置中的总开关
            # 如果 monitor_ppt 关闭，则所有检测/重启都不执行
            if not self.monitor_ppt_enabled:
                continue
            # 如果 check_before_ppt 关闭，则不执行自动检测启动
            if not self.check_before_ppt_enabled and not app.restart_on_ppt:
                continue

            # 单个程序的 auto_check 仍然生效（在总开关开启的前提下）
            if not app.auto_check:
                continue

            try:
                if app.restart_on_ppt:
                    # 标记了"打开PPT时重启"的程序：每次都执行重启
                    self.log.emit(f"  [监控] 正在重启：{app.name}")
                    ok, msg = app.restart()
                    if ok:
                        self.log.emit(f"  [监控] 已重启：{app.name}")
                    else:
                        self.log.emit(f"  [监控] 重启失败：{app.name} - {msg}")
                else:
                    # 未标记重启的程序：检测是否在运行，不在则启动
                    running = app.is_running()
                    if not running:
                        self.log.emit(f"  [监控] {app.name} 未运行，正在启动...")
                        ok, msg = app.launch()
                        if ok:
                            self.log.emit(f"  [监控] 已启动：{app.name}")
                        else:
                            self.log.emit(f"  [监控] 启动失败：{app.name} - {msg}")
                    else:
                        self.log.emit(f"  [监控] {app.name} 运行正常")
            except Exception as e:
                self.log.emit(f"  [监控] {app.name} 检测出错：{str(e)}")


# ============================================================
# 配置管理
# ============================================================
def get_config_path():
    """获取配置文件路径（兼容 exe 打包场景）"""
    if getattr(sys, 'frozen', False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).parent
    return base / "config.json"


CONFIG_FILE = get_config_path()

DEFAULT_CONFIG = {
    "check_before_ppt": True,
    "auto_startup": False,          # 开机自启动
    "start_mode": "tray",           # 启动模式: "tray"=折叠框, "window"=图形界面
    "monitor_ppt": True,            # 检测PPT程序运行（自动触发外置程序检测和重启）
    "window_width": 800,
    "window_height": 600,
    "external_apps": []
}


def load_config():
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except json.JSONDecodeError:
        # 配置文件损坏，自动重置
        show_warning(None, "配置文件损坏",
            f"配置文件 {CONFIG_FILE} 格式错误，已自动重置为默认设置。",
            "如果此问题反复出现，请删除该文件后重试。")
    except Exception:
        pass
    return DEFAULT_CONFIG.copy()


def save_config(config):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except PermissionError:
        show_warning(None, "保存失败",
            f"无法写入配置文件 {CONFIG_FILE}，权限不足。",
            "请以管理员身份运行本程序，或检查文件是否被设置为「只读」。")
    except Exception as e:
        show_warning(None, "保存失败",
            f"保存配置时出错：{str(e)}",
            "请检查程序所在文件夹是否可写。")


# ============================================================
# 外置应用数据模型
# ============================================================
class ExternalApp:
    def __init__(self, name="", exe_path="", auto_check=True, args="", restart_on_ppt=False):
        self.name = name
        self.exe_path = exe_path
        self.auto_check = auto_check
        self.args = args
        self.restart_on_ppt = restart_on_ppt  # 打开PPT时自动重启

    def to_dict(self):
        return {
            "name": self.name,
            "exe_path": self.exe_path,
            "auto_check": self.auto_check,
            "args": self.args,
            "restart_on_ppt": self.restart_on_ppt
        }

    @staticmethod
    def from_dict(d):
        return ExternalApp(
            name=d.get("name", ""),
            exe_path=d.get("exe_path", ""),
            auto_check=d.get("auto_check", True),
            args=d.get("args", ""),
            restart_on_ppt=d.get("restart_on_ppt", False)
        )

    def is_running(self):
        """检测程序是否正在运行"""
        if not self.exe_path:
            return False
        try:
            exe_name = os.path.basename(self.exe_path).lower()
            if not exe_name:
                return False
            result = subprocess.run(
                ["tasklist", "/FI", f"IMAGENAME eq {exe_name}", "/NH"],
                capture_output=True, text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                timeout=5
            )
            return exe_name in result.stdout.lower()
        except Exception:
            return False

    def kill(self):
        """强制结束进程"""
        if not self.exe_path:
            return False, "路径为空"
        try:
            exe_name = os.path.basename(self.exe_path).lower()
            subprocess.run(
                ["taskkill", "/F", "/IM", exe_name],
                capture_output=True, text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                timeout=5
            )
            return True, "已终止"
        except subprocess.TimeoutExpired:
            return False, "操作超时，进程可能未响应"
        except PermissionError:
            return False, "权限不足，请以管理员身份运行"
        except Exception as e:
            return False, str(e)

    def restart(self):
        """重启程序：先杀进程，再启动"""
        self.kill()
        time.sleep(1.5)
        return self.launch()

    def launch(self):
        """启动程序"""
        if not self.exe_path:
            return False, "路径为空"
        if not os.path.exists(self.exe_path):
            return False, "文件不存在，请检查路径是否正确"
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

            cmd = [self.exe_path]
            if self.args:
                cmd += self.args.split()

            subprocess.Popen(
                cmd,
                creationflags=subprocess.CREATE_NO_WINDOW,
                startupinfo=startupinfo
            )
            return True, "启动成功"
        except FileNotFoundError:
            return False, "找不到程序文件，请检查路径"
        except PermissionError:
            return False, "权限不足，请以管理员身份运行"
        except OSError as e:
            return False, f"系统错误：{str(e)}"
        except Exception as e:
            return False, str(e)


# ============================================================
# 检测线程
# ============================================================
class CheckThread(QThread):
    finished = pyqtSignal(list)

    def __init__(self, apps):
        super().__init__()
        self.apps = apps

    def run(self):
        results = []
        for app in self.apps:
            try:
                running = app.is_running() if app.auto_check else False
                results.append((app.name, app.exe_path, running))
            except Exception:
                results.append((app.name, app.exe_path, False))
        self.finished.emit(results)


# ============================================================
# PPT 智能搜索线程（U盘优先 > 桌面 > 其他磁盘，实时推送）
# ============================================================
class PptSearchThread(QThread):
    file_found = pyqtSignal(str, float, int)  # path, mtime, priority
    finished = pyqtSignal()
    status = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self._stopped = False

    def stop(self):
        """请求线程停止搜索"""
        self._stopped = True

    def run(self):
        self.status.emit("正在全盘搜索PPT文件...")

        # 1. 确定搜索范围：U盘 > 桌面 > 其他磁盘
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        usb_drives = []
        other_drives = []

        for letter in "DEFGHIJKLMNOPQRSTUVWXYZ":
            drive = f"{letter}:\\"
            try:
                if os.path.exists(drive):
                    try:
                        result = subprocess.run(
                            ["fsutil", "fsinfo", "drivetype", drive],
                            capture_output=True, text=True,
                            creationflags=subprocess.CREATE_NO_WINDOW,
                            timeout=3
                        )
                        if "Removable" in result.stdout or "可移动" in result.stdout:
                            usb_drives.append(drive)
                        else:
                            other_drives.append(drive)
                    except Exception:
                        other_drives.append(drive)
            except Exception:
                continue

        # 构建搜索队列：U盘(优先级0) > 桌面(优先级1) > 其他磁盘(优先级2)
        search_queue = []
        for d in usb_drives:
            search_queue.append((d, 0))
        search_queue.append((desktop, 1))
        for d in other_drives:
            if d not in usb_drives:
                search_queue.append((d, 2))

        # 2. 执行搜索 — 每找到一个立即通过信号发送
        found = 0
        max_results = 200

        for search_path, priority in search_queue:
            if found >= max_results or self._stopped:
                break
            if not os.path.exists(search_path):
                continue

            try:
                for root, dirs, files in os.walk(search_path):
                    if self._stopped:
                        break
                    dirs[:] = [d for d in dirs
                               if not d.startswith('$')
                               and not d.startswith('.')
                               and d not in ('System Volume Information', 'Recovery', 'Windows')]
                    rel = root[len(search_path):].strip(os.sep)
                    depth = rel.count(os.sep) + 1 if rel else 0
                    if depth > 6:
                        dirs.clear()
                        continue

                    if self._stopped:
                        break
                    for f in files:
                        if self._stopped:
                            break
                        try:
                            if f.lower().endswith(('.ppt', '.pptx', '.pps', '.ppsx')):
                                full_path = os.path.join(root, f)
                                try:
                                    mtime = os.path.getmtime(full_path)
                                except Exception:
                                    mtime = 0
                                # 实时发送找到的文件
                                self.file_found.emit(full_path, mtime, priority)
                                found += 1
                                if found >= max_results:
                                    break
                        except Exception:
                            continue
                    if found >= max_results or self._stopped:
                        break
            except Exception:
                continue

        self.finished.emit()


# ============================================================
# 编辑外置应用对话框
# ============================================================
class AppEditDialog(QDialog):
    def __init__(self, app=None, parent=None,
                 monitor_ppt_enabled=True, check_before_ppt_enabled=True):
        super().__init__(parent)
        self.app = app
        self.monitor_ppt_enabled = monitor_ppt_enabled
        self.check_before_ppt_enabled = check_before_ppt_enabled
        self.setWindowTitle("编辑外置应用" if app else "添加外置应用")
        self.setMinimumWidth(450)
        self.setup_ui()
        if app:
            self.load_app(app)

    def setup_ui(self):
        layout = QFormLayout(self)

        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("例如：翻页笔驱动、音响控制台")
        layout.addRow("应用名称：", self.name_edit)

        path_layout = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("选择或输入可执行文件路径")
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(self.browse_file)
        path_layout.addWidget(self.path_edit)
        path_layout.addWidget(browse_btn)
        layout.addRow("程序路径：", path_layout)

        self.args_edit = QLineEdit()
        self.args_edit.setPlaceholderText("例如：--silent 或 /port=8080")
        layout.addRow("启动参数：", self.args_edit)

        args_hint = QLabel(
            "💡 如果程序启动时需要附加额外指令才正常工作，请填写在这里。\n"
            "   大多数程序不需要填写此项，留空即可。"
        )
        args_hint.setStyleSheet("color: #888; font-size: 11px; padding-left: 5px;")
        args_hint.setWordWrap(True)
        layout.addRow("", args_hint)

        self.auto_check_cb = QCheckBox("自动检测运行状态")
        self.auto_check_cb.setChecked(True)
        if not self.check_before_ppt_enabled:
            self.auto_check_cb.setEnabled(False)
            self.auto_check_cb.setToolTip("设置中的「打开PPT前自动检测」总开关已关闭，该功能暂不可用")
        layout.addRow(self.auto_check_cb)

        self.restart_cb = QCheckBox("打开PPT时自动重启（先杀进程再重新打开）")
        self.restart_cb.setChecked(False)
        if not self.monitor_ppt_enabled:
            self.restart_cb.setEnabled(False)
            self.restart_cb.setToolTip("设置中的「检测到PPT时自动执行」总开关已关闭，该功能暂不可用")
        layout.addRow(self.restart_cb)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择可执行文件", "",
            "可执行文件 (*.exe *.bat *.cmd *.com);;所有文件 (*.*)"
        )
        if path:
            self.path_edit.setText(path)

    def load_app(self, app):
        self.name_edit.setText(app.name)
        self.path_edit.setText(app.exe_path)
        self.args_edit.setText(app.args)
        self.auto_check_cb.setChecked(app.auto_check)
        self.restart_cb.setChecked(app.restart_on_ppt)

    def validate_and_accept(self):
        name = self.name_edit.text().strip()
        path = self.path_edit.text().strip()
        if not name:
            show_warning(self, "提示", "请输入应用名称")
            return
        if not path:
            show_warning(self, "提示", "请选择程序路径")
            return
        self.accept()

    def get_app(self):
        return ExternalApp(
            name=self.name_edit.text().strip(),
            exe_path=self.path_edit.text().strip(),
            auto_check=self.auto_check_cb.isChecked(),
            args=self.args_edit.text().strip(),
            restart_on_ppt=self.restart_cb.isChecked()
        )


# ============================================================
# 主窗口
# ============================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = load_config()
        self.external_apps = []
        self.load_external_apps()
        self.ppt_files = []
        self._search_results = []
        self._closing = False

        self.setWindowTitle("Seewo Teacher Helper")
        self.setMinimumSize(600, 500)
        self.resize(
            self.config.get("window_width", 800),
            self.config.get("window_height", 600)
        )

        self.setup_ui()
        self.setup_tray()
        self.apply_style()

        # 注意：外置程序的检测和重置仅在检测到PPT打开时触发
        # 不再在启动时检测，避免程序被无故启动

        # 处理启动模式
        start_mode = self.config.get("start_mode", "tray")
        if start_mode == "tray":
            # 默认缩小到折叠框（系统托盘）
            QTimer.singleShot(100, self.hide_to_tray)

        # 启动PPT进程监控
        if self.config.get("monitor_ppt", True):
            self.start_ppt_monitor()

    # ==================== UI 构建 ====================
    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # 标题
        title_label = QLabel("Seewo Teacher Helper")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        desc_label = QLabel("触屏教学辅助工具 - 一键打开PPT / 外置程序管理")
        desc_label.setAlignment(Qt.AlignCenter)
        desc_label.setStyleSheet("color: #888; margin-bottom: 10px;")
        main_layout.addWidget(desc_label)

        # 选项卡
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self.ppt_tab = QWidget()
        self.setup_ppt_tab()
        self.tabs.addTab(self.ppt_tab, "打开PPT")

        self.apps_tab = QWidget()
        self.setup_apps_tab()
        self.tabs.addTab(self.apps_tab, "外置程序")

        self.settings_tab = QWidget()
        self.setup_settings_tab()
        self.tabs.addTab(self.settings_tab, "设置")

        self.log_tab = QWidget()
        self.setup_log_tab()
        self.tabs.addTab(self.log_tab, "日志")

    # ==================== PPT 选项卡 ====================
    def setup_ppt_tab(self):
        layout = QVBoxLayout(self.ppt_tab)
        layout.setSpacing(12)

        self.ppt_status_label = QLabel("点击下方按钮，自动搜索全盘PPT文件（U盘优先）")
        self.ppt_status_label.setStyleSheet("font-size: 13px; padding: 5px;")
        layout.addWidget(self.ppt_status_label)

        # 一键打开PPT按钮
        self.open_ppt_btn = QPushButton("一键打开PPT")
        btn_font = QFont()
        btn_font.setPointSize(16)
        btn_font.setBold(True)
        self.open_ppt_btn.setFont(btn_font)
        self.open_ppt_btn.setMinimumHeight(80)
        self.open_ppt_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 12px;
                padding: 20px;
                font-size: 18px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
            QPushButton:pressed {
                background-color: #a93226;
            }
        """)
        self.open_ppt_btn.clicked.connect(self.on_open_ppt_clicked)
        layout.addWidget(self.open_ppt_btn)

        # PPT文件列表
        self.ppt_list_widget = QListWidget()
        self.ppt_list_widget.setMinimumHeight(150)
        self.ppt_list_widget.setAlternatingRowColors(True)
        self.ppt_list_widget.itemDoubleClicked.connect(self.open_selected_ppt)
        self.ppt_list_widget.itemClicked.connect(self.on_ppt_item_clicked)
        self.ppt_list_widget.currentItemChanged.connect(self.on_ppt_selection_changed)
        layout.addWidget(QLabel("搜索到的PPT文件（按时间排序，U盘优先，双击打开）："))
        layout.addWidget(self.ppt_list_widget)

        # 按钮行
        btn_layout = QHBoxLayout()
        refresh_btn = QPushButton("搜索全盘PPT")
        refresh_btn.clicked.connect(self.on_refresh_clicked)
        browse_btn = QPushButton("手动选择文件")
        browse_btn.clicked.connect(self.on_browse_clicked)
        btn_layout.addWidget(refresh_btn)
        btn_layout.addWidget(browse_btn)

        self.stop_search_btn = QPushButton("停止搜索")
        self.stop_search_btn.setStyleSheet("""
            QPushButton {
                background-color: #e67e22;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #d35400;
            }
        """)
        self.stop_search_btn.clicked.connect(self.stop_search)
        self.stop_search_btn.setVisible(False)
        btn_layout.addWidget(self.stop_search_btn)

        layout.addLayout(btn_layout)

        layout.addStretch()

    def stop_search(self):
        """停止正在进行的全盘搜索"""
        if hasattr(self, 'search_thread') and self.search_thread.isRunning():
            self.search_thread.stop()
            self.search_thread.wait(2000)
            self.open_ppt_btn.setEnabled(True)
            self.open_ppt_btn.setText("一键打开PPT")
            self.ppt_status_label.setText("已停止搜索")
        self.stop_search_btn.setVisible(False)

    def on_open_ppt_clicked(self):
        """一键打开PPT按钮点击处理"""
        # 如果列表中有选中的项，直接打开选中的文件
        current_item = self.ppt_list_widget.currentItem()
        if current_item and current_item.data(Qt.UserRole):
            self.stop_search()
            self.open_ppt_file(current_item.data(Qt.UserRole))
            return

        # 否则执行全盘搜索
        self.open_ppt_btn.setEnabled(False)
        self.open_ppt_btn.setText("正在搜索...")
        self.ppt_list_widget.clear()
        self.ppt_files.clear()

        # 检测外置程序运行状态
        if self.config.get("check_before_ppt", True):
            self.check_external_apps()

        # 启动搜索线程（实时插入）
        self.start_search()

    def on_refresh_clicked(self):
        """搜索按钮点击处理"""
        self.open_ppt_btn.setEnabled(False)
        self.open_ppt_btn.setText("正在搜索...")
        self.ppt_list_widget.clear()
        self.ppt_files.clear()
        self.ppt_status_label.setText("正在全盘搜索PPT文件（U盘优先）...")

        self.start_search()

    def start_search(self):
        """启动搜索线程，实时插入结果"""
        self._search_results = []  # 暂存所有结果，用于最终计数

        self.search_thread = PptSearchThread()
        self.search_thread.status.connect(self.ppt_status_label.setText)
        self.search_thread.file_found.connect(self.on_file_found)
        self.search_thread.finished.connect(self.on_search_finished)
        self.search_thread.start()

        # 显示停止搜索按钮
        self.stop_search_btn.setVisible(True)

    def on_file_found(self, path, mtime, priority):
        """每找到一个文件立即插入列表（保持排序）"""
        self._search_results.append((path, mtime, priority))

        try:
            fname = os.path.basename(path)
            time_str = time.strftime("%Y-%m-%d %H:%M", time.localtime(mtime)) if mtime else ""
            display = f"{fname}  [{time_str}]  -  {path}" if time_str else f"{fname}  -  {path}"

            item = QListWidgetItem(display)
            item.setData(Qt.UserRole, path)

            # 按优先级(小在前) + 时间(新在前) 找到插入位置
            insert_idx = 0
            for i in range(self.ppt_list_widget.count()):
                existing_path = self.ppt_list_widget.item(i).data(Qt.UserRole)
                # 在 _search_results 中找到对应项的元组
                existing_info = None
                for p, m, pr in self._search_results:
                    if p == existing_path:
                        existing_info = (p, m, pr)
                        break
                if existing_info:
                    # 比较：priority 小的排前面；priority 相同时 mtime 大的（更新的）排前面
                    if priority < existing_info[2] or (priority == existing_info[2] and mtime > existing_info[1]):
                        insert_idx = i
                        break
                insert_idx = i + 1

            self.ppt_list_widget.insertItem(insert_idx, item)
            self.ppt_status_label.setText(f"已找到 {len(self._search_results)} 个PPT文件...")
        except Exception:
            pass

    def on_search_finished(self):
        """搜索完成回调"""
        self.ppt_files = [item[0] for item in self._search_results]
        self.open_ppt_btn.setEnabled(True)
        self.open_ppt_btn.setText("一键打开PPT")
        self.stop_search_btn.setVisible(False)

        if not self._search_results:
            self.ppt_status_label.setText("未找到PPT文件，请手动选择")
            self.ppt_list_widget.addItem("（未找到任何PPT文件）")
        else:
            self.ppt_status_label.setText(
                f"搜索完成，共找到 {len(self._search_results)} 个PPT文件（单击选中，再点红色按钮或双击打开）")

    def on_ppt_item_clicked(self, item):
        """单击列表项时更新按钮提示"""
        path = item.data(Qt.UserRole)
        if path:
            self.open_ppt_btn.setText(f"打开：{os.path.basename(path)}")

    def on_ppt_selection_changed(self, current, previous):
        """选中项变化时更新按钮文字"""
        if current is None:
            self.open_ppt_btn.setText("一键打开PPT")

    def on_browse_clicked(self):
        """浏览其他位置"""
        path, _ = QFileDialog.getOpenFileName(
            self, "选择PPT文件", "",
            "PPT文件 (*.ppt *.pptx *.pps *.ppsx);;所有文件 (*.*)"
        )
        if path:
            self.open_ppt_file(path)

    def open_selected_ppt(self, item):
        """双击列表项打开PPT"""
        path = item.data(Qt.UserRole)
        if path:
            self.stop_search()
            self.open_ppt_file(path)

    def open_ppt_file(self, path):
        """使用系统默认程序打开PPT文件"""
        if not path or not os.path.exists(path):
            show_warning(self, "文件未找到",
                f"文件不存在：\n{path}",
                "请检查文件是否被移动或删除，然后重新搜索。")
            return

        try:
            os.startfile(path)
            self.ppt_status_label.setText(f"正在打开：{os.path.basename(path)}")
            # 外置程序的检测和重启由 PPT 监控线程统一处理（on_ppt_detected）
        except Exception as e:
            show_error(self, e, f"打开文件：{path}")

    # ==================== 外置程序选项卡 ====================
    def setup_apps_tab(self):
        layout = QVBoxLayout(self.apps_tab)
        layout.setSpacing(10)

        info_label = QLabel(
            "管理需要在打开PPT前或开机时自动检测的外置程序\n"
            "（如翻页笔驱动、音响控制台等）"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666; padding: 5px;")
        layout.addWidget(info_label)

        self.app_list_widget = QListWidget()
        self.app_list_widget.setAlternatingRowColors(True)
        layout.addWidget(self.app_list_widget)

        btn_layout = QHBoxLayout()
        add_btn = QPushButton("添加程序")
        add_btn.clicked.connect(self.add_external_app)
        edit_btn = QPushButton("编辑")
        edit_btn.clicked.connect(self.edit_external_app)
        remove_btn = QPushButton("删除")
        remove_btn.clicked.connect(self.remove_external_app)
        check_btn = QPushButton("检测运行状态")
        check_btn.clicked.connect(self.check_external_apps)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(remove_btn)
        btn_layout.addWidget(check_btn)
        layout.addLayout(btn_layout)



        self.refresh_app_list()

    def load_external_apps(self):
        apps_data = self.config.get("external_apps", [])
        self.external_apps = [ExternalApp.from_dict(d) for d in apps_data]

    def save_external_apps(self):
        self.config["external_apps"] = [app.to_dict() for app in self.external_apps]
        save_config(self.config)

    def refresh_app_list(self):
        self.app_list_widget.clear()
        for i, app in enumerate(self.external_apps):
            restart_tag = " [打开PPT时重启]" if app.restart_on_ppt else ""
            text = f"  {app.name}  -  {os.path.basename(app.exe_path)}{restart_tag}"
            if app.args:
                text += f"  [参数: {app.args}]"
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, i)
            self.app_list_widget.addItem(item)

    def add_external_app(self):
        dialog = AppEditDialog(
            parent=self,
            monitor_ppt_enabled=self.config.get("monitor_ppt", True),
            check_before_ppt_enabled=self.config.get("check_before_ppt", True)
        )
        if dialog.exec_() == QDialog.Accepted:
            app = dialog.get_app()
            self.external_apps.append(app)
            self.save_external_apps()
            self.refresh_app_list()

    def edit_external_app(self):
        current = self.app_list_widget.currentRow()
        if current < 0 or current >= len(self.external_apps):
            show_info(self, "提示", "请先选择一个程序")
            return
        dialog = AppEditDialog(
            app=self.external_apps[current], parent=self,
            monitor_ppt_enabled=self.config.get("monitor_ppt", True),
            check_before_ppt_enabled=self.config.get("check_before_ppt", True)
        )
        if dialog.exec_() == QDialog.Accepted:
            self.external_apps[current] = dialog.get_app()
            self.save_external_apps()
            self.refresh_app_list()

    def remove_external_app(self):
        current = self.app_list_widget.currentRow()
        if current < 0 or current >= len(self.external_apps):
            show_info(self, "提示", "请先选择一个程序")
            return
        app = self.external_apps[current]
        reply = QMessageBox.question(
            self, "确认删除", f"确定要删除「{app.name}」吗？",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.external_apps.pop(current)
            self.save_external_apps()
            self.refresh_app_list()

    def check_external_apps(self):
        """检测所有外置程序的运行状态"""
        self.log_text.clear()
        self.log_text.append("正在检测外置程序运行状态...\n")

        apps_to_check = [app for app in self.external_apps if app.auto_check]
        if not apps_to_check:
            self.log_text.append("没有需要检测的外置程序")
            return

        self.check_thread = CheckThread(apps_to_check)
        self.check_thread.finished.connect(self.on_check_finished)
        self.check_thread.start()

    def on_check_finished(self, results):
        self.log_text.clear()
        not_running = []
        for name, exe_path, running in results:
            if running:
                self.log_text.append(f"[正常]  {name}")
            else:
                self.log_text.append(f"[未运行]  {name}")
                not_running.append((name, exe_path))

        if not_running:
            self.log_text.append(f"\n有 {len(not_running)} 个程序未运行")
            # 如果窗口隐藏，先显示窗口再弹窗
            if not self.isVisible():
                self.show_and_raise()
            reply = QMessageBox.question(
                self, "检测完成",
                f"有 {len(not_running)} 个程序未运行，是否启动它们？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.launch_apps(not_running)
        else:
            self.log_text.append("\n所有程序运行正常！")

    def launch_apps(self, apps):
        success_count = 0
        for name, exe_path in apps:
            app = ExternalApp(name=name, exe_path=exe_path)
            ok, msg = app.launch()
            if ok:
                success_count += 1
                self.log_text.append(f"  已启动：{name}")
            else:
                self.log_text.append(f"  启动失败：{name} - {msg}")

        if success_count:
            # 窗口隐藏时不弹窗，日志已写入日志选项卡
            if self.isVisible():
                show_info(self, "启动完成", f"成功启动 {success_count} 个程序")

    # ==================== 设置选项卡 ====================
    def setup_settings_tab(self):
        layout = QVBoxLayout(self.settings_tab)
        layout.setSpacing(15)

        # ----- 启动设置 -----
        group_startup = QGroupBox("启动设置")
        group_startup_layout = QVBoxLayout(group_startup)

        self.auto_startup_cb = QCheckBox("开机自动启动")
        self.auto_startup_cb.setChecked(self.config.get("auto_startup", False))
        self.auto_startup_cb.toggled.connect(self.on_auto_startup_toggled)
        group_startup_layout.addWidget(self.auto_startup_cb)

        self.start_mode_tray_cb = QCheckBox("启动时缩小到系统托盘（折叠框）")
        self.start_mode_tray_cb.setChecked(self.config.get("start_mode", "tray") == "tray")
        self.start_mode_tray_cb.toggled.connect(self.on_start_mode_toggled)
        group_startup_layout.addWidget(self.start_mode_tray_cb)

        layout.addWidget(group_startup)

        # ----- 检测设置 -----
        group1 = QGroupBox("检测设置")
        group1_layout = QVBoxLayout(group1)

        self.monitor_ppt_cb = QCheckBox("【总开关】检测到PPT时自动执行外置程序检测和重启")
        self.monitor_ppt_cb.setChecked(self.config.get("monitor_ppt", True))
        self.monitor_ppt_cb.toggled.connect(self.on_monitor_ppt_toggled)
        group1_layout.addWidget(self.monitor_ppt_cb)

        self.check_before_ppt_cb = QCheckBox("【总开关】打开PPT前自动检测外置程序运行状态")
        self.check_before_ppt_cb.setChecked(self.config.get("check_before_ppt", True))
        self.check_before_ppt_cb.toggled.connect(
            lambda v: self.update_config("check_before_ppt", v))
        group1_layout.addWidget(self.check_before_ppt_cb)

        monitor_hint = QLabel(
            "💡 总开关关闭时，即使单个程序勾选了对应功能也不会生效。\n"
            "   「检测到PPT时」：监控PPT进程，自动触发外置程序检测和重启\n"
            "   「打开PPT前」：通过本程序打开PPT前检测外置程序运行状态"
        )
        monitor_hint.setStyleSheet("color: #888; font-size: 11px; padding-left: 20px;")
        monitor_hint.setWordWrap(True)
        group1_layout.addWidget(monitor_hint)

        layout.addWidget(group1)

        # ----- 关于 -----
        group3 = QGroupBox("关于")
        group3_layout = QVBoxLayout(group3)
        about_text = QLabel(
            "Seewo Teacher Helper v2.1\n\n"
            "专为触屏教学场景设计，帮助教师快速打开PPT课件，\n"
            "并自动检测外置设备驱动程序的运行状态。\n\n"
            "支持自定义添加外置程序，灵活适配不同教学环境。"
        )
        about_text.setWordWrap(True)
        about_text.setStyleSheet("color: #555;")
        group3_layout.addWidget(about_text)
        layout.addWidget(group3)

        layout.addStretch()

    # ==================== 日志选项卡 ====================
    def setup_log_tab(self):
        layout = QVBoxLayout(self.log_tab)
        layout.setSpacing(10)

        header = QLabel("运行日志")
        header_font = QFont()
        header_font.setPointSize(12)
        header_font.setBold(True)
        header.setFont(header_font)
        layout.addWidget(header)

        hint = QLabel(
            "记录 PPT 进程监控、外置程序检测与重启等所有运行信息，便于调试和排查问题。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #888; font-size: 11px;")
        layout.addWidget(hint)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #ddd;
                border-radius: 8px;
                padding: 8px;
                background-color: #1e1e1e;
                color: #d4d4d4;
                font-family: Consolas, "Courier New", monospace;
                font-size: 12px;
            }
        """)
        layout.addWidget(self.log_text)

        btn_layout = QHBoxLayout()
        clear_btn = QPushButton("清空日志")
        clear_btn.clicked.connect(self.log_text.clear)
        btn_layout.addStretch()
        btn_layout.addWidget(clear_btn)
        layout.addLayout(btn_layout)

    def on_auto_startup_toggled(self, enabled):
        """开机自启动开关"""
        ok, msg = set_auto_startup(enabled)
        if ok:
            self.config["auto_startup"] = enabled
            save_config(self.config)
        else:
            show_warning(self, "设置失败",
                f"无法{'启用' if enabled else '关闭'}开机自启动：{msg}",
                "请尝试以管理员身份运行本程序。")
            # 恢复复选框状态
            self.auto_startup_cb.blockSignals(True)
            self.auto_startup_cb.setChecked(not enabled)
            self.auto_startup_cb.blockSignals(False)

    def on_start_mode_toggled(self, tray_mode):
        """启动模式切换：窗口/折叠框"""
        mode = "tray" if tray_mode else "window"
        self.config["start_mode"] = mode
        save_config(self.config)

    def on_monitor_ppt_toggled(self, enabled):
        """PPT监控开关"""
        self.config["monitor_ppt"] = enabled
        save_config(self.config)
        if enabled:
            self.start_ppt_monitor()
        else:
            if hasattr(self, 'ppt_monitor'):
                self.ppt_monitor.stop()
                self.ppt_monitor.wait(2000)

    def update_config(self, key, value):
        self.config[key] = value
        save_config(self.config)

    # ==================== 系统托盘 ====================
    def setup_tray(self):
        try:
            self.tray_icon = QSystemTrayIcon(self)
            self.tray_icon.setToolTip("Seewo Teacher Helper")

            # 创建托盘图标：使用 QStyle 标准图标或绘制一个简单图标
            icon = self.style().standardIcon(QStyle.SP_ComputerIcon)
            if icon.isNull():
                # 如果标准图标不可用，手动绘制一个简单图标
                pixmap = QPixmap(32, 32)
                pixmap.fill(Qt.transparent)
                painter = QPainter(pixmap)
                painter.setRenderHint(QPainter.Antialiasing)
                painter.setBrush(QBrush(QColor("#4a90d9")))
                painter.setPen(QPen(QColor("#2c5f8a"), 1))
                painter.drawRoundedRect(2, 2, 28, 28, 6, 6)
                painter.setPen(QPen(QColor("white"), 2))
                painter.drawText(pixmap.rect(), Qt.AlignCenter, "S")
                painter.end()
                icon = QIcon(pixmap)
            self.tray_icon.setIcon(icon)

            tray_menu = QMenu()
            show_action = QAction("显示主窗口", self)
            show_action.triggered.connect(self.show_and_raise)
            tray_menu.addAction(show_action)

            open_ppt_action = QAction("一键打开PPT", self)
            open_ppt_action.triggered.connect(self.on_open_ppt_clicked)
            tray_menu.addAction(open_ppt_action)

            check_action = QAction("检测外置程序", self)
            check_action.triggered.connect(self.check_external_apps)
            tray_menu.addAction(check_action)

            tray_menu.addSeparator()

            quit_action = QAction("退出", self)
            quit_action.triggered.connect(self.quit_app)
            tray_menu.addAction(quit_action)

            self.tray_icon.setContextMenu(tray_menu)
            self.tray_icon.activated.connect(self.on_tray_activated)
            self.tray_icon.show()
        except Exception:
            self.tray_icon = None

    def show_and_raise(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_and_raise()

    def hide_to_tray(self):
        """隐藏窗口到系统托盘（折叠框）"""
        self.hide()
        if self.tray_icon:
            try:
                self.tray_icon.showMessage(
                    "Seewo Teacher Helper",
                    "程序已最小化到系统托盘",
                    QSystemTrayIcon.Information,
                    2000
                )
            except Exception:
                pass

    def start_ppt_monitor(self):
        """启动PPT进程监控"""
        self.ppt_monitor = PptMonitorThread()
        self.ppt_monitor.ppt_opened.connect(self.on_ppt_detected)
        self.ppt_monitor.ppt_closed.connect(self.on_ppt_closed)
        self.ppt_monitor.start()

    def on_ppt_detected(self):
        """检测到PPT程序正在运行 → 执行外置程序检测和重启"""
        self.log_text.append("\n[监控] 检测到PPT程序正在运行...")

        # 以设置中的总开关为准，传递给工作线程
        monitor_ppt = self.config.get("monitor_ppt", True)
        check_before_ppt = self.config.get("check_before_ppt", True)

        # 在工作线程中执行检测和重启，避免阻塞主线程
        self._ppt_check_thread = PptCheckThread(
            self.external_apps,
            monitor_ppt_enabled=monitor_ppt,
            check_before_ppt_enabled=check_before_ppt
        )
        self._ppt_check_thread.log.connect(self._on_ppt_check_log)
        self._ppt_check_thread.start()

    def _on_ppt_check_log(self, message):
        """接收 PPT 检测线程的日志消息，更新 UI"""
        self.log_text.append(message)

    def on_ppt_closed(self):
        """PPT程序关闭时的处理（预留）"""
        pass

    def quit_app(self):
        self._closing = True
        if hasattr(self, 'ppt_monitor'):
            self.ppt_monitor.stop()
            self.ppt_monitor.wait(2000)
        if self.tray_icon:
            self.tray_icon.hide()
        QApplication.quit()

    def closeEvent(self, event):
        """关闭窗口时最小化到托盘"""
        if self._closing:
            event.accept()
            return

        try:
            self.config["window_width"] = self.width()
            self.config["window_height"] = self.height()
            save_config(self.config)
        except Exception:
            pass

        event.ignore()
        self.hide()

        if self.tray_icon:
            try:
                self.tray_icon.showMessage(
                    "Seewo Teacher Helper",
                    "程序已最小化到系统托盘",
                    QSystemTrayIcon.Information,
                    2000
                )
            except Exception:
                pass

    # ==================== 样式 ====================
    def apply_style(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f8f9fa;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ddd;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QPushButton {
                padding: 8px 16px;
                border: 1px solid #ccc;
                border-radius: 6px;
                background-color: #fff;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #e8f0fe;
                border-color: #4a90d9;
            }
            QPushButton:pressed {
                background-color: #d0e0f8;
            }
            QTabWidget::pane {
                border: 1px solid #ddd;
                border-radius: 8px;
                padding: 10px;
                background-color: white;
            }
            QTabBar::tab {
                padding: 8px 20px;
                margin-right: 2px;
                border: 1px solid #ddd;
                border-bottom: none;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                background-color: #f0f0f0;
            }
            QTabBar::tab:selected {
                background-color: white;
                border-bottom-color: white;
                font-weight: bold;
            }
            QTabBar::tab:hover {
                background-color: #e8f0fe;
            }
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 8px;
                padding: 5px;
                font-size: 13px;
            }
            QListWidget::item {
                padding: 8px;
                border-bottom: 1px solid #f0f0f0;
            }
            QListWidget::item:hover {
                background-color: #e8f0fe;
            }
            QListWidget::item:selected {
                background-color: #4a90d9;
                color: white;
            }
        """)


# ============================================================
# 程序入口
# ============================================================
def main():
    # 修复 PyQt5 在虚拟环境中找不到 Qt 平台插件的问题
    try:
        import PyQt5
        pyqt5_dir = os.path.dirname(PyQt5.__file__)
        plugins_dir = os.path.join(pyqt5_dir, "Qt5", "plugins")
        if os.path.isdir(plugins_dir):
            os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = os.path.join(plugins_dir, "platforms")
    except ImportError:
        show_warning(None, "组件缺失",
            "未找到 PyQt5 组件，程序可能无法正常显示。",
            "请重新安装本程序，或运行：pip install PyQt5")
    except Exception:
        pass

    try:
        app = QApplication(sys.argv)
        app.setApplicationName("Seewo Teacher Helper")
        app.setOrganizationName("SeewoHelper")

        window = MainWindow()
        window.show()

        exit_code = app.exec_()
        sys.exit(exit_code)
    except Exception as e:
        try:
            from error_helper import show_error
            error_app = QApplication(sys.argv)
            # 在启动失败时，用 Exception 构造一个伪异常来触发通用建议
            show_error(None, e, "程序启动")
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()
