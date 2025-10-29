import sys
import winreg
import os
import subprocess
import logging
import configparser
import ctypes
from ctypes import wintypes

class MSLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("pt", wintypes.POINT),
        ("mouseData", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", wintypes.LPARAM),
    ]

import threading
import time
import psutil
import win32gui
import win32process
import win32con
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import QSharedMemory

# =========================
# Language Support
# =========================
class Translator:
    def __init__(self):
        self.languages = {
            'zh_CN': {
                'app_title': '鼠标滚轮防抖工具',
                'settings': '防抖设置',
                'block_interval': '时间阈值:',
                'direction_threshold': '方向改变阈值:',
                'enable_blocking': '启用滚轮防抖',
                'startup': '开机自动启动',
                'status': '实时状态',
                'total_events': '总滚轮事件:',
                'blocked_events': '已拦截事件:',
                'current_direction': '当前方向:',
                'direction_up': '↑ 向上',
                'direction_down': '↓ 向下',
                'direction_none': '无',
                'status_label': '状态:',
                'how_it_works': '工作原理',
                'how_it_works_text': '''• 时间阈值：距离上次滚轮事件的时间在阈值内，如果出现反方向滚轮事件，则阻塞该事件
• 次数阈值：连续的反方向滚轮事件（每个事件距离上次滚轮时间在时间阈值内）达到次数阈值，则改变初始方向并允许此事件
• 这样可以有效防止因鼠标滚轮抖动导致的意外滚动''',
                'restart_hook': '重启钩子',
                'minimize_to_tray': '最小化到托盘',
                'quit': '退出',
                'tray_show': '显示窗口',
                'tray_quit': '退出',
                'confirm_quit': '确认退出',
                'quit_message': '确定要退出鼠标滚轮防抖工具吗？',
                'tray_message_title': '鼠标滚轮防抖工具',
                'tray_message_content': '程序已最小化到系统托盘，点击托盘图标可以重新打开窗口。',
                'hook_started': '钩子已启动',
                'hook_restarting': '正在重启钩子...',
                'hook_restarted': '钩子已重启',
                'hook_failed': '钩子启动失败',
                'seconds': ' 秒',
                'status_waiting': '等待滚轮事件',
                'status_disabled': '拦截已关闭',
                'status_initial_up': '初始方向: 向上',
                'status_initial_down': '初始方向: 向下',
                'status_same_direction': '同方向滚动',
                'status_direction_changed': '方向改变 (连续{count}次)',
                'status_blocked': '已拦截 (连续{current}/{threshold}次)'
            },
            'en_US': {
                'app_title': 'Mouse Scroll Stabilizer',
                'settings': 'Stabilization Settings',
                'block_interval': 'Time Threshold:',
                'direction_threshold': 'Direction Change Threshold:',
                'enable_blocking': 'Enable Scroll Stabilization',
                'startup': 'Start Automatically at Boot',
                'status': 'Real-time Status',
                'total_events': 'Total Scroll Events:',
                'blocked_events': 'Blocked Events:',
                'current_direction': 'Current Direction:',
                'direction_up': '↑ Up',
                'direction_down': '↓ Down',
                'direction_none': 'None',
                'status_label': 'Status:',
                'how_it_works': 'How It Works',
                'how_it_works_text': '''• Time Threshold: If a reverse scroll event occurs within the threshold time after the last scroll event, it will be blocked
• Count Threshold: When consecutive reverse scroll events (each within the time threshold) reach the count threshold, the initial direction changes and the event is allowed
• This effectively prevents accidental scrolling caused by mouse wheel jitter''',
                'restart_hook': 'Restart Hook',
                'minimize_to_tray': 'Minimize to Tray',
                'quit': 'Quit',
                'tray_show': 'Show Window',
                'tray_quit': 'Quit',
                'confirm_quit': 'Confirm Quit',
                'quit_message': 'Are you sure you want to quit Mouse Scroll Stabilizer?',
                'tray_message_title': 'Mouse Scroll Stabilizer',
                'tray_message_content': 'The program has been minimized to the system tray. Click the tray icon to reopen the window.',
                'hook_started': 'Hook Started',
                'hook_restarting': 'Restarting Hook...',
                'hook_restarted': 'Hook Restarted',
                'hook_failed': 'Hook Failed to Start',
                'seconds': ' seconds',
                'status_waiting': 'Waiting for scroll events',
                'status_disabled': 'Blocking disabled',
                'status_initial_up': 'Initial direction: Up',
                'status_initial_down': 'Initial direction: Down',
                'status_same_direction': 'Same direction scrolling',
                'status_direction_changed': 'Direction changed (consecutive {count} times)',
                'status_blocked': 'Blocked (consecutive {current}/{threshold} times)'
            }
        }
        self.current_lang = 'zh_CN'
    
    def set_language(self, lang):
        if lang in self.languages:
            self.current_lang = lang
    
    def tr(self, key, **kwargs):
        text = self.languages[self.current_lang].get(key, key)
        # Format text with kwargs if needed
        if kwargs:
            try:
                text = text.format(**kwargs)
            except KeyError:
                pass
        return text

# =========================
# Custom INI Settings Class
# =========================
class IniSettings:
    def __init__(self, org_name, app_name):
        self.file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config", f"{app_name}.ini")
        self.config = configparser.ConfigParser()
        self._load_settings()

    def _load_settings(self):
        if os.path.exists(self.file_path):
            self.config.read(self.file_path)

    def value(self, key, default=None, type=None):
        section, option = self._parse_key(key)
        if self.config.has_option(section, option):
            val = self.config.get(section, option)
            if type == int:
                return int(val)
            elif type == float:
                return float(val)
            elif type == bool:
                return val.lower() == 'true'
            return val
        return default

    def setValue(self, key, value):
        section, option = self._parse_key(key)
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, option, str(value))

    def _parse_key(self, key):
        if '/' in key:
            parts = key.split('/')
            return parts[0], parts[1]
        return 'General', key

    def sync(self):
        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)
        with open(self.file_path, 'w') as configfile:
            self.config.write(configfile)

# ===============
# App Settings
# ===============
class Settings(IniSettings):
    def __init__(self):
        super().__init__("ScrollLockApp", "Settings")

    # Blocking logic
    def get_interval(self) -> float:
        return self.value("block_interval", 0.50, type=float)

    def set_interval(self, v: float):
        self.setValue("block_interval", v)

    def get_direction_change_threshold(self) -> int:
        return self.value("direction_change_threshold", 3, type=int)

    def set_direction_change_threshold(self, v: int):
        self.setValue("direction_change_threshold", v)

    # App control
    def get_enabled(self) -> bool:
        return self.value("enabled", True, type=bool)

    def set_enabled(self, v: bool):
        self.setValue("enabled", v)
        
    # Startup
    def get_startup(self) -> bool:
        return self.value("start_on_boot", False, type=bool)

    def set_startup(self, v: bool):
        self.setValue("start_on_boot", v)
        
    # Language
    def get_language(self) -> str:
        return self.value("language", "zh_CN", type=str)

    def set_language(self, v: str):
        self.setValue("language", v)

# ===========================
# Windows Startup Management
# ===========================
def configure_startup(enable: bool):
    """Create or remove HKCU Run entry for this script."""
    run_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    name = "MouseScrollLock"
    path = f'"{sys.executable}" "{os.path.abspath(__file__)}"'
    
    try:
        if enable:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key, 0, winreg.KEY_WRITE) as key:
                winreg.SetValueEx(key, name, 0, winreg.REG_SZ, path)
            print(f"开机启动已启用: {path}")
        else:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key, 0, winreg.KEY_WRITE) as key:
                try:
                    winreg.DeleteValue(key, name)
                except FileNotFoundError:
                    pass
            print("开机启动已禁用")
    except Exception as e:
        print(f"配置开机启动失败: {e}")

# =======================
# Low-Level Mouse Hooker
# =======================
class MouseHook:
    def __init__(self, settings: Settings, translator: Translator):
        self.settings = settings
        self.translator = translator

        # cached settings
        self.block_interval = self.settings.get_interval()
        self.enabled = self.settings.get_enabled()
        self.direction_change_threshold = self.settings.get_direction_change_threshold()

        # state
        self.last_dir = None            # 1: up, -1: down
        self.last_time = 0.0
        self._consecutive_opposite_events = 0

        # 统计信息 - 使用简单的变量而不是Qt信号
        self.total_events = 0
        self.blocked_events = 0
        self.current_direction = 0
        self.last_status = self.translator.tr('status_waiting')

        # win32 - 使用cdll而不是windll来避免126错误
        self.user32 = ctypes.cdll.user32
        self.kernel32 = ctypes.cdll.kernel32
        self.hook_id = None
        self.hook_cb = None
        
        # 确保DLL函数有正确的参数类型
        self.user32.SetWindowsHookExA.argtypes = [
            ctypes.c_int, 
            ctypes.c_void_p, 
            wintypes.HINSTANCE, 
            ctypes.c_uint
        ]
        self.user32.SetWindowsHookExA.restype = ctypes.c_void_p
        
        self.user32.CallNextHookEx.argtypes = [
            ctypes.c_void_p, 
            ctypes.c_int, 
            wintypes.WPARAM, 
            wintypes.LPARAM
        ]
        self.user32.CallNextHookEx.restype = ctypes.c_long
        
        self.kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
        self.kernel32.GetModuleHandleW.restype = wintypes.HMODULE

    def reload_settings(self):
        self.block_interval = self.settings.get_interval()
        self.enabled = self.settings.get_enabled()
        self.direction_change_threshold = self.settings.get_direction_change_threshold()
        self._consecutive_opposite_events = 0

    def get_status(self):
        """获取当前状态信息"""
        return {
            "total_events": self.total_events,
            "blocked_events": self.blocked_events,
            "current_direction": self.current_direction,
            "status": self.last_status
        }

    def start(self):
        # Callback type: LowLevelMouseProc
        CMPFUNC = ctypes.WINFUNCTYPE(
            ctypes.c_int, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM
        )

        def hook_proc(nCode, wParam, lParam):
            if nCode == 0 and wParam == win32con.WM_MOUSEWHEEL:
                if not self.enabled:
                    self.total_events += 1
                    self.current_direction = 0
                    self.last_status = self.translator.tr('status_disabled')
                    return self.user32.CallNextHookEx(self.hook_id, nCode, wParam, lParam)

                ms = ctypes.cast(lParam, ctypes.POINTER(MSLLHOOKSTRUCT)).contents
                # HIWORD(mouseData) contains the wheel delta, signed short
                delta_short = ctypes.c_short((ms.mouseData >> 16) & 0xFFFF).value
                current_dir = 1 if delta_short > 0 else -1

                now = time.time()

                # First event or interval elapsed: establish (or re-establish) a direction
                if (self.last_dir is None) or (now - self.last_time >= self.block_interval):
                    self.last_dir = current_dir
                    self.last_time = now
                    self._consecutive_opposite_events = 0
                    
                    self.total_events += 1
                    self.current_direction = current_dir
                    if current_dir > 0:
                        self.last_status = self.translator.tr('status_initial_up')
                    else:
                        self.last_status = self.translator.tr('status_initial_down')
                    
                    return self.user32.CallNextHookEx(self.hook_id, nCode, wParam, lParam)

                # Within interval:
                if current_dir == self.last_dir:
                    # continuing same direction keeps the "burst" alive
                    self.last_time = now
                    self._consecutive_opposite_events = 0
                    
                    self.total_events += 1
                    self.current_direction = current_dir
                    self.last_status = self.translator.tr('status_same_direction')
                    
                    return self.user32.CallNextHookEx(self.hook_id, nCode, wParam, lParam)
                else:
                    # opposite event inside interval
                    self._consecutive_opposite_events += 1
                    if self._consecutive_opposite_events >= self.direction_change_threshold:
                        # deliberate change — switch direction
                        self.last_dir = current_dir
                        self.last_time = now
                        self._consecutive_opposite_events = 0
                        
                        self.total_events += 1
                        self.current_direction = current_dir
                        self.last_status = self.translator.tr('status_direction_changed', count=self.direction_change_threshold)
                        
                        return self.user32.CallNextHookEx(self.hook_id, nCode, wParam, lParam)
                    else:
                        # suppress the jitter
                        self.blocked_events += 1
                        self.current_direction = current_dir
                        self.last_status = self.translator.tr('status_blocked', 
                                                             current=self._consecutive_opposite_events, 
                                                             threshold=self.direction_change_threshold)
                        
                        return 1  # block

            return self.user32.CallNextHookEx(self.hook_id, nCode, wParam, lParam)

        self.hook_cb = CMPFUNC(hook_proc)
        
        # 获取当前模块句柄
        hmod = self.kernel32.GetModuleHandleW(None)
        if not hmod:
            error_code = self.kernel32.GetLastError()
            print(f"获取模块句柄失败，错误代码: {error_code}")
            return False

        self.hook_id = self.user32.SetWindowsHookExA(
            win32con.WH_MOUSE_LL,
            self.hook_cb,
            hmod,
            0,
        )

        if not self.hook_id:
            error_code = self.kernel32.GetLastError()
            print(f"钩子安装失败，错误代码: {error_code}")
            return False

        print("鼠标钩子安装成功")
        
        # 初始化状态
        self.total_events = 0
        self.blocked_events = 0
        self.current_direction = 0
        self.last_status = self.translator.tr('status_waiting')

        # Message loop (runs in this thread)
        msg = wintypes.MSG()
        while True:
            b = self.user32.GetMessageA(ctypes.byref(msg), None, 0, 0)
            if b == 0:  # WM_QUIT
                break
            self.user32.TranslateMessage(ctypes.byref(msg))
            self.user32.DispatchMessageA(ctypes.byref(msg))

# 系统托盘类
class SystemTrayIcon(QtWidgets.QSystemTrayIcon):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        
        # 设置托盘图标
        app = QtWidgets.QApplication.instance()
        self.setIcon(app.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon))
        
        # 创建托盘菜单
        self.create_menu()
        
        # 连接激活信号
        self.activated.connect(self.on_tray_activated)
        
    def create_menu(self):
        """创建托盘菜单"""
        menu = QtWidgets.QMenu()
        
        show_action = menu.addAction(self.main_window.translator.tr('tray_show'))
        show_action.triggered.connect(self.main_window.show_normal)
        
        menu.addSeparator()
        
        quit_action = menu.addAction(self.main_window.translator.tr('tray_quit'))
        quit_action.triggered.connect(self.main_window.quit_application)
        
        self.setContextMenu(menu)
        
    def on_tray_activated(self, reason):
        """处理托盘图标激活"""
        if reason == QtWidgets.QSystemTrayIcon.DoubleClick:
            self.main_window.show_normal()

# 主窗口类
class ScrollLockApp(QtWidgets.QMainWindow):
    def __init__(self, settings, translator):
        super().__init__()
        self.settings = settings
        self.translator = translator
        self.hook = None
        self.hook_thread = None
        self.tray_icon = None
        
        # 初始化UI
        self.init_ui()
        
        # 创建状态更新定时器
        self.status_timer = QtCore.QTimer()
        self.status_timer.timeout.connect(self.update_status)
        self.status_timer.start(100)  # 每100ms更新一次状态
        
    def init_ui(self):
        # 设置窗口标题和大小
        self.setWindowTitle(self.translator.tr('app_title'))
        self.resize(450, 600)
        self.setMinimumSize(400, 550)
        
        # 设置窗口图标
        app = QtWidgets.QApplication.instance()
        self.setWindowIcon(app.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon))
        
        # 创建中央窗口部件
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QtWidgets.QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # 标题区域
        title_layout = QtWidgets.QHBoxLayout()
        
        # 应用图标
        app_icon = QtWidgets.QLabel()
        app_icon.setPixmap(QtWidgets.QApplication.instance().style().standardIcon(
            QtWidgets.QStyle.SP_ComputerIcon).pixmap(24, 24))
        title_layout.addWidget(app_icon)
        
        # 标题文本
        self.title_label = QtWidgets.QLabel(self.translator.tr('app_title'))
        self.title_label.setStyleSheet("font-size: 16pt; font-weight: bold;")
        title_layout.addWidget(self.title_label)
        
        # 语言选择下拉框
        self.language_combo = QtWidgets.QComboBox()
        self.language_combo.addItem("中文", "zh_CN")
        self.language_combo.addItem("English", "en_US")
        self.language_combo.setCurrentText("中文" if self.settings.get_language() == "zh_CN" else "English")
        self.language_combo.currentIndexChanged.connect(self.change_language)
        title_layout.addWidget(self.language_combo)
        
        main_layout.addLayout(title_layout)
        
        # 设置卡片
        self.settings_card = QtWidgets.QGroupBox(self.translator.tr('settings'))
        settings_layout = QtWidgets.QVBoxLayout(self.settings_card)
        
        # 设置表单
        form_layout = QtWidgets.QFormLayout()
        form_layout.setVerticalSpacing(10)
        form_layout.setHorizontalSpacing(15)
        
        # 时间阈值设置
        self.interval_label = QtWidgets.QLabel(self.translator.tr('block_interval'))
        self.interval_spin = QtWidgets.QDoubleSpinBox()
        self.interval_spin.setRange(0.05, 2.0)
        self.interval_spin.setSingleStep(0.05)
        self.interval_spin.setValue(self.settings.get_interval())
        self.interval_spin.setSuffix(self.translator.tr('seconds'))
        self.interval_spin.valueChanged.connect(self.update_settings)
        form_layout.addRow(self.interval_label, self.interval_spin)
        
        # 方向改变阈值设置
        self.threshold_label = QtWidgets.QLabel(self.translator.tr('direction_threshold'))
        self.threshold_spin = QtWidgets.QSpinBox()
        self.threshold_spin.setRange(1, 10)
        self.threshold_spin.setValue(self.settings.get_direction_change_threshold())
        self.threshold_spin.valueChanged.connect(self.update_settings)
        form_layout.addRow(self.threshold_label, self.threshold_spin)
        
        # 启用/禁用拦截
        self.enable_checkbox = QtWidgets.QCheckBox(self.translator.tr('enable_blocking'))
        self.enable_checkbox.setChecked(self.settings.get_enabled())
        self.enable_checkbox.stateChanged.connect(self.toggle_interception)
        form_layout.addRow("", self.enable_checkbox)
        
        # 开机启动设置
        self.startup_checkbox = QtWidgets.QCheckBox(self.translator.tr('startup'))
        self.startup_checkbox.setChecked(self.settings.get_startup())
        self.startup_checkbox.stateChanged.connect(self.toggle_startup)
        form_layout.addRow("", self.startup_checkbox)
        
        settings_layout.addLayout(form_layout)
        main_layout.addWidget(self.settings_card)
        
        # 状态卡片
        self.status_card = QtWidgets.QGroupBox(self.translator.tr('status'))
        status_layout = QtWidgets.QVBoxLayout(self.status_card)
        
        # 状态网格布局
        status_grid = QtWidgets.QGridLayout()
        status_grid.setVerticalSpacing(10)
        status_grid.setHorizontalSpacing(15)
        
        # 总事件数
        self.total_label = QtWidgets.QLabel(self.translator.tr('total_events'))
        self.total_events_label = QtWidgets.QLabel("0")
        status_grid.addWidget(self.total_label, 0, 0)
        status_grid.addWidget(self.total_events_label, 0, 1)
        
        # 拦截事件数
        self.blocked_label = QtWidgets.QLabel(self.translator.tr('blocked_events'))
        self.blocked_events_label = QtWidgets.QLabel("0")
        status_grid.addWidget(self.blocked_label, 0, 2)
        status_grid.addWidget(self.blocked_events_label, 0, 3)
        
        # 当前方向
        self.direction_label = QtWidgets.QLabel(self.translator.tr('current_direction'))
        self.direction_value_label = QtWidgets.QLabel(self.translator.tr('direction_none'))
        self.direction_value_label.setAlignment(QtCore.Qt.AlignCenter)
        status_grid.addWidget(self.direction_label, 1, 0)
        status_grid.addWidget(self.direction_value_label, 1, 1)
        
        # 状态信息
        self.status_info_label = QtWidgets.QLabel(self.translator.tr('status_label'))
        self.status_value_label = QtWidgets.QLabel("正在初始化...")
        self.status_value_label.setAlignment(QtCore.Qt.AlignCenter)
        status_grid.addWidget(self.status_info_label, 1, 2)
        status_grid.addWidget(self.status_value_label, 1, 3)
        
        status_layout.addLayout(status_grid)
        main_layout.addWidget(self.status_card)
        
        # 说明卡片
        self.info_card = QtWidgets.QGroupBox(self.translator.tr('how_it_works'))
        info_layout = QtWidgets.QVBoxLayout(self.info_card)
        
        # 说明文本
        self.info_text = QtWidgets.QLabel(self.translator.tr('how_it_works_text'))
        self.info_text.setWordWrap(True)
        info_layout.addWidget(self.info_text)
        
        main_layout.addWidget(self.info_card)
        
        # 底部按钮
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.setSpacing(10)
        
        self.restart_button = QtWidgets.QPushButton(self.translator.tr('restart_hook'))
        self.restart_button.clicked.connect(self.restart_hook)
        button_layout.addWidget(self.restart_button)
        
        self.minimize_button = QtWidgets.QPushButton(self.translator.tr('minimize_to_tray'))
        self.minimize_button.clicked.connect(self.hide_to_tray)
        button_layout.addWidget(self.minimize_button)
        
        button_layout.addStretch()
        
        self.quit_button = QtWidgets.QPushButton(self.translator.tr('quit'))
        self.quit_button.clicked.connect(self.quit_application)
        button_layout.addWidget(self.quit_button)
        
        main_layout.addLayout(button_layout)
        
    def set_tray_icon(self, tray_icon):
        """设置托盘图标引用"""
        self.tray_icon = tray_icon
        
    def change_language(self):
        """切换语言"""
        lang = self.language_combo.currentData()
        self.translator.set_language(lang)
        self.settings.set_language(lang)
        self.settings.sync()
        self.update_ui_text()
        
    def update_ui_text(self):
        """更新UI文本"""
        # 更新窗口标题
        self.setWindowTitle(self.translator.tr('app_title'))
        
        # 更新标题
        self.title_label.setText(self.translator.tr('app_title'))
        
        # 更新设置卡片
        self.settings_card.setTitle(self.translator.tr('settings'))
        self.interval_label.setText(self.translator.tr('block_interval'))
        self.interval_spin.setSuffix(self.translator.tr('seconds'))
        self.threshold_label.setText(self.translator.tr('direction_threshold'))
        self.enable_checkbox.setText(self.translator.tr('enable_blocking'))
        self.startup_checkbox.setText(self.translator.tr('startup'))
        
        # 更新状态卡片
        self.status_card.setTitle(self.translator.tr('status'))
        self.total_label.setText(self.translator.tr('total_events'))
        self.blocked_label.setText(self.translator.tr('blocked_events'))
        self.direction_label.setText(self.translator.tr('current_direction'))
        self.status_info_label.setText(self.translator.tr('status_label'))
        
        # 更新说明卡片
        self.info_card.setTitle(self.translator.tr('how_it_works'))
        self.info_text.setText(self.translator.tr('how_it_works_text'))
        
        # 更新按钮
        self.restart_button.setText(self.translator.tr('restart_hook'))
        self.minimize_button.setText(self.translator.tr('minimize_to_tray'))
        self.quit_button.setText(self.translator.tr('quit'))
        
        # 更新托盘菜单
        if self.tray_icon:
            self.tray_icon.create_menu()
            
        # 更新方向显示
        self.update_direction_display()
        
    def update_direction_display(self):
        """更新方向显示"""
        if hasattr(self, 'direction_value_label'):
            if self.hook and hasattr(self.hook, 'current_direction'):
                direction = self.hook.current_direction
                if direction == 1:
                    self.direction_value_label.setText(self.translator.tr('direction_up'))
                    self.direction_value_label.setStyleSheet("color: #107c10; font-weight: bold;")
                elif direction == -1:
                    self.direction_value_label.setText(self.translator.tr('direction_down'))
                    self.direction_value_label.setStyleSheet("color: #0078d7; font-weight: bold;")
                else:
                    self.direction_value_label.setText(self.translator.tr('direction_none'))
                    self.direction_value_label.setStyleSheet("color: #666666;")
        
    def start_hook(self):
        """启动鼠标钩子线程"""
        try:
            # 停止现有的钩子
            if self.hook_thread and self.hook_thread.is_alive():
                # 我们无法直接停止wheel.py的钩子线程，因为它运行在消息循环中
                # 这里我们只是重新启动线程
                pass
            
            # 重新初始化钩子
            self.hook = MouseHook(self.settings, self.translator)
            
            self.hook_thread = threading.Thread(target=self.hook.start)
            self.hook_thread.daemon = True
            self.hook_thread.start()
            print("钩子线程已启动")
            
            # 重置状态显示
            self.total_events_label.setText("0")
            self.blocked_events_label.setText("0")
            self.direction_value_label.setText(self.translator.tr('direction_none'))
            self.status_value_label.setText(self.translator.tr('hook_started'))
            
        except Exception as e:
            print(f"启动钩子线程失败: {e}")
            self.status_value_label.setText(self.translator.tr('hook_failed'))
        
    def restart_hook(self):
        """重启钩子"""
        print("重启鼠标钩子...")
        self.status_value_label.setText(self.translator.tr('hook_restarting'))
        # 由于wheel.py的钩子线程运行在消息循环中，我们无法直接停止
        # 这里我们只是重新加载设置
        if self.hook:
            self.hook.reload_settings()
        self.status_value_label.setText(self.translator.tr('hook_restarted'))
        
    def update_settings(self):
        """更新防抖设置"""
        self.settings.set_interval(self.interval_spin.value())
        self.settings.set_direction_change_threshold(self.threshold_spin.value())
        self.settings.set_enabled(self.enable_checkbox.isChecked())
        self.settings.sync()
        
        # 应用设置
        if self.hook:
            self.hook.reload_settings()
            
    def toggle_startup(self, state):
        """切换开机启动状态"""
        enabled = state == QtCore.Qt.Checked
        self.settings.set_startup(enabled)
        self.settings.sync()
        configure_startup(enabled)
        
    def toggle_interception(self, state):
        """切换拦截状态"""
        enabled = state == QtCore.Qt.Checked
        self.settings.set_enabled(enabled)
        self.settings.sync()
        if self.hook:
            self.hook.reload_settings()
    
    def update_status(self):
        """更新状态显示 - 通过定时器定期调用"""
        if self.hook:
            status = self.hook.get_status()
            self.total_events_label.setText(str(status["total_events"]))
            self.blocked_events_label.setText(str(status["blocked_events"]))
            
            # 更新方向显示
            direction = status["current_direction"]
            if direction == 1:
                self.direction_value_label.setText(self.translator.tr('direction_up'))
                self.direction_value_label.setStyleSheet("color: #107c10; font-weight: bold;")
            elif direction == -1:
                self.direction_value_label.setText(self.translator.tr('direction_down'))
                self.direction_value_label.setStyleSheet("color: #0078d7; font-weight: bold;")
            else:
                self.direction_value_label.setText(self.translator.tr('direction_none'))
                self.direction_value_label.setStyleSheet("color: #666666;")
                
            self.status_value_label.setText(status["status"])
        
    def hide_to_tray(self):
        """隐藏到系统托盘"""
        self.hide()
        if self.tray_icon:
            self.tray_icon.showMessage(
                self.translator.tr('tray_message_title'), 
                self.translator.tr('tray_message_content'),
                QtWidgets.QSystemTrayIcon.Information, 
                2000
            )
        
    def show_normal(self):
        """从托盘恢复显示"""
        self.show()
        self.raise_()
        self.activateWindow()
        
    def quit_application(self):
        """退出应用程序"""
        # 保存设置
        self.settings.sync()
        
        # 停止状态更新定时器
        if self.status_timer and self.status_timer.isActive():
            self.status_timer.stop()
        
        # 退出应用
        QtWidgets.QApplication.quit()
        
    def closeEvent(self, event):
        """处理关闭事件 - 改为最小化到托盘"""
        event.ignore()  # 忽略关闭事件
        self.hide_to_tray()

# =========
#  main()
# =========
def main():
    # 启用高DPI缩放
    if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
        QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
        QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
    
    app = QtWidgets.QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)  # 重要：确保关闭窗口不会退出应用
    
    # 设置应用程序属性
    app.setApplicationName("鼠标滚轮防抖工具")
    app.setApplicationVersion("1.0")
    
    # 加载设置和翻译器
    settings = Settings()
    translator = Translator()
    translator.set_language(settings.get_language())
    
    # 配置开机启动（根据设置）
    if settings.get_startup():
        configure_startup(True)
    
    # 创建主窗口
    main_window = ScrollLockApp(settings, translator)
    
    # 创建系统托盘图标
    tray_icon = SystemTrayIcon(main_window)
    tray_icon.show()
    
    # 设置主窗口的托盘图标引用
    main_window.set_tray_icon(tray_icon)
    
    # 显示主窗口
    main_window.show()
    
    # 启动钩子线程
    main_window.start_hook()
    
    # 运行应用程序
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()