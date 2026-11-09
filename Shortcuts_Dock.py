import os
import sys
import time
import ctypes
import json
import re
import shutil
import winreg
from ctypes import wintypes
import win32com.client
from pathlib import Path


from PySide6.QtCore import Qt, QSize, QFileInfo, QTimer, QSettings, QPropertyAnimation, QEasingCurve, QPoint, QMimeData, QUrl, QEvent
from PySide6.QtGui import QAction, QActionGroup, QCursor, QDrag, QIcon, QPixmap, QPainter, QColor, QPen, QPainterPath
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QBoxLayout,
    QPushButton,
    QMenu,
    QSystemTrayIcon,
    QStyle,
    QFileIconProvider,
    QMessageBox,
    QWidgetAction,
    QSlider,
    QLabel,
    QVBoxLayout,
    QToolTip,
    QInputDialog,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLineEdit,
    QScrollArea,
    QScrollBar,
    QSizePolicy,
)


# Store user shortcuts beside the actual EXE, never inside PyInstaller's
# temporary _MEIPASS extraction directory. In source mode, use the .py folder.
BASE_DIR = (Path(sys.executable).resolve().parent if getattr(sys, "frozen", False)
            else Path(__file__).resolve().parent)
SHORTCUTS_DIR = BASE_DIR / "Shortcuts"
VERSION_INFO_PATH = (Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "version_info.txt")


def application_version():
    """Read the same ProductVersion field used by PyInstaller's version resource."""
    try:
        data = VERSION_INFO_PATH.read_text(encoding="utf-8-sig")
        match = re.search(r"StringStruct\(\s*['\"]ProductVersion['\"]\s*,\s*['\"]([^'\"]+)['\"]", data)
        if match:
            return match.group(1)
    except OSError:
        pass
    return "Unknown"

ICON_SIZE = 24
BUTTON_SIZE = 40
SPACING = 4
MARGIN = 4

POSITIONS = ("Bottom", "Top", "Left", "Right")

# Read actual physical mouse coordinates and monitor bounds from Windows.
# This avoids rounding at 200%/300% display scaling.
user32 = ctypes.WinDLL("user32", use_last_error=True)
MONITOR_DEFAULTTONEAREST = 2

class MONITORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("rcMonitor", wintypes.RECT),
        ("rcWork", wintypes.RECT),
        ("dwFlags", wintypes.DWORD),
    ]

user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int,
                                 ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                 wintypes.UINT]
user32.SetWindowPos.restype = wintypes.BOOL
user32.GetCursorPos.argtypes = [ctypes.POINTER(wintypes.POINT)]
user32.GetCursorPos.restype = wintypes.BOOL
user32.MonitorFromPoint.argtypes = [wintypes.POINT, wintypes.DWORD]
user32.MonitorFromPoint.restype = wintypes.HANDLE
user32.GetMonitorInfoW.argtypes = [wintypes.HANDLE, ctypes.POINTER(MONITORINFO)]
user32.GetMonitorInfoW.restype = wintypes.BOOL


def native_cursor_and_monitor():
    point = wintypes.POINT()
    if not user32.GetCursorPos(ctypes.byref(point)):
        return None
    handle = user32.MonitorFromPoint(point, MONITOR_DEFAULTTONEAREST)
    info = MONITORINFO()
    info.cbSize = ctypes.sizeof(info)
    if not handle or not user32.GetMonitorInfoW(handle, ctypes.byref(info)):
        return None
    return point, info.rcMonitor


def windows_system_dark_theme():
    """Taskbar/system theme (not the independent app theme)."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            return winreg.QueryValueEx(key, "SystemUsesLightTheme")[0] == 0
    except OSError:
        return True


def windows_taskbar_tint():
    """Neutral taskbar-like surface; do not mix in DWM's wallpaper/accent tint."""
    dark = windows_system_dark_theme()
    transparency = True
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            transparency = bool(winreg.QueryValueEx(key, "EnableTransparency")[0])
    except OSError:
        pass
    # Windows 11 taskbar uses a neutral system surface unless accent-on-taskbar
    # is enabled. The old DwmGetColorizationColor blending made it brown.
    rgb = (32, 32, 32) if dark else (243, 243, 243)
    alpha = 205 if transparency else 255
    border = (255, 255, 255, 48) if dark else (30, 30, 30, 42)
    hover = (255, 255, 255, 48) if dark else (0, 0, 0, 30)
    pressed = (255, 255, 255, 76) if dark else (0, 0, 0, 55)
    return (f"rgba({rgb[0]}, {rgb[1]}, {rgb[2]}, {alpha})",
            f"rgba({border[0]}, {border[1]}, {border[2]}, {border[3]})",
            f"rgba({hover[0]}, {hover[1]}, {hover[2]}, {hover[3]})",
            f"rgba({pressed[0]}, {pressed[1]}, {pressed[2]}, {pressed[3]})")


def style_icon_context_menu(menu):
    """Explicit system-matched colors: Qt popup palettes can disagree with Windows."""
    if windows_system_dark_theme():
        background, foreground, selection, separator = (
            "#2b2b2b", "#f5f5f5", "#414141", "#555555")
    else:
        background, foreground, selection, separator = (
            "#f9f9f9", "#202020", "#e5e5e5", "#dedede")
    menu.setStyleSheet(f"""
        QMenu {{ background-color: {background}; color: {foreground};
                 border: 1px solid {separator}; padding: 4px 2px; }}
        QMenu::item {{ background-color: transparent; color: {foreground};
                       padding: 6px 24px 6px 14px; }}
        QMenu::item:selected {{ background-color: {selection}; color: {foreground}; }}
        QMenu::item:disabled {{ color: {'#858585' if windows_system_dark_theme() else '#888888'}; }}
        QMenu::separator {{ background: {separator}; height: 1px; margin: 4px 8px; }}
    """)


def set_native_dialog_theme(dialog):
    """Set the Win32 title-bar theme before the dialog is displayed."""
    dark = windows_system_dark_theme()
    # Creating the native handle here does not show the dialog.
    hwnd = int(dialog.winId())
    dwm = ctypes.WinDLL("dwmapi", use_last_error=True)
    dwm.DwmSetWindowAttribute.argtypes = [wintypes.HWND, wintypes.DWORD,
                                         ctypes.c_void_p, wintypes.DWORD]
    dwm.DwmSetWindowAttribute.restype = ctypes.c_long
    value = ctypes.c_int(1 if dark else 0)
    for attribute in (20, 19):  # Windows 10/11 immersive dark title bar
        if dwm.DwmSetWindowAttribute(hwnd, attribute, ctypes.byref(value),
                                     ctypes.sizeof(value)) == 0:
            break


def style_rename_dialog(dialog):
    """Apply the Windows system theme to the entire rename dialog."""
    dark = windows_system_dark_theme()
    bg = "#292929" if dark else "#f9f9f9"
    fg = "#f5f5f5" if dark else "#202020"
    field = "#383838" if dark else "#ffffff"
    border = "#666666" if dark else "#b8b8b8"
    hover = "#454545" if dark else "#e8e8e8"
    dialog.setStyleSheet(f"""
        QInputDialog {{ background-color: {bg}; color: {fg}; }}
        QInputDialog QLabel {{ color: {fg}; background: transparent; }}
        QInputDialog QLineEdit {{ background-color: {field}; color: {fg};
            border: 1px solid {border}; border-radius: 4px; padding: 5px; }}
        QInputDialog QPushButton {{ background-color: {field}; color: {fg};
            border: 1px solid {border}; border-radius: 4px;
            padding: 5px 16px; min-width: 65px; }}
        QInputDialog QPushButton:hover {{ background-color: {hover}; }}
    """)


MIME_DOCK = "application/x-shortcuts-dock-item"


class RoundedTooltip(QLabel):
    """Paint the complete tooltip on a transparent native window surface."""

    def __init__(self, parent=None, flags=Qt.WindowType.Widget):
        super().__init__(parent, flags)
        self.tip_background = QColor("#303030")
        self.tip_foreground = QColor("#ffffff")
        self.tip_border = QColor("#646464")

    def set_tip_colors(self, background, foreground, border):
        self.tip_background = QColor(background)
        self.tip_foreground = QColor(foreground)
        self.tip_border = QColor(border)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = self.rect().adjusted(1, 1, -1, -1)
        painter.setPen(QPen(self.tip_border, 1))
        painter.setBrush(self.tip_background)
        painter.drawRoundedRect(rect, 6, 6)
        painter.setPen(self.tip_foreground)
        painter.drawText(self.contentsRect(), self.alignment(), self.text())
        painter.end()


class ShortcutButton(QPushButton):
    def __init__(self, dock, path):
        super().__init__(dock.content)
        self.dock = dock
        self.path = path
        self.press_point = None
        self.setAcceptDrops(True)
        self.setMouseTracking(True)

    def paintEvent(self, event):
        """Paint the taskbar-style tile explicitly, independent of inherited QSS.

        The scroll content has a transparent stylesheet which can override
        QPushButton hover backgrounds on some Qt/Windows configurations.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        dark = windows_system_dark_theme()
        if self.isDown():
            tint = QColor(255, 255, 255, 76) if dark else QColor(0, 0, 0, 55)
        elif self.underMouse():
            tint = QColor(255, 255, 255, 48) if dark else QColor(0, 0, 0, 30)
        else:
            tint = None
        if tint is not None:
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(tint)
            painter.drawRoundedRect(self.rect().adjusted(1, 1, -1, -1), 5, 5)
        # Paint the QIcon through Qt itself. This preserves its actual icon
        # engine (including Windows shell icons) and handles high-DPI scaling.
        icon = self.icon()
        if not icon.isNull():
            size = self.iconSize()
            icon_rect = self.rect()
            icon_rect.setLeft((self.width() - size.width()) // 2)
            icon_rect.setTop((self.height() - size.height()) // 2)
            icon_rect.setWidth(size.width())
            icon_rect.setHeight(size.height())
            icon.paint(painter, icon_rect, Qt.AlignmentFlag.AlignCenter,
                       QIcon.Mode.Normal, QIcon.State.Off)
        painter.end()

    def enterEvent(self, event):
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.update()
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.press_point = event.position().toPoint()
        super().mousePressEvent(event)
        self.update()

    def mouseMoveEvent(self, event):
        if (self.press_point is not None and
                event.buttons() & Qt.MouseButton.LeftButton and
                (event.position().toPoint() - self.press_point).manhattanLength() >=
                QApplication.startDragDistance()):
            self.press_point = None
            self.dock.start_shortcut_drag(self)
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self.press_point = None
        super().mouseReleaseEvent(event)
        self.update()

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        style_icon_context_menu(menu)
        open_action = menu.addAction("Open")
        rename_action = menu.addAction("Rename")
        remove_action = menu.addAction("Remove from Dock")
        self.dock.icon_menu_open = True
        self.dock.hide_deadline = None
        try:
            menu.adjustSize()
            menu_size = menu.sizeHint()
            dock_anchor = self.dock.mapToGlobal(QPoint(0, 0))
            icon_center = self.mapToGlobal(self.rect().center())
            if self.dock.position == "Bottom":
                point = QPoint(icon_center.x() - menu_size.width() // 2,
                               dock_anchor.y() - menu_size.height())
            elif self.dock.position == "Top":
                point = QPoint(icon_center.x() - menu_size.width() // 2,
                               dock_anchor.y() + self.dock.height())
            elif self.dock.position == "Left":
                point = QPoint(dock_anchor.x() + self.dock.width(),
                               icon_center.y() - menu_size.height() // 2)
            else:
                point = QPoint(dock_anchor.x() - menu_size.width(),
                               icon_center.y() - menu_size.height() // 2)
            # Keep the whole popup on-screen, nudging it inward near edges.
            available = self.dock.get_target_screen().availableGeometry()
            point.setX(max(available.left(), min(point.x(), available.right() - menu_size.width() + 1)))
            point.setY(max(available.top(), min(point.y(), available.bottom() - menu_size.height() + 1)))
            selected = menu.exec(point)
        finally:
            self.dock.icon_menu_open = False
            self.dock.hide_deadline = None
        if selected == open_action:
            self.dock.launch_shortcut(self.path)
        elif selected == rename_action:
            self.dock.rename_shortcut(self.path)
        elif selected == remove_action:
            self.dock.remove_shortcut(self.path)

    def dragEnterEvent(self, event):
        self.dock.handle_drag_hover(event, self)

    def dragMoveEvent(self, event):
        self.dock.handle_drag_hover(event, self)

    def dragLeaveEvent(self, event):
        event.accept()

    def dropEvent(self, event):
        self.dock.handle_drop(event, self)


class ShortcutsDock(QWidget):
    def __init__(self):
        super().__init__()

        SHORTCUTS_DIR.mkdir(exist_ok=True)

        self.settings = QSettings(
            "ShortcutsDock", "ShortcutsDock"
        )

        self.position = self.settings.value(
            "position", "Bottom"
        )

        if self.position not in POSITIONS:
            self.position = "Bottom"

        self.primary_only = self.settings.value(
            "primary_only", False, type=bool
        )

        self.offset = max(0, min(200, int(self.settings.value("offset", 0))))
        self.icon_spacing = max(0, min(40, int(self.settings.value("icon_spacing", SPACING))))
        self.reverse_scroll = self.settings.value("reverse_scroll", False, type=bool)

        self.auto_hide = self.settings.value("auto_hide", False, type=bool)
        self.hide_delay = max(0, min(3000, int(self.settings.value("hide_delay", 500))))
        self.hide_deadline = None
        self.appearance_delay = max(0, min(3000, int(self.settings.value("appearance_delay", 0))))
        self.appearance_since = None
        self.dragging = False
        self.drag_button = None
        self.drag_original_order = None
        self.drag_placeholder = None
        self.push_animations = []
        self.drag_preview_index = None
        self.external_drag_active = False
        self.tooltip_button = None
        self.tooltip_started = None
        self.tooltip_visible_for = None
        self.icon_tooltip = RoundedTooltip(None, Qt.WindowType.ToolTip | Qt.WindowType.FramelessWindowHint)
        self.icon_tooltip.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        # The tooltip is its own native window. Without a translucent window
        # surface, Qt paints the area outside the rounded QLabel as opaque.
        self.icon_tooltip.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.icon_tooltip.setAutoFillBackground(False)
        self.icon_tooltip.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.icon_tooltip.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_tooltip.setMargin(7)
        self.known_shortcut_names = set()
        self.icon_menu_open = False
        self.tray_menu_open = False
        self.rename_dialog_open = False
        self.setAcceptDrops(True)
        self.activation_armed = False
        self.last_activation_zone = False
        self.current_screen = None
        self.animating = False
        self.animation = QPropertyAnimation(self, b"pos", self)
        self.animation.setDuration(230)
        self.animation.setEasingCurve(QEasingCurve.Type.OutCubic)
        self.animation.finished.connect(self.animation_finished)
        self.animation_hiding = False
        self.visible_position = QPoint(0, 0)

        self.setWindowTitle("Shortcuts Dock")

        self.setWindowFlags(
            Qt.WindowType.Tool
            | Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
        )

        self.setAttribute(
            Qt.WidgetAttribute.WA_TranslucentBackground
        )

        self._last_theme_state = self.read_theme_state()
        bg, border, hover, pressed = windows_taskbar_tint()
        self.setStyleSheet(f"""
            QWidget#dockPanel {{
                background-color: {bg};
                border: 1px solid {border};
                border-radius: 8px;
            }}
            QWidget#dockPanel QPushButton {{
                background-color: transparent;
                border: none;
                border-radius: 5px;
            }}
            QWidget#dockPanel QPushButton:hover {{ background-color: {hover}; }}
            QWidget#dockPanel QPushButton:pressed {{ background-color: {pressed}; }}
        """)

        self.panel = QWidget(self)
        self.panel.setObjectName("dockPanel")
        self.scroll_area = QScrollArea(self.panel)
        self.scroll_area.setFrameShape(QScrollArea.Shape.NoFrame)
        self.scroll_area.setStyleSheet("QScrollArea, QScrollArea > QWidget, QScrollArea > QWidget > QWidget { background: transparent; border: 0px; }")
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setWidgetResizable(False)
        self.scroll_area.setAcceptDrops(True)
        self.scroll_area.viewport().setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.scroll_area.viewport().setAutoFillBackground(False)
        self.scroll_area.viewport().setAcceptDrops(True)
        self.scroll_area.viewport().installEventFilter(self)
        self.content = QWidget()
        self.content.setStyleSheet("background: transparent; border: none;")
        self.content.setAutoFillBackground(False)
        self.content.setAcceptDrops(True)
        self.content.installEventFilter(self)
        self.scroll_area.setWidget(self.content)

        self.layout = QBoxLayout(
            QBoxLayout.Direction.LeftToRight,
            self.content
        )

        self.layout.setContentsMargins(
            MARGIN, MARGIN, MARGIN, MARGIN
        )
        self.layout.setSpacing(self.icon_spacing)

        self.icon_provider = QFileIconProvider()

        self.load_shortcuts()
        self.known_shortcut_names = {p.name.casefold() for p in self.shortcut_paths()}
        self.create_tray()
        self.update_orientation()

        # Monitor the mouse and display geometry.
        self.monitor_timer = QTimer(self)
        self.monitor_timer.timeout.connect(
            self.update_monitor
        )
        self.monitor_timer.start(10 if self.auto_hide else 250)
        # Detect additions/deletions made directly in the Shortcuts folder.
        # Polling also catches Explorer operations that filesystem watchers can miss.
        self.shortcuts_timer = QTimer(self)
        self.shortcuts_timer.timeout.connect(self.check_shortcuts_folder)
        self.shortcuts_timer.start(700)
        # Windows does not reliably deliver theme-change events to tray-only Qt
        # windows. Poll the small registry values so changes apply while running.
        self.theme_timer = QTimer(self)
        self.theme_timer.timeout.connect(self.check_windows_theme)
        self.theme_timer.start(250)
        # Show shortcut names on Qt hover events, without a polling interval.
        if self.auto_hide:
            self.hide()

    @staticmethod
    def read_theme_state():
        transparency = True
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
                transparency = bool(winreg.QueryValueEx(key, "EnableTransparency")[0])
        except OSError:
            pass
        return windows_system_dark_theme(), transparency

    def check_windows_theme(self):
        state = self.read_theme_state()
        if state == self._last_theme_state:
            return
        self._last_theme_state = state
        bg, border, hover, pressed = windows_taskbar_tint()
        self.setStyleSheet(f"""
            QWidget#dockPanel {{
                background-color: {bg};
                border: 1px solid {border};
                border-radius: 8px;
            }}
            QWidget#dockPanel QPushButton {{
                background-color: transparent;
                border: none;
                border-radius: 5px;
            }}
            QWidget#dockPanel QPushButton:hover {{ background-color: {hover}; }}
            QWidget#dockPanel QPushButton:pressed {{ background-color: {pressed}; }}
        """)
        dark = state[0]
        self.icon_tooltip.set_tip_colors(
            "#303030" if dark else "#fafafa",
            "#ffffff" if dark else "#202020",
            "#646464" if dark else "#b8b8b8")
        if hasattr(self, 'tray'):
            menu = self.tray.contextMenu()
            if menu is not None:
                self.style_tray_menu(menu)
                self.style_tray_slider_labels()
        self.panel.update()
        self.update()

    def get_shortcut_icon(self, shortcut):
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            link = shell.CreateShortcut(str(shortcut))

            icon_location = link.IconLocation
            target_path = link.TargetPath

            # Use the shortcut's custom icon file when available.
            if icon_location:
                icon_path = os.path.expandvars(
                    icon_location.rsplit(",", 1)[0].strip('"')
                )
                if os.path.isfile(icon_path):
                    icon = self.icon_provider.icon(QFileInfo(icon_path))
                    if not icon.isNull():
                        return icon

            # The target's icon has no shortcut-arrow overlay.
            if target_path:
                target_path = os.path.expandvars(target_path)
                # Folder shortcuts launch Explorer with the folder as an argument.
                if Path(target_path).name.lower() == "explorer.exe":
                    folder = str(getattr(link, "Arguments", "")).strip().strip('"')
                    if os.path.isdir(folder):
                        return self.icon_provider.icon(QFileInfo(folder))
                if os.path.isdir(target_path):
                    return self.icon_provider.icon(QFileInfo(target_path))
                if os.path.isfile(target_path):
                    icon = self.icon_provider.icon(QFileInfo(target_path))
                    if not icon.isNull():
                        return icon

        except Exception as error:
            print(f"Icon extraction error for {shortcut.name}: {error}")

        # If the shortcut cannot be resolved, use a generic file icon
        # instead of the .lnk icon with its arrow overlay.
        return self.icon_provider.icon(
            QFileInfo(str(shortcut.with_suffix(".exe")))
        )

    def shortcut_paths(self):
        files = [p for p in SHORTCUTS_DIR.iterdir()
                 if p.is_file() and p.suffix.lower() == ".lnk"]
        try:
            order = json.loads(self.settings.value("shortcut_order", "[]"))
            if not isinstance(order, list):
                order = []
        except (TypeError, ValueError):
            order = []
        positions = {name.casefold(): i for i, name in enumerate(order)
                     if isinstance(name, str)}
        return sorted(files, key=lambda p: (positions.get(p.name.casefold(), 999999),
                                            p.name.casefold()))

    def check_shortcuts_folder(self):
        current = {p.name.casefold() for p in SHORTCUTS_DIR.iterdir()
                   if p.is_file() and p.suffix.lower() == ".lnk"}
        if current == self.known_shortcut_names or self.dragging or self.icon_menu_open:
            return
        # Compact the stored order when a shortcut disappears, and append
        # newly copied shortcuts in their first-seen order.
        paths = self.shortcut_paths()
        self.save_order(paths)
        self.known_shortcut_names = current
        self.load_shortcuts()
        self.position_dock()

    def clear_icon_tooltip(self):
        self.tooltip_button = None
        self.tooltip_started = None
        self.tooltip_visible_for = None
        self.icon_tooltip.hide()

    def hideEvent(self, event):
        self.clear_icon_tooltip()
        super().hideEvent(event)

    def update_icon_tooltip(self, entered_button=None):
        if (not self.isVisible() or self.animating or self.dragging
                or self.icon_menu_open or self.tray_menu_open):
            self.tooltip_button = None
            self.tooltip_started = None
            self.tooltip_visible_for = None
            self.icon_tooltip.hide()
            return

        cursor = QCursor.pos()
        hovered = None
        if (isinstance(entered_button, ShortcutButton) and entered_button.isVisible()
                and entered_button.rect().contains(entered_button.mapFromGlobal(cursor))):
            hovered = entered_button
        else:
            for i in range(self.layout.count()):
                widget = self.layout.itemAt(i).widget()
                if (isinstance(widget, ShortcutButton) and widget.isVisible()
                        and widget.rect().contains(widget.mapFromGlobal(cursor))):
                    hovered = widget
                    break

        if hovered is not self.tooltip_button:
            self.tooltip_button = hovered
            self.tooltip_started = time.monotonic() if hovered else None
            self.tooltip_visible_for = None
            self.icon_tooltip.hide()

        if hovered and self.tooltip_visible_for is not hovered:
            dark = windows_system_dark_theme()
            background = "#303030" if dark else "#fafafa"
            foreground = "#ffffff" if dark else "#202020"
            border = "#646464" if dark else "#b8b8b8"
            self.icon_tooltip.set_tip_colors(background, foreground, border)
            self.icon_tooltip.setText(hovered.path.stem)
            self.icon_tooltip.adjustSize()
            width, height = self.icon_tooltip.width(), self.icon_tooltip.height()
            icon_center = hovered.mapToGlobal(hovered.rect().center())
            dock_rect = self.geometry()
            gap = 8
            if self.position == "Bottom":
                x = icon_center.x() - width // 2
                y = dock_rect.top() - height - gap
            elif self.position == "Top":
                x = icon_center.x() - width // 2
                y = dock_rect.bottom() + gap
            elif self.position == "Left":
                x = dock_rect.right() + gap
                y = icon_center.y() - height // 2
            else:
                x = dock_rect.left() - width - gap
                y = icon_center.y() - height // 2
            screen = QApplication.screenAt(icon_center) or QApplication.primaryScreen()
            if screen:
                bounds = screen.availableGeometry()
                x = max(bounds.left() + 3, min(x, bounds.right() - width - 2))
                y = max(bounds.top() + 3, min(y, bounds.bottom() - height - 2))
            self.icon_tooltip.move(x, y)
            self.icon_tooltip.show()
            self.tooltip_visible_for = hovered

    def save_order(self, paths):
        self.settings.setValue("shortcut_order", json.dumps([p.name for p in paths]))

    def load_shortcuts(self):
        while self.layout.count():
            item = self.layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        shortcuts = self.shortcut_paths()
        if not shortcuts:
            button = QPushButton("+")
            button.setToolTip("Drag files, folders, or shortcuts onto the dock")
            button.setFixedSize(BUTTON_SIZE, BUTTON_SIZE)
            button.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
            self.layout.addWidget(button)

        for shortcut in shortcuts:
            button = ShortcutButton(self, shortcut)
            button.installEventFilter(self)
            button.setIcon(self.get_shortcut_icon(shortcut))
            button.setIconSize(QSize(ICON_SIZE, ICON_SIZE))
            button.setFixedSize(BUTTON_SIZE, BUTTON_SIZE)
            button.setToolTip("")
            button.clicked.connect(lambda checked=False, path=shortcut:
                                   self.launch_shortcut(path))
            self.layout.addWidget(button)
        self.update_orientation()
        self.known_shortcut_names = {p.name.casefold() for p in shortcuts}

    def unique_path(self, folder, stem, suffix=".lnk"):
        path = folder / (stem + suffix)
        number = 1
        while path.exists():
            path = folder / f"{stem} ({number}){suffix}"
            number += 1
        return path

    def create_shortcut_from_path(self, source):
        source = Path(source)
        if not source.exists():
            return None
        destination = self.unique_path(SHORTCUTS_DIR, source.stem if source.is_file() else source.name)
        if source.suffix.lower() == ".lnk":
            shutil.copy2(source, destination)
        else:
            shell = win32com.client.Dispatch("WScript.Shell")
            link = shell.CreateShortcut(str(destination))
            if source.is_dir():
                # Explorer reliably opens directory shortcuts, including paths with spaces.
                link.TargetPath = str(Path(os.environ.get("WINDIR", r"C:\Windows")) / "explorer.exe")
                link.Arguments = f'"{source}"'
                link.WorkingDirectory = str(source)
            else:
                link.TargetPath = str(source)
                link.WorkingDirectory = str(source.parent)
            link.Save()
        return destination

    def add_sources(self, sources, target=None, before=True):
        order = self.shortcut_paths()
        created = []
        for source in sources:
            try:
                new_path = self.create_shortcut_from_path(source)
                if new_path:
                    created.append(new_path)
            except Exception as error:
                QMessageBox.warning(self, "Add Shortcut", f"Could not add {source}:\n{error}")
        if not created:
            return False
        if target in order:
            index = order.index(target) + (0 if before else 1)
            order[index:index] = created
        else:
            order.extend(created)
        self.save_order(order)
        self.load_shortcuts()
        return True

    def reorder_shortcut(self, source, target, before):
        order = self.shortcut_paths()
        if source not in order or target not in order or source == target:
            return
        order.remove(source)
        index = order.index(target) + (0 if before else 1)
        order.insert(index, source)
        self.save_order(order)
        self.load_shortcuts()

    def rename_shortcut(self, path):
        # A compact, frameless editor styled like the dock's context menus.
        # Apply its palette before showing, avoiding native-title-bar theme flashes.
        dark = windows_system_dark_theme()
        bg = "#2b2b2b" if dark else "#f9f9f9"
        fg = "#f5f5f5" if dark else "#202020"
        field = "#383838" if dark else "#ffffff"
        border = "#555555" if dark else "#dedede"
        hover = "#414141" if dark else "#e5e5e5"
        dialog = QDialog(self)
        dialog.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint |
                              Qt.WindowType.WindowStaysOnTopHint)
        dialog.setModal(True)
        dialog.setObjectName("simpleRenameDialog")
        dialog.setStyleSheet(f"""
            QDialog#simpleRenameDialog {{ background-color: {bg};
                border: 1px solid {border}; }}
            QLabel {{ color: {fg}; background: transparent; border: none; }}
            QLineEdit {{ color: {fg}; background: {field};
                border: 1px solid {border}; padding: 4px; }}
            QPushButton {{ color: {fg}; background: transparent;
                border: 1px solid {border}; padding: 4px 12px; min-width: 55px; }}
            QPushButton:hover {{ background: {hover}; }}
        """)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)
        layout.addWidget(QLabel("Rename shortcut:"))
        editor = QLineEdit(path.stem)
        editor.selectAll()
        editor.returnPressed.connect(dialog.accept)
        layout.addWidget(editor)
        buttons = QHBoxLayout()
        buttons.setSpacing(6)
        buttons.addStretch()
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(dialog.reject)
        confirm = QPushButton("OK")
        confirm.setDefault(True)
        confirm.clicked.connect(dialog.accept)
        buttons.addWidget(cancel)
        buttons.addWidget(confirm)
        layout.addLayout(buttons)
        dialog.setFixedSize(285, 122)

        # Keep the editor on the inward side of the dock and within the screen.
        dock_rect = self.frameGeometry()
        screen_rect = self.get_target_screen().availableGeometry()
        gap = 10
        if self.position == "Bottom":
            x = dock_rect.center().x() - dialog.width() // 2
            y = dock_rect.top() - dialog.height() - gap
        elif self.position == "Top":
            x = dock_rect.center().x() - dialog.width() // 2
            y = dock_rect.bottom() + gap
        elif self.position == "Left":
            x = dock_rect.right() + gap
            y = dock_rect.center().y() - dialog.height() // 2
        else:
            x = dock_rect.left() - dialog.width() - gap
            y = dock_rect.center().y() - dialog.height() // 2
        x = max(screen_rect.left(), min(x, screen_rect.right() - dialog.width() + 1))
        y = max(screen_rect.top(), min(y, screen_rect.bottom() - dialog.height() + 1))
        dialog.move(x, y)
        self.rename_dialog_open = True
        self.hide_deadline = None
        try:
            ok = dialog.exec() == QDialog.DialogCode.Accepted
        finally:
            self.rename_dialog_open = False
            self.hide_deadline = None
            self.appearance_since = None
        name = editor.text()
        if not ok or not name.strip():
            return
        name = name.strip()
        if any(c in name for c in '<>:"/\\|?*') or name.endswith(('.', ' ')):
            QMessageBox.warning(self, "Rename Shortcut", "Invalid Windows filename.")
            return
        new_path = path.with_name(name + ".lnk")
        if new_path != path and new_path.exists():
            QMessageBox.warning(self, "Rename Shortcut", "A shortcut with that name already exists.")
            return
        order = [new_path if p == path else p for p in self.shortcut_paths()]
        try:
            path.rename(new_path)
            self.save_order(order)
            self.load_shortcuts()
        except OSError as error:
            QMessageBox.warning(self, "Rename Shortcut", str(error))

    def remove_shortcut(self, path):
        try:
            path.unlink()
            self.save_order([p for p in self.shortcut_paths() if p != path])
            self.load_shortcuts()
        except OSError as error:
            QMessageBox.warning(self, "Remove Shortcut", str(error))

    def animate_icon_layout(self, change):
        """Animate surviving icons between their old and new layout positions."""
        widgets = [self.layout.itemAt(i).widget() for i in range(self.layout.count())]
        old = {w: w.pos() for w in widgets if isinstance(w, ShortcutButton) and w.isVisible()}
        for anim in self.push_animations:
            anim.stop()
        self.push_animations.clear()
        change()
        self.layout.invalidate()
        self.layout.activate()
        for w, start in old.items():
            if self.layout.indexOf(w) < 0:
                continue
            end = w.pos()
            if start == end:
                continue
            w.move(start)
            anim = QPropertyAnimation(w, b"pos", self)
            anim.setDuration(145)
            anim.setStartValue(start)
            anim.setEndValue(end)
            anim.setEasingCurve(QEasingCurve.Type.OutCubic)
            self.push_animations.append(anim)
            anim.start()

    def move_drag_gap(self, index):
        gap = self.drag_placeholder
        if gap is None:
            return
        current = self.layout.indexOf(gap)
        if current == index:
            return
        def change():
            if current >= 0:
                self.layout.removeWidget(gap)
            if index >= 0:
                self.layout.insertWidget(min(index, self.layout.count()), gap)
                gap.show()
            else:
                gap.hide()
        self.animate_icon_layout(change)
        self.drag_preview_index = index

    def start_shortcut_drag(self, button):
        self.dragging = True
        self.drag_button = button
        self.drag_original_order = self.shortcut_paths()
        self.hide_deadline = None
        self.icon_tooltip.hide()
        self.internal_drop = False
        index = self.layout.indexOf(button)
        gap = QWidget(self.content)
        gap.setFixedSize(BUTTON_SIZE, BUTTON_SIZE)
        gap.setStyleSheet("background: transparent; border: none;")
        self.drag_placeholder = gap
        self.drag_preview_index = index
        self.layout.removeWidget(button)
        button.hide()
        self.layout.insertWidget(index, gap)
        self.layout.activate()

        drag = QDrag(button)
        mime = QMimeData()
        mime.setData(MIME_DOCK, str(button.path).encode("utf-8"))
        mime.setUrls([QUrl.fromLocalFile(str(button.path))])
        drag.setMimeData(mime)
        drag.setPixmap(button.icon().pixmap(QSize(ICON_SIZE, ICON_SIZE)))
        drag.setHotSpot(QPoint(ICON_SIZE // 2, ICON_SIZE // 2))
        # Override Qt/Windows drag-operation badge. The floating icon itself
        # is already drawn by setPixmap; no additional '+' cursor is needed.
        clean_cursor = QPixmap(1, 1)
        clean_cursor.fill(Qt.GlobalColor.transparent)
        drag.setDragCursor(clean_cursor, Qt.DropAction.MoveAction)
        drag.setDragCursor(clean_cursor, Qt.DropAction.CopyAction)
        self.drag_preview_timer = QTimer(self)
        self.drag_preview_timer.timeout.connect(self.update_drag_preview)
        self.drag_preview_timer.start(16)
        result = drag.exec(Qt.DropAction.CopyAction | Qt.DropAction.MoveAction,
                           Qt.DropAction.MoveAction)
        self.drag_preview_timer.stop()
        self.drag_preview_timer.deleteLater()
        self.drag_preview_timer = None
        for anim in self.push_animations:
            anim.stop()
        self.push_animations.clear()
        gap_index = self.layout.indexOf(gap)
        if gap_index >= 0:
            self.layout.removeWidget(gap)
        gap.hide()
        gap.deleteLater()
        self.drag_placeholder = None
        was_internal = self.internal_drop
        self.internal_drop = False
        self.dragging = False
        self.drag_button = None
        self.hide_deadline = None
        if was_internal:
            self.layout.insertWidget(gap_index if gap_index >= 0 else index, button)
            button.show()
            self.layout.activate()
            self.save_order(self.shortcut_paths_from_layout())
        else:
            # A cancelled drag restores the old order; a successful Explorer
            # drop removes the original shortcut from the dock.
            if result in (Qt.DropAction.CopyAction, Qt.DropAction.MoveAction):
                if button.path.exists():
                    self.remove_shortcut(button.path)
            else:
                self.save_order(self.drag_original_order)
                self.load_shortcuts()
        self.drag_original_order = None
        self.drag_preview_index = None

    def shortcut_paths_from_layout(self):
        return [self.layout.itemAt(i).widget().path
                for i in range(self.layout.count())
                if isinstance(self.layout.itemAt(i).widget(), ShortcutButton)]

    def preview_insertion(self, target_button, before):
        if not self.dragging or self.drag_placeholder is None:
            return
        if target_button is None or target_button is self.drag_button:
            return
        target = self.layout.indexOf(target_button)
        current = self.layout.indexOf(self.drag_placeholder)
        if target < 0:
            return
        index = target + (0 if before else 1)
        if current >= 0 and current < index:
            index -= 1
        self.move_drag_gap(index)

    def update_drag_preview(self):
        if not self.dragging or self.drag_placeholder is None:
            return
        global_pos = QCursor.pos()
        local = self.content.mapFromGlobal(global_pos)
        if not self.content.rect().contains(local):
            self.move_drag_gap(-1)
            return
        horizontal = self.position in ("Bottom", "Top")
        coord = local.x() if horizontal else local.y()
        buttons = [self.layout.itemAt(i).widget() for i in range(self.layout.count())
                   if isinstance(self.layout.itemAt(i).widget(), ShortcutButton)]
        if not buttons:
            self.move_drag_gap(0)
            return
        index = 0
        for widget in buttons:
            center = (widget.x() + widget.width() / 2 if horizontal
                      else widget.y() + widget.height() / 2)
            if coord >= center:
                index += 1
        self.move_drag_gap(index)

    def external_drag_preview(self, event, target_button=None):
        """Show an animated insertion gap for files dragged from Explorer."""
        if self.dragging:
            return
        self.external_drag_active = True
        self.hide_deadline = None
        if self.drag_placeholder is None:
            gap = QWidget(self.content)
            gap.setFixedSize(BUTTON_SIZE, BUTTON_SIZE)
            gap.setStyleSheet("background: transparent; border: none;")
            self.drag_placeholder = gap
        # Drag events may be delivered to a different child as the insertion
        # placeholder moves. The drag source is often an Explorer HWND (or a
        # widget outside our hierarchy), so mapFrom(source, ...) is invalid.
        # Native cursor coordinates are stable across child/layout changes.
        pos = self.content.mapFromGlobal(QCursor.pos())
        horizontal = self.position in ("Bottom", "Top")
        coord = pos.x() if horizontal else pos.y()
        buttons = [self.layout.itemAt(i).widget() for i in range(self.layout.count())
                   if isinstance(self.layout.itemAt(i).widget(), ShortcutButton)]
        index = 0
        for button in buttons:
            midpoint = (button.x() + button.width() / 2 if horizontal
                        else button.y() + button.height() / 2)
            if coord >= midpoint:
                index += 1
        self.move_drag_gap(index)

    def clear_external_drag_preview(self):
        if not self.external_drag_active:
            return
        self.external_drag_active = False
        gap = self.drag_placeholder
        self.drag_placeholder = None
        self.drag_preview_index = None
        if gap is None:
            return
        for anim in self.push_animations:
            anim.stop()
        self.push_animations.clear()
        def change():
            self.layout.removeWidget(gap)
            gap.hide()
        self.animate_icon_layout(change)
        gap.deleteLater()

    def dragLeaveEvent(self, event):
        # Changing the placeholder can cause a transient child drag-leave.
        # Only clear the preview after verifying the cursor really left.
        QTimer.singleShot(0, self.clear_external_preview_if_outside)
        event.accept()

    def clear_external_preview_if_outside(self):
        if (self.external_drag_active and
                not self.rect().contains(self.mapFromGlobal(QCursor.pos()))):
            self.clear_external_drag_preview()

    def handle_drag_hover(self, event, target_button=None):
        if event.mimeData().hasFormat(MIME_DOCK):
            self.update_drag_preview()
            event.setDropAction(Qt.DropAction.MoveAction)
            event.accept()
        elif event.mimeData().hasUrls():
            self.external_drag_preview(event, target_button)
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()

    def eventFilter(self, obj, event):
        if isinstance(obj, ShortcutButton):
            if event.type() == QEvent.Type.Enter:
                self.update_icon_tooltip(obj)
            elif event.type() == QEvent.Type.Leave:
                if self.tooltip_button is obj:
                    self.clear_icon_tooltip()
            if event.type() == QEvent.Type.ToolTip:
                # Our explicit hover timer handles tooltips, including native z-order.
                return True
            if event.type() == QEvent.Type.Wheel:
                self.scroll_icons(event.angleDelta().y())
                event.accept()
                return True
        if event.type() == QEvent.Type.Wheel and (obj is self.content or
                obj is self.scroll_area.viewport()):
            self.scroll_icons(event.angleDelta().y())
            event.accept()
            return True
        if obj in (self.content, self.scroll_area.viewport()) and event.type() in (
                QEvent.Type.DragEnter, QEvent.Type.DragMove, QEvent.Type.DragLeave, QEvent.Type.Drop):
            if event.type() == QEvent.Type.Drop:
                self.handle_drop(event)
            elif event.type() == QEvent.Type.DragLeave:
                QTimer.singleShot(0, self.clear_external_preview_if_outside)
                event.accept()
            else:
                self.handle_drag_hover(event)
            return True
        return super().eventFilter(obj, event)

    def scroll_icons(self, delta):
        horizontal = self.position in ("Bottom", "Top")
        bar = (self.scroll_area.horizontalScrollBar() if horizontal
               else self.scroll_area.verticalScrollBar())
        if bar.maximum() == 0 or delta == 0:
            return
        direction = -1 if delta > 0 else 1
        if self.reverse_scroll:
            direction *= -1
        # One wheel notch = one icon pitch (icon + its spacing).
        # Snap to the icon grid instead of accumulating arbitrary pixel offsets.
        pitch = BUTTON_SIZE + self.icon_spacing
        steps = max(1, round(abs(delta) / 120))
        current = bar.value()
        # Snap instantly to the next whole icon.
        if direction > 0:
            index = (current // pitch) + steps
        else:
            index = -((-current) // pitch) - steps
        target = max(bar.minimum(), min(bar.maximum(), index * pitch))
        bar.setValue(target)
        self.hide_deadline = None

    def change_spacing(self, value):
        self.icon_spacing = value
        self.settings.setValue("icon_spacing", value)
        self.spacing_label.setText(f"Icon spacing: {value} px")
        self.layout.setSpacing(value)
        self.update_orientation()

    def change_reverse_scroll(self, checked):
        self.reverse_scroll = checked
        self.settings.setValue("reverse_scroll", checked)

    def dragEnterEvent(self, event):
        self.handle_drag_hover(event)

    def dragMoveEvent(self, event):
        self.handle_drag_hover(event)

    def dropEvent(self, event):
        self.handle_drop(event)

    def handle_drop(self, event, target_button=None):
        target = target_button.path if target_button else None
        before = True
        if target_button:
            p = event.position()
            before = (p.x() < target_button.width() / 2 if
                      self.position in ("Bottom", "Top") else
                      p.y() < target_button.height() / 2)
        if event.mimeData().hasFormat(MIME_DOCK):
            try:
                source = Path(bytes(event.mimeData().data(MIME_DOCK)).decode("utf-8"))
            except (UnicodeError, ValueError):
                event.ignore()
                return
            self.internal_drop = True
            # Live preview already rearranged the widgets; commit on drop.
            event.setDropAction(Qt.DropAction.MoveAction)
            event.accept()
            return
        if event.mimeData().hasUrls():
            # Use the live gap, not the icon under the cursor, to commit the
            # exact position that was previewed throughout the external drag.
            insertion = self.layout.indexOf(self.drag_placeholder) if self.external_drag_active else -1
            existing = self.shortcut_paths_from_layout()
            if insertion >= 0:
                target = existing[insertion] if insertion < len(existing) else None
                before = True
            self.clear_external_drag_preview()
            sources = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
            if self.add_sources(sources, target, before):
                event.setDropAction(Qt.DropAction.CopyAction)
                event.accept()
                return
        event.ignore()

    def launch_shortcut(self, shortcut):
        try:
            os.startfile(str(shortcut))
        except OSError as error:
            QMessageBox.warning(
                self,
                "Shortcut Error",
                f"Unable to launch:\n"
                f"{shortcut}\n\n{error}"
            )

    def update_orientation(self):
        horizontal = self.position in ("Bottom", "Top")
        self.layout.setDirection(QBoxLayout.Direction.LeftToRight if horizontal
                                 else QBoxLayout.Direction.TopToBottom)
        count = max(1, self.layout.count())
        length = MARGIN * 2 + count * BUTTON_SIZE + (count - 1) * self.icon_spacing
        thickness = MARGIN * 2 + BUTTON_SIZE
        screen = self.get_target_screen()
        if screen is not None:
            available = screen.availableGeometry()
            max_length = (available.width() if horizontal else available.height())
        else:
            max_length = length
        visible_length = min(length, max(thickness, max_length))
        width, height = ((visible_length, thickness) if horizontal
                         else (thickness, visible_length))
        old_horizontal = self.scroll_area.horizontalScrollBar().value()
        old_vertical = self.scroll_area.verticalScrollBar().value()
        self.setFixedSize(width, height)
        self.panel.setGeometry(0, 0, width, height)
        self.scroll_area.setGeometry(0, 0, width, height)
        self.content.setFixedSize(length if horizontal else thickness,
                                  thickness if horizontal else length)
        self.layout.activate()
        if horizontal:
            self.scroll_area.horizontalScrollBar().setValue(old_horizontal)
        else:
            self.scroll_area.verticalScrollBar().setValue(old_vertical)
        self.position_dock()

    def get_target_screen(self):
        if self.primary_only:
            return QApplication.primaryScreen()

        screen = QApplication.screenAt(
            QCursor.pos()
        )

        return screen or QApplication.primaryScreen()

    def position_dock(self, screen=None):
        screen = screen or self.get_target_screen()

        if screen is None:
            return

        self.current_screen = screen

        # availableGeometry excludes the Windows
        # taskbar's reserved area.
        geometry = screen.availableGeometry()

        if self.position == "Bottom":
            x = (
                geometry.x()
                + (geometry.width() - self.width()) // 2
            )
            y = geometry.bottom() - self.height() + 1 - self.offset

        elif self.position == "Top":
            x = (
                geometry.x()
                + (geometry.width() - self.width()) // 2
            )
            y = geometry.top() + self.offset

        elif self.position == "Left":
            x = geometry.left() + self.offset
            y = (
                geometry.y()
                + (geometry.height() - self.height()) // 2
            )

        else:  # Right
            x = geometry.right() - self.width() + 1 - self.offset
            y = (
                geometry.y()
                + (geometry.height() - self.height()) // 2
            )

        self.visible_position = QPoint(x, y)
        if not self.animating and (self.isVisible() or not self.auto_hide):
            self.move(self.visible_position)

    def taskbar_on_dock_edge(self, screen):
        """True only when Windows reserves space on our selected edge."""
        full = screen.geometry()
        work = screen.availableGeometry()
        return {
            "Bottom": work.bottom() < full.bottom(),
            "Top": work.top() > full.top(),
            "Left": work.left() > full.left(),
            "Right": work.right() < full.right(),
        }[self.position]

    def keep_above_windows(self):
        """Keep the dock in the topmost band, directly BEHIND the taskbar."""
        flags = 0x0001 | 0x0002 | 0x0010  # NOSIZE | NOMOVE | NOACTIVATE
        user32.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
        user32.FindWindowW.restype = wintypes.HWND
        user32.FindWindowExW.argtypes = [wintypes.HWND, wintypes.HWND,
                                         wintypes.LPCWSTR, wintypes.LPCWSTR]
        user32.FindWindowExW.restype = wintypes.HWND
        dock = int(self.winId())
        # First make the dock topmost, so regular windows cannot obscure it.
        user32.SetWindowPos(dock, -1, 0, 0, 0, 0, flags)
        taskbars = []
        primary = user32.FindWindowW("Shell_TrayWnd", None)
        if primary:
            taskbars.append(primary)
        previous = None
        while True:
            secondary = user32.FindWindowExW(None, previous,
                                              "Shell_SecondaryTrayWnd", None)
            if not secondary:
                break
            taskbars.append(secondary)
            previous = secondary
        # Windows' taskbars must be in the topmost band as well. Place the
        # dock immediately AFTER the taskbars, rather than raising the dock
        # above them or dropping it below ordinary application windows.
        for taskbar in taskbars:
            user32.SetWindowPos(taskbar, -1, 0, 0, 0, 0, flags)
        if taskbars:
            user32.SetWindowPos(dock, taskbars[-1], 0, 0, 0, 0, flags)

    def keep_below_taskbar(self, screen):
        """Maintain the same z-order during sliding as when fully shown."""
        self.keep_above_windows()

    def hidden_position(self, screen):
        # Always start beyond the OUTER monitor edge, even when the
        # Windows taskbar occupies this edge. This matters especially
        # with transparent taskbars: no stationary dock should be visible.
        rect = screen.geometry()
        p = self.visible_position
        if self.position == "Bottom":
            return QPoint(p.x(), rect.bottom() + 1)
        if self.position == "Top":
            return QPoint(p.x(), rect.top() - self.height())
        if self.position == "Left":
            return QPoint(rect.left() - self.width(), p.y())
        return QPoint(rect.right() + 1, p.y())

    def animate_dock(self, show):
        screen = self.get_target_screen()
        if screen is None:
            return

        self.animation.stop()
        self.animating = False
        self.position_dock(screen)
        target = self.visible_position if show else self.hidden_position(screen)
        start = self.pos() if self.isVisible() else self.hidden_position(screen)

        if show:
            # Qt/Windows can restore a top-level window's previous
            # position on show(). Reposition it AFTER show(), while
            # transparent, before allowing its first visible frame.
            self.setWindowOpacity(0.0)
            self.show()
            self.move(start)
            self.keep_below_taskbar(screen)

        self.animation_hiding = not show
        self.animating = True
        self.animation.setDuration(230)
        self.animation.setKeyValues([(0.0, start), (1.0, target)])

        # Traverse the taskbar-covered portion almost immediately, so the
        # dock begins emerging without a perceptible wait. Preserve the
        # original outer-screen endpoints and total 230 ms duration.
        if self.taskbar_on_dock_edge(screen):
            work = screen.availableGeometry()
            if self.position == "Bottom":
                boundary = QPoint(self.visible_position.x(), work.bottom() + 1)
            elif self.position == "Top":
                boundary = QPoint(self.visible_position.x(), work.top() - self.height())
            elif self.position == "Left":
                boundary = QPoint(work.left() - self.width(), self.visible_position.y())
            else:
                boundary = QPoint(work.right() + 1, self.visible_position.y())
            if show:
                self.animation.setKeyValueAt(0.06, boundary)
            else:
                self.animation.setKeyValueAt(0.94, boundary)

        if show:
            # Start the slide before exposing the window. Keeping it
            # transparent during show/z-order adjustment avoids a flash
            # of the full dock behind transparent taskbars.
            self.animation.start()
            self.setWindowOpacity(1.0)
        else:
            self.keep_above_windows()
            self.animation.start()

    def animation_finished(self):
        self.animating = False
        if self.animation_hiding and self.auto_hide:
            self.hide()
        elif not self.animation_hiding:
            self.move(self.visible_position)
            self.keep_above_windows()

    def cursor_on_taskbar(self, cursor, screen):
        # Windows reserves taskbar space in availableGeometry().
        # Compare it with the full monitor rectangle.
        full = screen.geometry()
        available = screen.availableGeometry()
        return full.contains(cursor) and not available.contains(cursor)

    def in_activation_zone(self, cursor, screen):
        """Check the actual outer screen edge in physical Windows pixels."""
        native = native_cursor_and_monitor()
        if native is None:
            # Fallback if the native API is temporarily unavailable.
            area = screen.geometry()
            if not area.contains(cursor):
                return False
            edge = 4
            if self.position == "Bottom":
                return cursor.y() >= area.bottom() - edge + 1
            if self.position == "Top":
                return cursor.y() <= area.top() + edge - 1
            if self.position == "Left":
                return cursor.x() <= area.left() + edge - 1
            return cursor.x() >= area.right() - edge + 1

        point, bounds = native
        if self.primary_only and QApplication.screenAt(cursor) is not screen:
            return False
        # Four logical pixels, converted to physical pixels for this screen.
        edge = max(4, round(4 * screen.devicePixelRatio()))
        if self.position == "Bottom":
            return point.y >= bounds.bottom - edge
        if self.position == "Top":
            return point.y < bounds.top + edge
        if self.position == "Left":
            return point.x < bounds.left + edge
        return point.x >= bounds.right - edge

    def update_monitor(self):
        screen = self.get_target_screen()
        if screen is None:
            return

        cursor = QCursor.pos()
        if not self.auto_hide:
            if not self.underMouse():
                self.position_dock(screen)
            return

        zone = self.in_activation_zone(cursor, screen)
        if self.dragging or self.icon_menu_open or self.tray_menu_open or self.rename_dialog_open:
            self.hide_deadline = None
            return
        if self.animating:
            return

        if not self.isVisible():
            # The cursor may already be at the edge when the dock hides.
            # Do not require an edge-entry transition to reveal it.
            if zone:
                if self.appearance_since is None:
                    self.appearance_since = time.monotonic()
                if (time.monotonic() - self.appearance_since) * 1000 >= self.appearance_delay:
                    self.animate_dock(True)
                    self.hide_deadline = None
                    self.activation_armed = True
                    self.appearance_since = None
            else:
                self.appearance_since = None
            self.last_activation_zone = zone
            return

        self.last_activation_zone = zone
        if (zone or self.underMouse() or self.geometry().contains(cursor)
                or self.cursor_on_taskbar(cursor, screen) or self.icon_tooltip.isVisible()):
            self.hide_deadline = None
            self.activation_armed = False
            return

        # After revealing, allow the cursor to travel from the edge strip
        # to the dock even if the offset leaves a gap between them.
        if self.activation_armed:
            self.activation_armed = False
            self.hide_deadline = None
            return

        now = time.monotonic()
        if self.hide_deadline is None:
            self.hide_deadline = now + self.hide_delay / 1000.0
        elif now >= self.hide_deadline:
            self.animate_dock(False)
            self.hide_deadline = None

    def change_auto_hide(self, enabled):
        self.auto_hide = enabled
        self.appearance_since = None
        self.settings.setValue("auto_hide", enabled)
        self.monitor_timer.setInterval(10 if enabled else 250)
        self.hide_deadline = None
        self.activation_armed = False
        self.last_activation_zone = False
        self.delay_slider.setEnabled(enabled)
        self.appearance_slider.setEnabled(enabled)
        self.animation.stop()
        self.animating = False
        if enabled:
            self.hide()
        else:
            self.show_dock()

    def change_appearance_delay(self, value):
        self.appearance_delay = value
        self.settings.setValue("appearance_delay", value)
        self.appearance_label.setText(f"Appearance delay: {value} ms")
        self.appearance_since = None

    def change_hide_delay(self, value):
        self.hide_delay = value
        self.settings.setValue("hide_delay", value)
        self.delay_label.setText(f"Hide delay: {value} ms")
        self.hide_deadline = None

    def change_position(self, position):
        self.position = position
        self.settings.setValue(
            "position", position
        )
        self.update_orientation()

    def change_monitor_mode(self, primary_only):
        self.primary_only = primary_only

        self.settings.setValue(
            "primary_only", primary_only
        )

        self.update_orientation()

    def change_offset(self, value):
        self.offset = value
        self.settings.setValue("offset", value)
        self.offset_label.setText(f"Offset: {value} px")
        self.position_dock()

    def make_tray_icon(self):
        # Draw a compact dock with three application tiles, not a computer icon.
        pixmap = QPixmap(64, 64)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor("#a6b3c1"), 2))
        painter.setBrush(QColor("#303c49"))
        painter.drawRoundedRect(4, 17, 56, 31, 10, 10)
        for x, color in [(11, "#5ba9ed"), (26, "#ecb55d"), (41, "#6cc8a1")]:
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(color))
            painter.drawRoundedRect(x, 24, 12, 15, 3, 3)
        painter.end()
        return QIcon(pixmap)

    def style_tray_menu(self, menu):
        style_icon_context_menu(menu)
        for action in menu.actions():
            if action.menu():
                self.style_tray_menu(action.menu())

    def style_tray_slider_labels(self):
        """Keep embedded QWidgetAction labels readable in the Windows theme."""
        foreground = "#f5f5f5" if windows_system_dark_theme() else "#202020"
        for label in (self.delay_label, self.appearance_label,
                      self.spacing_label, self.offset_label):
            label.setStyleSheet(f"QLabel {{ color: {foreground}; background: transparent; }}")

    def style_menu_slider(self, slider):
        dark = windows_system_dark_theme()
        track = "#606060" if dark else "#b8b8b8"
        handle = "#e8e8e8" if dark else "#444444"
        border = "#888888" if dark else "#666666"
        slider.setMinimumHeight(30)
        slider.setStyleSheet(f"""
            QSlider::groove:horizontal {{ height: 6px; border-radius: 3px; background: {track}; }}
            QSlider::handle:horizontal {{ background: {handle}; border: 2px solid {border};
                width: 22px; height: 22px; margin: -10px 0; border-radius: 12px; }}
            QSlider::handle:horizontal:hover {{ border: 2px solid #479be8; }}
        """)

    def show_about_dialog(self):
        """Simple About window matching Visual Studio Build Stopwatch."""
        dialog = QDialog(None)
        dialog.setWindowTitle("About Shortcuts Dock")
        dialog.setWindowIcon(self.make_tray_icon())
        dialog.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.WindowCloseButtonHint)
        dialog.setModal(True)
        dark = windows_system_dark_theme()
        bg = "#292929" if dark else "#f9f9f9"
        fg = "#f5f5f5" if dark else "#202020"
        border = "#555555" if dark else "#d0d0d0"
        button_bg = "#383838" if dark else "#ffffff"
        button_hover = "#454545" if dark else "#e8e8e8"
        dialog.setStyleSheet(f"""
            QDialog {{ background-color: {bg}; color: {fg}; }}
            QLabel {{ color: {fg}; background: transparent; }}
            QPushButton {{ background-color: {button_bg}; color: {fg};
                border: 1px solid {border}; border-radius: 4px;
                padding: 6px 18px; min-width: 78px; }}
            QPushButton:hover {{ background-color: {button_hover}; }}
        """)
        outer = QVBoxLayout(dialog)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(9)
        name = QLabel("Shortcuts Dock")
        font = name.font()
        font.setPointSize(12)
        font.setBold(True)
        name.setFont(font)
        name.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(name)
        for text in (f"Version {application_version()}",
                     "Written by Gal Shemesh (PacketLauncher)"):
            label = QLabel(text)
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            outer.addWidget(label)
        github = QLabel('<a href="https://github.com/packetlauncher">github.com/packetlauncher</a>')
        github.setOpenExternalLinks(True)
        github.setAlignment(Qt.AlignmentFlag.AlignCenter)
        github.setStyleSheet("QLabel { color: #4da3ff; background: transparent; }")
        outer.addWidget(github)
        outer.addSpacing(5)
        ok = QPushButton("OK")
        ok.clicked.connect(dialog.accept)
        ok.setDefault(True)
        outer.addWidget(ok, 0, Qt.AlignmentFlag.AlignHCenter)
        dialog.adjustSize()
        set_native_dialog_theme(dialog)
        self.rename_dialog_open = True
        self.hide_deadline = None
        try:
            dialog.exec()
        finally:
            self.rename_dialog_open = False
            self.hide_deadline = None

    def create_tray(self):
        self.tray = QSystemTrayIcon(self)

        icon = self.make_tray_icon()

        self.tray.setIcon(icon)
        self.tray.setToolTip("Shortcuts Dock")

        menu = QMenu()

        # Position submenu
        position_menu = menu.addMenu("Position")

        position_group = QActionGroup(self)
        position_group.setExclusive(True)

        for position in POSITIONS:
            action = QAction(position, self)
            action.setCheckable(True)
            action.setChecked(
                position == self.position
            )

            action.triggered.connect(
                lambda checked=False, p=position:
                    self.change_position(p)
            )

            position_group.addAction(action)
            position_menu.addAction(action)

        # Monitor submenu
        monitor_menu = menu.addMenu("Monitor")

        monitor_group = QActionGroup(self)
        monitor_group.setExclusive(True)

        follow_action = QAction(
            "Follow Mouse Monitor", self
        )
        follow_action.setCheckable(True)
        follow_action.setChecked(
            not self.primary_only
        )
        follow_action.triggered.connect(
            lambda: self.change_monitor_mode(False)
        )

        primary_action = QAction(
            "Primary Monitor Only", self
        )
        primary_action.setCheckable(True)
        primary_action.setChecked(
            self.primary_only
        )
        primary_action.triggered.connect(
            lambda: self.change_monitor_mode(True)
        )

        monitor_group.addAction(follow_action)
        monitor_group.addAction(primary_action)

        monitor_menu.addAction(follow_action)
        monitor_menu.addAction(primary_action)

        menu.addSeparator()

        auto_hide_action = QAction("Auto Hide", self)
        auto_hide_action.setCheckable(True)
        auto_hide_action.setChecked(self.auto_hide)
        auto_hide_action.toggled.connect(self.change_auto_hide)
        menu.addAction(auto_hide_action)

        reverse_action = QAction("Reverse Icon Scrolling", self)
        reverse_action.setCheckable(True)
        reverse_action.setChecked(self.reverse_scroll)
        reverse_action.toggled.connect(self.change_reverse_scroll)
        menu.addAction(reverse_action)

        menu.addSeparator()

        delay_action = QWidgetAction(menu)
        delay_widget = QWidget(menu)
        delay_layout = QVBoxLayout(delay_widget)
        delay_layout.setContentsMargins(12, 6, 12, 8)
        delay_layout.setSpacing(4)
        self.delay_label = QLabel(f"Hide delay: {self.hide_delay} ms")
        delay_layout.addWidget(self.delay_label)
        self.delay_slider = QSlider(Qt.Orientation.Horizontal)
        self.delay_slider.setRange(0, 3000)
        self.delay_slider.setSingleStep(50)
        self.delay_slider.setPageStep(250)
        self.delay_slider.setValue(self.hide_delay)
        self.delay_slider.setMinimumWidth(190)
        self.delay_slider.setEnabled(self.auto_hide)
        self.delay_slider.valueChanged.connect(self.change_hide_delay)
        self.style_menu_slider(self.delay_slider)
        delay_layout.addWidget(self.delay_slider)
        delay_action.setDefaultWidget(delay_widget)
        menu.addAction(delay_action)

        appearance_action = QWidgetAction(menu)
        appearance_widget = QWidget(menu)
        appearance_layout = QVBoxLayout(appearance_widget)
        appearance_layout.setContentsMargins(12, 6, 12, 8)
        appearance_layout.setSpacing(4)
        self.appearance_label = QLabel(f"Appearance delay: {self.appearance_delay} ms")
        appearance_layout.addWidget(self.appearance_label)
        self.appearance_slider = QSlider(Qt.Orientation.Horizontal)
        self.appearance_slider.setRange(0, 3000)
        self.appearance_slider.setSingleStep(50)
        self.appearance_slider.setPageStep(250)
        self.appearance_slider.setValue(self.appearance_delay)
        self.appearance_slider.setMinimumWidth(190)
        self.appearance_slider.setEnabled(self.auto_hide)
        self.appearance_slider.valueChanged.connect(self.change_appearance_delay)
        self.style_menu_slider(self.appearance_slider)
        appearance_layout.addWidget(self.appearance_slider)
        appearance_action.setDefaultWidget(appearance_widget)
        menu.addAction(appearance_action)

        menu.addSeparator()

        spacing_action = QWidgetAction(menu)
        spacing_widget = QWidget(menu)
        spacing_layout = QVBoxLayout(spacing_widget)
        spacing_layout.setContentsMargins(12, 6, 12, 8)
        spacing_layout.setSpacing(4)
        self.spacing_label = QLabel(f"Icon spacing: {self.icon_spacing} px")
        spacing_layout.addWidget(self.spacing_label)
        spacing_slider = QSlider(Qt.Orientation.Horizontal)
        spacing_slider.setRange(0, 40)
        spacing_slider.setValue(self.icon_spacing)
        spacing_slider.setMinimumWidth(190)
        spacing_slider.valueChanged.connect(self.change_spacing)
        self.style_menu_slider(spacing_slider)
        spacing_layout.addWidget(spacing_slider)
        spacing_action.setDefaultWidget(spacing_widget)
        menu.addAction(spacing_action)

        # Embedded slider: stays interactive while the tray menu is open.
        offset_action = QWidgetAction(menu)
        offset_widget = QWidget(menu)
        offset_layout = QVBoxLayout(offset_widget)
        offset_layout.setContentsMargins(12, 6, 12, 8)
        offset_layout.setSpacing(4)

        self.offset_label = QLabel(f"Offset: {self.offset} px")
        offset_layout.addWidget(self.offset_label)

        offset_slider = QSlider(Qt.Orientation.Horizontal)
        offset_slider.setRange(0, 200)
        offset_slider.setSingleStep(1)
        offset_slider.setPageStep(10)
        offset_slider.setValue(self.offset)
        offset_slider.setMinimumWidth(190)
        offset_slider.valueChanged.connect(self.change_offset)
        self.style_menu_slider(offset_slider)
        offset_layout.addWidget(offset_slider)

        offset_action.setDefaultWidget(offset_widget)
        menu.addAction(offset_action)
        menu.addSeparator()

        about_action = QAction("About...", self)
        about_action.triggered.connect(self.show_about_dialog)
        menu.addAction(about_action)
        menu.addSeparator()

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(
            QApplication.quit
        )
        menu.addAction(exit_action)

        # Keep an already-visible dock on screen while the tray settings
        # menu (including its sliders and submenus) is open.
        menu.aboutToShow.connect(self.tray_menu_opened)
        menu.aboutToHide.connect(self.tray_menu_closed)
        self.style_tray_menu(menu)
        self.tray.setContextMenu(menu)

        # Re-read Windows theme when the user opens the menu after changing it.
        self.style_tray_slider_labels()
        menu.aboutToShow.connect(lambda: (self.style_tray_menu(menu), self.style_tray_slider_labels()))
        self.tray.activated.connect(
            self.tray_activated
        )
        self.tray.show()

    def tray_menu_opened(self):
        self.tray_menu_open = True
        self.hide_deadline = None
        self.appearance_since = None
        # Do not force a hidden dock to appear; only prevent an existing
        # visible dock from hiding while its settings are being adjusted.

    def tray_menu_closed(self):
        self.tray_menu_open = False
        self.hide_deadline = None
        self.activation_armed = True
        self.appearance_since = None

    def tray_activated(self, reason):
        if reason == (
            QSystemTrayIcon.ActivationReason.Trigger
        ):
            self.show_dock()

    def show_dock(self):
        self.hide_deadline = None
        self.activation_armed = True
        if self.auto_hide:
            self.animate_dock(True)
        else:
            self.position_dock()
            self.show()
            screen = self.get_target_screen()
            self.keep_above_windows()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    app.setQuitOnLastWindowClosed(False)

    # A named Windows mutex prevents launching another copy, even when the
    # application is started from a different working directory.
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR]
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL
    instance_mutex = kernel32.CreateMutexW(None, False, "Local\\ShortcutsDock_SingleInstance_8C0E93B7")
    if not instance_mutex:
        QMessageBox.warning(None, "Shortcuts Dock", "Unable to check whether Shortcuts Dock is already running.")
        sys.exit(1)
    if ctypes.get_last_error() == 183:  # ERROR_ALREADY_EXISTS
        kernel32.CloseHandle(instance_mutex)
        sys.exit(0)

    dock = ShortcutsDock()
    if not dock.auto_hide:
        dock.show()

    try:
        exit_code = app.exec()
    finally:
        kernel32.CloseHandle(instance_mutex)
    sys.exit(exit_code)
