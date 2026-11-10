import os
import sys
import time
import ctypes
import json
import re
import shutil
import subprocess
import struct
import hashlib
import uuid
import winreg
from ctypes import wintypes
import win32com.client
from pathlib import Path


from PySide6.QtCore import Qt, QSize, QFileInfo, QTimer, QSettings, QPropertyAnimation, QVariantAnimation, QEasingCurve, QPoint, QRect, QMimeData, QUrl, QEvent, QThread, Signal, QBuffer, QIODevice
from PySide6.QtGui import QAction, QActionGroup, QCursor, QDrag, QIcon, QPixmap, QPainter, QColor, QPen, QPainterPath, QImage, QPalette
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
    QListWidget,
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

SEPARATOR_SUFFIX = ".dockseparator"
ICON_SIZE = 24
BUTTON_SIZE = 40
ICON_SIZES = (16, 20, 24, 32, 40, 48, 64, 96, 128, 256)
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


def windows_accent_color():
    """Read the actual Windows accent selected in Personalization > Colors."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                r"Software\\Microsoft\\Windows\\DWM") as key:
            raw = int(winreg.QueryValueEx(key, "AccentColor")[0])
            return QColor(raw & 255, (raw >> 8) & 255, (raw >> 16) & 255)
    except (OSError, ValueError, TypeError):
        return QApplication.palette().color(QPalette.ColorRole.Highlight)


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


class IconSizeSlider(QSlider):
    """Continuous 1px dragging; wheel jumps to adjacent standard icon size."""

    def wheelEvent(self, event):
        delta = event.angleDelta().y() or event.angleDelta().x()
        if not delta:
            event.ignore()
            return
        direction = 1 if delta > 0 else -1
        # Use the nearest standard size in the requested direction, not the
        # nearest size overall. A partial wheel delta counts as one step.
        for _ in range(max(1, abs(delta) // 120)):
            current = self.value()
            if direction > 0:
                next_size = next((size for size in ICON_SIZES if size > current), ICON_SIZES[-1])
            else:
                next_size = next((size for size in reversed(ICON_SIZES) if size < current), ICON_SIZES[0])
            self.setValue(next_size)
        event.accept()


class WindowsAppsLoader(QThread):
    """Query Windows in a worker thread so the app picker opens immediately."""
    loaded = Signal(object, object, str)

    def run(self):
        try:
            script = (
                "[Console]::OutputEncoding=[Text.Encoding]::UTF8; "
                "$a=@(Get-StartApps | Select-Object Name,AppID); "
                "$p=@(Get-AppxPackage | Select-Object PackageFamilyName,InstallLocation); "
                "@{Apps=$a;Packages=$p} | ConvertTo-Json -Depth 5 -Compress"
            )
            proc = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive",
                 "-ExecutionPolicy", "Bypass", "-Command", script],
                capture_output=True, text=True, encoding="utf-8-sig",
                errors="replace", timeout=45,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            if proc.returncode:
                raise RuntimeError(proc.stderr.strip() or "Windows app query failed")
            data = json.loads(proc.stdout)
            apps = data.get("Apps") or []
            packages = data.get("Packages") or []
            if isinstance(apps, dict):
                apps = [apps]
            if isinstance(packages, dict):
                packages = [packages]
            result = sorted(
                [(str(a.get("Name", "")).strip(), str(a.get("AppID", "")).strip())
                 for a in apps if isinstance(a, dict) and a.get("Name") and a.get("AppID")],
                key=lambda item: item[0].casefold())
            roots = {}
            for package in packages:
                if not isinstance(package, dict):
                    continue
                family = package.get("PackageFamilyName")
                folder = package.get("InstallLocation")
                if family and folder:
                    roots.setdefault(family, []).append(folder)
            self.loaded.emit(result, roots, "")
        except Exception as error:
            self.loaded.emit([], {}, str(error))


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
        # Prefer an actual embedded ICO frame at its native dimensions.
        # For intermediate sizes, scale the nearest suitable native frame.
        icon = self.icon()
        if not icon.isNull():
            requested = self.iconSize().width()
            source, native_size = self._icon_source_pixmap(requested)
            if not source.isNull():
                target = self.rect()
                target.setLeft((self.width() - requested) // 2)
                target.setTop((self.height() - requested) // 2)
                target.setWidth(requested)
                target.setHeight(requested)
                # Preserve pixel-art crispness at exact integer enlargements.
                # Otherwise use smooth interpolation for noninteger scaling.
                integer_upscale = (native_size > 0 and requested >= native_size
                                   and requested % native_size == 0)
                if integer_upscale and requested > native_size:
                    cache_key = (source.cacheKey(), requested)
                    scaled = self._integer_cache.get(cache_key)
                    if scaled is None:
                        image = source.toImage()
                        image.setDevicePixelRatio(1.0)
                        image = image.scaled(
                            requested, requested,
                            Qt.AspectRatioMode.IgnoreAspectRatio,
                            Qt.TransformationMode.FastTransformation)
                        scaled = QPixmap.fromImage(image)
                        self._integer_cache[cache_key] = scaled
                    painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, False)
                    painter.drawPixmap(target.topLeft(), scaled)
                else:
                    painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform,
                                          not integer_upscale)
                    painter.drawPixmap(target, source, source.rect())
        painter.end()

    def drag_icon_pixmap(self):
        """Draw the exact same pixels as the dock icon, without the hover tile."""
        requested = self.iconSize().width()
        result = QPixmap(requested, requested)
        result.fill(Qt.GlobalColor.transparent)
        source, native = self._icon_source_pixmap(requested)
        if source.isNull():
            return result
        painter = QPainter(result)
        integer = native > 0 and requested > native and requested % native == 0
        if integer:
            key = (source.cacheKey(), requested)
            scaled = self._integer_cache.get(key)
            if scaled is None:
                image = source.toImage()
                image.setDevicePixelRatio(1.0)
                scaled = QPixmap.fromImage(image.scaled(
                    requested, requested, Qt.AspectRatioMode.IgnoreAspectRatio,
                    Qt.TransformationMode.FastTransformation))
                self._integer_cache[key] = scaled
            painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, False)
            painter.drawPixmap(QPoint(0, 0), scaled)
        else:
            painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
            painter.drawPixmap(QRect(0, 0, requested, requested), source, source.rect())
        painter.end()
        return result

    def _icon_source_pixmap(self, requested):
        icon = self.icon()
        key = icon.cacheKey()
        if getattr(self, "_source_icon_key", None) != key:
            self._source_icon_key = key
            self._source_pixmaps = {}
            self._trim_cache = {}
            self._integer_cache = {}
            self._native_sizes = sorted({sz.width() for sz in icon.availableSizes()
                                         if sz.width() == sz.height() and sz.width() > 0})

        sizes = self._native_sizes
        if sizes:
            # Use the exact frame when present; for in-between sizes prefer
            # downscaling a larger frame, except above the largest native frame.
            if requested in sizes:
                native = requested
            else:
                integer_sources = [n for n in sizes if n < requested and requested % n == 0]
                native = max(integer_sources) if integer_sources else next(
                    (n for n in sizes if n >= requested), sizes[-1])
        else:
            # Shell icon engines may not advertise their actual ICO sizes.
            # Request the current dimension directly instead of stretching a
            # single 256px image across every possible size.
            native = requested
        if native not in self._source_pixmaps:
            self._source_pixmaps[native] = icon.pixmap(QSize(native, native))
        pixmap = self._source_pixmaps[native]
        actual = round(pixmap.width() / max(1.0, pixmap.devicePixelRatio()))
        # The Windows shell icon engine sometimes returns a 16px bitmap even
        # when asked for 32px or more. In that case prefer the largest cached
        # *real* bitmap, rather than repeatedly enlarging a tiny fallback.
        if not sizes and actual < requested:
            for probe in (256, 128, 96, 64, 48, 32):
                if probe <= actual:
                    break
                if probe not in self._source_pixmaps:
                    self._source_pixmaps[probe] = icon.pixmap(QSize(probe, probe))
                candidate = self._source_pixmaps[probe]
                candidate_size = round(candidate.width() / max(1.0, candidate.devicePixelRatio()))
                if candidate_size > actual:
                    pixmap, actual = candidate, candidate_size
                    break
        # Some Windows icon engines return a 256px canvas containing only a
        # 16/32px picture. Crop the empty padding rather than scaling the canvas.
        # Keep this cached: scanning pixels during every paint causes flicker.
        cache_key = (native, pixmap.cacheKey())
        if cache_key not in self._trim_cache:
            trimmed_pixmap = pixmap
            if not pixmap.isNull() and pixmap.width() >= 48:
                image = pixmap.toImage().convertToFormat(QImage.Format.Format_ARGB32)
                left, top = image.width(), image.height()
                right = bottom = -1
                for y in range(image.height()):
                    for x in range(image.width()):
                        if (image.pixel(x, y) >> 24) & 255 >= 16:
                            left = min(left, x)
                            top = min(top, y)
                            right = max(right, x)
                            bottom = max(bottom, y)
                if right >= left and bottom >= top:
                    width, height = right - left + 1, bottom - top + 1
                    if max(width, height) <= max(image.width(), image.height()) * .62:
                        # Center the actual graphic in a square transparent canvas.
                        side = max(width, height)
                        crop = pixmap.copy(left, top, width, height)
                        square = QPixmap(side, side)
                        square.fill(Qt.GlobalColor.transparent)
                        cp = QPainter(square)
                        cp.drawPixmap((side - width)//2, (side - height)//2, crop)
                        cp.end()
                        trimmed_pixmap = square
            self._trim_cache[cache_key] = trimmed_pixmap
        pixmap = self._trim_cache[cache_key]
        actual = round(pixmap.width() / max(1.0, pixmap.devicePixelRatio()))
        return pixmap, (native if sizes and pixmap is self._source_pixmaps.get(native) else actual)

    def prepare_icon_frames(self):
        """Warm icon and integer-size caches outside resize/paint callbacks."""
        if self.icon().isNull():
            return
        # Accessing the native representations once also caches expensive
        # transparent-border detection for each frame.
        self._icon_source_pixmap(self.iconSize().width())
        native_sizes = list(self._native_sizes)
        for native in native_sizes:
            self._icon_source_pixmap(native)
        for requested in ICON_SIZES:
            source, native = self._icon_source_pixmap(requested)
            if source.isNull() or native <= 0 or requested <= native or requested % native:
                continue
            key = (source.cacheKey(), requested)
            if key not in self._integer_cache:
                image = source.toImage()
                image.setDevicePixelRatio(1.0)
                image = image.scaled(requested, requested,
                    Qt.AspectRatioMode.IgnoreAspectRatio,
                    Qt.TransformationMode.FastTransformation)
                self._integer_cache[key] = QPixmap.fromImage(image)

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
        location_action = menu.addAction("Open file location")
        rename_action = menu.addAction("Rename")
        windows_apps_action = menu.addAction("Add Windows app...")
        remove_action = menu.addAction("Remove from Dock")
        menu.addSeparator()
        add_separator_action = menu.addAction("Add separator")
        menu.addSeparator()
        exit_action = menu.addAction("Exit dock")
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
        if selected in (rename_action, add_separator_action):
            self.dock.allow_context_return()
        if selected == open_action:
            self.dock.launch_shortcut(self.path)
        elif selected == location_action:
            self.dock.open_file_location(self.path)
        elif selected == rename_action:
            self.dock.rename_shortcut(self.path)
        elif selected == windows_apps_action:
            self.dock.browse_windows_apps()
        elif selected == remove_action:
            self.dock.remove_shortcut(self.path)
        elif selected == add_separator_action:
            self.dock.add_separator(self.path)
        elif selected == exit_action:
            QApplication.instance().quit()

    def dragEnterEvent(self, event):
        self.dock.handle_drag_hover(event, self)

    def dragMoveEvent(self, event):
        self.dock.handle_drag_hover(event, self)

    def dragLeaveEvent(self, event):
        event.accept()

    def dropEvent(self, event):
        self.dock.handle_drop(event, self)


class SeparatorButton(ShortcutButton):
    """A movable, persistent visual spacer; never exported to Explorer."""
    def __init__(self, dock, path):
        super().__init__(dock, path)
        self.setCursor(Qt.CursorShape.SizeAllCursor)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        dark = windows_system_dark_theme()
        if self.underMouse() or self.isDown():
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(255,255,255,28) if dark else QColor(0,0,0,18))
            painter.drawRoundedRect(self.rect().adjusted(2,2,-2,-2), 5, 5)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(255,255,255,95) if dark else QColor(0,0,0,85))
        if self.dock.position in ("Top", "Bottom"):
            painter.drawRoundedRect((self.width()-2)//2, self.height()//4, 2, self.height()//2, 1, 1)
        else:
            painter.drawRoundedRect(self.width()//4, (self.height()-2)//2, self.width()//2, 2, 1, 1)
        painter.end()

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        style_icon_context_menu(menu)
        add = menu.addAction("Add separator")
        remove = menu.addAction("Remove separator")
        windows_apps_action = menu.addAction("Add Windows app...")
        menu.addSeparator()
        exit_action = menu.addAction("Exit dock")
        self.dock.icon_menu_open = True
        self.dock.hide_deadline = None
        try:
            pos = self.mapToGlobal(event.pos())
            selected = menu.exec(pos)
        finally:
            self.dock.icon_menu_open = False
            self.dock.hide_deadline = None
        if selected in (add, remove):
            self.dock.allow_context_return()
        if selected == add:
            self.dock.add_separator(self.path)
        elif selected == remove:
            self.dock.remove_separator(self.path)
        elif selected == windows_apps_action:
            self.dock.browse_windows_apps()
        elif selected == exit_action:
            QApplication.instance().quit()

    def mouseMoveEvent(self, event):
        # The same internal gap animation, but never an Explorer drag.
        super().mouseMoveEvent(event)


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
        saved_icon_size = int(self.settings.value("icon_size", ICON_SIZE))
        self.icon_size = max(ICON_SIZES[0], min(ICON_SIZES[-1], saved_icon_size))
        self.button_size = self.icon_size + 16
        self.separator_distance = max(0, min(100, int(self.settings.value("separator_distance", 12))))
        self._rendered_icon_size = self.icon_size
        self._center_overflow_on_resize = True
        self.reverse_scroll = self.settings.value("reverse_scroll", False, type=bool)

        self.auto_hide = self.settings.value("auto_hide", False, type=bool)
        self.hide_delay = max(0, min(3000, int(self.settings.value("hide_delay", 500))))
        self.hide_deadline = None
        self.context_grace_until = 0.0
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
        # A single interruptible animation interpolates the rendered size.
        # The slider stores its exact requested value independently.
        self.size_animation = QVariantAnimation(self)
        self.size_animation.setDuration(115)
        self.size_animation.setEasingCurve(QEasingCurve.Type.OutCubic)
        self.size_animation.valueChanged.connect(self.render_icon_size)
        self.size_animation.finished.connect(lambda: self.render_icon_size(self.icon_size))
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
            arguments = str(link.Arguments or "")
            if "shell:AppsFolder" in arguments and not (
                    icon_location and Path(
                        os.path.expandvars(icon_location.rsplit(",", 1)[0].strip('"'))
                    ).is_file()):
                app_id = arguments.split("shell:AppsFolder", 1)[1].lstrip("\\\\/")
                if app_id:
                    return self.windows_app_icon(app_id, shortcut.stem)

            # Use the shortcut's custom icon file when available.
            if icon_location:
                icon_path = os.path.expandvars(
                    icon_location.rsplit(",", 1)[0].strip('"')
                )
                if os.path.isfile(icon_path):
                    # QFileIconProvider returns a Windows shell icon, which may
                    # substitute a 16px representation for 32px at large sizes.
                    # Read .ico files directly so Qt can select their actual
                    # embedded 16/32/48/256px frames instead.
                    if Path(icon_path).suffix.lower() == ".ico":
                        icon = QIcon(icon_path)
                        if not icon.isNull():
                            return icon
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
                    if Path(target_path).suffix.lower() == ".ico":
                        icon = QIcon(target_path)
                        if not icon.isNull():
                            return icon
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
                 if p.is_file() and p.suffix.lower() in (".lnk", SEPARATOR_SUFFIX)]
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
                   if p.is_file() and p.suffix.lower() in (".lnk", SEPARATOR_SUFFIX)}
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

    def allow_context_return(self, seconds=1.5):
        """Keep the dock visible briefly after a context-menu operation."""
        self.context_grace_until = max(self.context_grace_until, time.monotonic() + seconds)
        self.hide_deadline = None

    def add_separator(self, after=None):
        """Create a real marker file so separators survive restarts/reordering."""
        path = SHORTCUTS_DIR / ("separator-" + uuid.uuid4().hex + SEPARATOR_SUFFIX)
        path.touch()
        order = [p for p in self.shortcut_paths() if p != path]
        if after in order:
            order.insert(order.index(after) + 1, path)
        else:
            order.append(path)
        self.save_order(order)
        self.load_shortcuts()
        self.allow_context_return()

    def remove_separator(self, path):
        if path.suffix.lower() != SEPARATOR_SUFFIX:
            return
        try:
            path.unlink()
        except OSError as error:
            QMessageBox.warning(self, "Remove Separator", str(error))
            return
        self.save_order([p for p in self.shortcut_paths() if p != path])
        self.load_shortcuts()
        self._center_overflow_on_resize = True
        self.update_orientation()
        self.allow_context_return()

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
            button.setFixedSize(self.button_size, self.button_size)
            button.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
            self.layout.addWidget(button)

        for shortcut in shortcuts:
            is_separator = shortcut.suffix.lower() == SEPARATOR_SUFFIX
            button = SeparatorButton(self, shortcut) if is_separator else ShortcutButton(self, shortcut)
            button.installEventFilter(self)
            if not is_separator:
                button.setIcon(self.get_shortcut_icon(shortcut))
            button.setIconSize(QSize(self.icon_size, self.icon_size))
            if is_separator:
                self.size_separator(button)
            else:
                button.setFixedSize(self.button_size, self.button_size)
            button.setToolTip("")
            if not is_separator:
                button.clicked.connect(lambda checked=False, path=shortcut:
                                       self.launch_shortcut(path))
            self.layout.addWidget(button)
            if isinstance(button, ShortcutButton):
                QTimer.singleShot(0, button.prepare_icon_frames)
        self._center_overflow_on_resize = True
        self.update_orientation()
        self._reset_scroll_for_fitting_content()
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
            self.allow_context_return()
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
        except OSError as error:
            QMessageBox.warning(self, "Remove Shortcut", str(error))
            return
        order = [p for p in self.shortcut_paths() if p != path]
        self.save_order(order)
        self.known_shortcut_names = {p.name.casefold() for p in order
                                     if p.suffix.lower() == ".lnk"}
        removed = next(
            (self.layout.itemAt(i).widget() for i in range(self.layout.count())
             if isinstance(self.layout.itemAt(i).widget(), ShortcutButton)
             and self.layout.itemAt(i).widget().path == path), None)
        if removed is None:
            self.load_shortcuts()
            self.allow_context_return()
            return

        # Build the final layout geometry while the OLD viewport is still
        # wide enough to contain the final icon. Qt must not paint a resized
        # viewport against the stale layout/scroll offset.
        self.setUpdatesEnabled(False)
        self.scroll_area.setUpdatesEnabled(False)
        self.content.setUpdatesEnabled(False)
        try:
            self.clear_icon_tooltip()
            for animation in self.push_animations:
                animation.stop()
            self.push_animations.clear()
            self.layout.removeWidget(removed)
            removed.hide()
            removed.deleteLater()
            if not order:
                self.load_shortcuts()
            else:
                horizontal = self.position in ("Bottom", "Top")
                count = self.layout.count()
                thickness = MARGIN * 2 + self.button_size
                length = MARGIN * 2 + max(0, count - 1) * self.icon_spacing
                for i in range(count):
                    widget = self.layout.itemAt(i).widget()
                    if isinstance(widget, SeparatorButton):
                        self.size_separator(widget)
                    length += ((widget.width() if horizontal else widget.height())
                               if widget else self.button_size)

                # Reset the old scroll position BEFORE changing the viewport.
                self.scroll_area.horizontalScrollBar().setValue(0)
                self.scroll_area.verticalScrollBar().setValue(0)
                self.content.move(0, 0)
                self.content.setFixedSize(
                    length if horizontal else thickness,
                    thickness if horizontal else length)
                self.layout.invalidate()
                self.layout.activate()
                self.layout.setGeometry(self.content.rect())

                screen = self.get_target_screen()
                if screen:
                    available = screen.availableGeometry()
                    limit = available.width() if horizontal else available.height()
                else:
                    limit = length
                visible = min(length, max(thickness, limit))
                width, height = ((visible, thickness) if horizontal
                                 else (thickness, visible))
                self.set_dock_geometry(width, height, screen)
                self.panel.setGeometry(0, 0, width, height)
                self.scroll_area.setGeometry(0, 0, width, height)

                if length > limit:
                    # Only genuinely oversized docks need scrolling.
                    bar = (self.scroll_area.horizontalScrollBar() if horizontal
                           else self.scroll_area.verticalScrollBar())
                    bar.setValue((bar.minimum() + bar.maximum()) // 2)
                else:
                    self.scroll_area.horizontalScrollBar().setValue(0)
                    self.scroll_area.verticalScrollBar().setValue(0)
                    self.content.move(0, 0)
                self._center_overflow_on_resize = False
        finally:
            self.content.setUpdatesEnabled(True)
            self.scroll_area.setUpdatesEnabled(True)
            self.setUpdatesEnabled(True)
            self.update()
        self.allow_context_return()

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
        gap.setFixedSize(self.button_size, self.button_size)
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
        if not isinstance(button, SeparatorButton):
            mime.setUrls([QUrl.fromLocalFile(str(button.path))])
        drag.setMimeData(mime)
        if isinstance(button, SeparatorButton):
            pix = QPixmap(self.button_size, self.button_size)
            pix.fill(Qt.GlobalColor.transparent)
            pp = QPainter(pix)
            pp.setPen(QPen(QColor(150, 150, 150), 2))
            if self.position in ("Bottom", "Top"):
                pp.drawLine(self.button_size//2, 8, self.button_size//2, self.button_size-8)
            else:
                pp.drawLine(8, self.button_size//2, self.button_size-8, self.button_size//2)
            pp.end()
            drag.setPixmap(pix)
        else:
            drag.setPixmap(button.drag_icon_pixmap())
        drag.setHotSpot(QPoint(drag.pixmap().width() // 2, drag.pixmap().height() // 2))
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
        was_internal = self.internal_drop or isinstance(button, SeparatorButton)
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
            if isinstance(self.drag_button, SeparatorButton):
                local.setX(max(0, min(local.x(), self.content.width()-1)))
                local.setY(max(0, min(local.y(), self.content.height()-1)))
            else:
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
            gap.setFixedSize(self.button_size, self.button_size)
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
            if event.type() == QEvent.Type.Enter and not isinstance(obj, SeparatorButton):
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
        pitch = self.button_size + self.icon_spacing
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

    def change_icon_size(self, new_size):
        new_size = max(ICON_SIZES[0], min(ICON_SIZES[-1], int(new_size)))
        if new_size == self.icon_size:
            return
        self.icon_size = new_size
        self.settings.setValue("icon_size", new_size)
        self.icon_size_label.setText(f"Icon size: {new_size} × {new_size} px")

        # A dragged slider already delivers a continuous 1px stream. Animating
        # every single change restarts easing 60+ times a second and creates the
        # lag/jump seen in the recording. Render those changes directly.
        slider = self.sender()
        self.size_animation.stop()
        if isinstance(slider, QSlider) and slider.isSliderDown():
            self.render_icon_size(new_size)
        else:
            # Wheel/keyboard steps are discrete; interpolate between them.
            self.size_animation.setStartValue(float(self._rendered_icon_size))
            self.size_animation.setEndValue(float(new_size))
            self.size_animation.start()

    def render_icon_size(self, value):
        size = max(ICON_SIZES[0], min(ICON_SIZES[-1], round(float(value))))
        if size == self._rendered_icon_size:
            return
        self._rendered_icon_size = size
        self._center_overflow_on_resize = True
        self.button_size = size + 16
        self.setUpdatesEnabled(False)
        for i in range(self.layout.count()):
            widget = self.layout.itemAt(i).widget()
            if widget is not None:
                if isinstance(widget, ShortcutButton):
                    widget.setIconSize(QSize(size, size))
                if isinstance(widget, SeparatorButton):
                    self.size_separator(widget)
                else:
                    widget.setFixedSize(self.button_size, self.button_size)
        try:
            self.update_orientation()
        finally:
            self.setUpdatesEnabled(True)
            self.update()

    def size_separator(self, widget):
        # Distance is the free space on each side of the divider line.
        axis = max(8, 8 + 2 * self.separator_distance)
        if self.position in ("Top", "Bottom"):
            widget.setFixedSize(axis, self.button_size)
        else:
            widget.setFixedSize(self.button_size, axis)

    def change_separator_distance(self, value):
        self.separator_distance = int(value)
        self.settings.setValue("separator_distance", self.separator_distance)
        self.separator_distance_label.setText(
            f"Separator distance: {self.separator_distance} px each side")
        for i in range(self.layout.count()):
            widget = self.layout.itemAt(i).widget()
            if isinstance(widget, SeparatorButton):
                self.size_separator(widget)
        self.update_orientation()

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

    def open_file_location(self, shortcut):
        """Open the actual shortcut target's working folder, never the dock folder."""
        try:
            link = win32com.client.Dispatch("WScript.Shell").CreateShortcut(str(shortcut))
            working = os.path.expandvars(str(link.WorkingDirectory or "").strip('" '))
            target = os.path.expandvars(str(link.TargetPath or "").strip('" '))
            # For shortcuts to Explorer, the arguments usually specify the
            # actual directory; the target executable is just explorer.exe.
            arguments = str(link.Arguments or "").strip()
            if target and Path(target).name.lower() == "explorer.exe" and arguments:
                candidate = os.path.expandvars(arguments.strip('" '))
                if Path(candidate).is_dir():
                    working = candidate
            if working and Path(working).is_dir():
                folder = Path(working)
            elif target and Path(target).exists():
                folder = Path(target) if Path(target).is_dir() else Path(target).parent
            else:
                raise FileNotFoundError("The shortcut target or working directory could not be found.")
            os.startfile(str(folder))
        except Exception as error:
            QMessageBox.warning(self, "Open file location", str(error))

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

    def settle_dock_layout(self):
        """Synchronize actual scroll-widget position with the new dock width."""
        self.layout.invalidate()
        self.layout.activate()
        self.update_orientation()
        self._reset_scroll_for_fitting_content()

    def _reset_scroll_for_fitting_content(self):
        """A dock that fits on screen must start at the FIRST icon.

        QScrollArea may retain the previous horizontal scroll offset after a
        shortcut is removed, even after the viewport has been made smaller.
        That makes the rightmost surviving shortcut appear cut off.
        """
        horizontal = self.position in ("Bottom", "Top")
        bar = (self.scroll_area.horizontalScrollBar() if horizontal
               else self.scroll_area.verticalScrollBar())
        screen = self.get_target_screen()
        available = screen.availableGeometry() if screen else None
        available_length = ((available.width() if horizontal else available.height())
                            if available else 0)
        actual_length = (self.content.width() if horizontal else self.content.height())
        if not available or actual_length <= available_length:
            # No overflow: Qt must not preserve a scroll offset from the
            # previous, longer dock. Explicitly reset the widget position.
            bar.setValue(0)
            self.scroll_area.horizontalScrollBar().setValue(0)
            self.scroll_area.verticalScrollBar().setValue(0)
        else:
            bar.setValue((bar.minimum() + bar.maximum()) // 2)
        self.scroll_area.viewport().update()

    def schedule_dock_layout_settle(self):
        """Repair once Qt and the 145ms icon animation have both settled."""
        self._layout_settle_serial = getattr(self, "_layout_settle_serial", 0) + 1
        serial = self._layout_settle_serial

        def settle():
            if serial == self._layout_settle_serial:
                self.settle_dock_layout()

        QTimer.singleShot(0, settle)
        QTimer.singleShot(180, settle)

    def update_orientation(self):
        horizontal = self.position in ("Bottom", "Top")
        self.layout.setDirection(QBoxLayout.Direction.LeftToRight if horizontal
                                 else QBoxLayout.Direction.TopToBottom)
        count = max(1, self.layout.count())
        length = MARGIN * 2 + (count - 1) * self.icon_spacing
        for i in range(self.layout.count()):
            widget = self.layout.itemAt(i).widget()
            if isinstance(widget, SeparatorButton):
                self.size_separator(widget)
            length += (widget.width() if horizontal else widget.height()) if widget else self.button_size
        thickness = MARGIN * 2 + self.button_size
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
        # Resize and reposition in ONE geometry operation. Resizing first and
        # moving later makes Windows display an intermediate corner-anchored
        # frame, perceived as a jump at high DPI.
        self.set_dock_geometry(width, height, screen)
        self.panel.setGeometry(0, 0, width, height)
        self.scroll_area.setGeometry(0, 0, width, height)
        self.content.setFixedSize(length if horizontal else thickness,
                                  thickness if horizontal else length)
        self.layout.invalidate()
        self.layout.activate()
        # The viewport is clipped at the screen size. When resizing, center
        # the entire icon strip under it so BOTH ends overflow equally.
        if self._center_overflow_on_resize:
            bar = (self.scroll_area.horizontalScrollBar() if horizontal
                   else self.scroll_area.verticalScrollBar())
            bar.setValue((bar.minimum() + bar.maximum()) // 2)
            self._center_overflow_on_resize = False
        elif horizontal:
            self.scroll_area.horizontalScrollBar().setValue(old_horizontal)
        else:
            self.scroll_area.verticalScrollBar().setValue(old_vertical)
        # A reduced dock may still have the previous scroll position.
        # Force the first icon to x=0 when the entire strip fits.
        if length <= max_length:
            self.scroll_area.horizontalScrollBar().setValue(0)
            self.scroll_area.verticalScrollBar().setValue(0)
        # set_dock_geometry already applied the correct anchored position.

    def get_target_screen(self):
        if self.primary_only:
            return QApplication.primaryScreen()

        screen = QApplication.screenAt(
            QCursor.pos()
        )

        return screen or QApplication.primaryScreen()

    def anchored_position(self, screen, width, height):
        geometry = screen.availableGeometry()
        if self.position == "Bottom":
            return QPoint(geometry.x() + (geometry.width() - width) // 2,
                          geometry.bottom() - height + 1 - self.offset)
        if self.position == "Top":
            return QPoint(geometry.x() + (geometry.width() - width) // 2,
                          geometry.top() + self.offset)
        if self.position == "Left":
            return QPoint(geometry.left() + self.offset,
                          geometry.y() + (geometry.height() - height) // 2)
        return QPoint(geometry.right() - width + 1 - self.offset,
                      geometry.y() + (geometry.height() - height) // 2)

    def set_dock_geometry(self, width, height, screen=None):
        screen = screen or self.get_target_screen()
        if screen is None:
            self.resize(width, height)
            return
        self.current_screen = screen
        position = self.anchored_position(screen, width, height)
        self.visible_position = position
        if not self.animating:
            # QWindow geometry is committed together: center stays fixed and
            # the edge facing the taskbar never shifts during a size frame.
            if self.geometry() != QRect(position.x(), position.y(), width, height):
                self.setGeometry(position.x(), position.y(), width, height)
        else:
            # Do not override the active auto-hide/show slide position.
            self.resize(width, height)

    def position_dock(self, screen=None):
        screen = screen or self.get_target_screen()
        if screen is None:
            return
        self.current_screen = screen
        self.visible_position = self.anchored_position(screen, self.width(), self.height())
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
        if (self.dragging or self.icon_menu_open or self.tray_menu_open
                or self.rename_dialog_open or time.monotonic() < self.context_grace_until):
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
        # Crossing the screen to another edge needs more time.
        self.allow_context_return(3.0)

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
                      self.spacing_label, self.icon_size_label, self.separator_distance_label, self.offset_label):
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
        # Windows' system hyperlink/accent color is exposed through the
        # current Qt application palette. Apply it to the anchor itself:
        # styling QLabel alone does not override HTML <a> colors.
        link_color = windows_accent_color().name()
        github.setText(
            f'<a href="https://github.com/packetlauncher" '
            f'style="color: {link_color};">github.com/packetlauncher</a>')
        github.setStyleSheet("QLabel { background: transparent; }")
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

    def contextMenuEvent(self, event):
        # Right-click the empty dock surface to add an installed app.
        menu = QMenu(self)
        style_icon_context_menu(menu)
        add = menu.addAction("Add Windows app...")
        menu.addSeparator()
        exit_action = menu.addAction("Exit dock")
        self.icon_menu_open = True
        self.hide_deadline = None
        try:
            chosen = menu.exec(self.mapToGlobal(event.pos()))
        finally:
            self.icon_menu_open = False
            self.allow_context_return()
        if chosen == add:
            self.browse_windows_apps()
        elif chosen == exit_action:
            QApplication.instance().quit()

    def windows_app_icon(self, app_id, name):
        """Look up actual installed-app art, preferring package manifest assets."""
        if not hasattr(self, "_windows_app_icon_cache"):
            self._windows_app_icon_cache = {}
        if app_id in self._windows_app_icon_cache:
            return self._windows_app_icon_cache[app_id]
        icon = QIcon()
        # Packaged apps have a family name before '!'. Resolve their installed
        # manifest and choose the highest-resolution logo asset available.
        package_family = app_id.split("!", 1)[0]
        for root in getattr(self, "_windows_app_packages", {}).get(package_family, []):
            try:
                import xml.etree.ElementTree as ET
                manifest = Path(root) / "AppxManifest.xml"
                if not manifest.is_file():
                    continue
                tree = ET.parse(str(manifest))
                candidates = []
                for node in tree.iter():
                    for key, val in node.attrib.items():
                        if key.rsplit("}", 1)[-1].lower() in (
                                "square44x44logo", "square150x150logo", "logo"):
                            candidates.append(val)
                for relative in candidates:
                    stem = Path(root) / relative.replace("\\\\", "/")
                    files = [stem]
                    files.extend(stem.parent.glob(
                        stem.stem + ".scale-*" + stem.suffix))
                    files.extend(stem.parent.glob(
                        stem.stem + ".targetsize-*" + stem.suffix))
                    files.extend(stem.parent.glob(
                        stem.stem + ".scale-*.png"))
                    files.extend(stem.parent.glob(
                        stem.stem + ".targetsize-*.png"))
                    for candidate in sorted(files, key=lambda f: (
                            "targetsize-256" in f.name, "scale-400" in f.name,
                            "scale-200" in f.name), reverse=True):
                        if candidate.is_file():
                            test = QIcon(str(candidate))
                            if not test.isNull():
                                icon = test
                                break
                    if not icon.isNull():
                        break
                if not icon.isNull():
                    break
            except (OSError, ValueError, ET.ParseError):
                continue
        if icon.isNull():
            for folder in (Path(os.environ.get("PROGRAMDATA", r"C:\\ProgramData"))
                           / "Microsoft/Windows/Start Menu/Programs",
                           Path(os.environ.get("APPDATA", ""))
                           / "Microsoft/Windows/Start Menu/Programs"):
                if folder.is_dir():
                    for shortcut in folder.rglob("*.lnk"):
                        if shortcut.stem.casefold() == name.casefold():
                            icon = self.get_shortcut_icon(shortcut)
                            if not icon.isNull():
                                break
                if not icon.isNull():
                    break
        if icon.isNull():
            icon = self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)
        self._windows_app_icon_cache[app_id] = icon
        return icon

    def browse_windows_apps(self):
        """List installed Start-menu/Store apps and create a portable .lnk."""
        dialog = QDialog(None)
        dialog.setWindowTitle("Add Windows app to dock")
        dialog.setWindowIcon(self.make_tray_icon())
        dialog.resize(480, 540)
        layout = QVBoxLayout(dialog)
        search = QLineEdit(dialog)
        search.setPlaceholderText("Search installed Windows apps...")
        layout.addWidget(search)
        app_list = QListWidget(dialog)
        layout.addWidget(app_list)
        controls = QHBoxLayout()
        add_button = QPushButton("Add to Dock", dialog)
        cancel_button = QPushButton("Cancel", dialog)
        controls.addStretch()
        controls.addWidget(add_button)
        controls.addWidget(cancel_button)
        layout.addLayout(controls)
        add_button.setEnabled(False)
        app_list.itemSelectionChanged.connect(
            lambda: add_button.setEnabled(bool(app_list.selectedItems())))
        cancel_button.clicked.connect(dialog.reject)
        app_list.itemDoubleClicked.connect(lambda _: dialog.accept())
        add_button.clicked.connect(dialog.accept)
        dark = windows_system_dark_theme()
        bg = "#292929" if dark else "#f9f9f9"
        fg = "#f5f5f5" if dark else "#202020"
        dialog.setStyleSheet(
            f"QDialog {{ background: {bg}; color: {fg}; }}"
            f"QLabel, QLineEdit, QListWidget {{ color: {fg}; }}"
            f"QLineEdit, QListWidget {{ background: {'#202020' if dark else '#ffffff'}; }}"
        )
        from PySide6.QtWidgets import QListWidgetItem
        app_list.setIconSize(QSize(40, 40))
        app_list.setSpacing(3)
        app_list.setStyleSheet(app_list.styleSheet() +
            "QListWidget::item { min-height: 48px; padding: 3px 6px; }")

        def filter_apps(query):
            for i in range(app_list.count()):
                item = app_list.item(i)
                item.setHidden(query.casefold() not in item.text().casefold())
        search.textChanged.connect(filter_apps)

        def fill_apps(apps):
            app_list.clear()
            for name, app_id in apps:
                item = QListWidgetItem(name)
                item.setData(Qt.ItemDataRole.UserRole, app_id)
                app_list.addItem(item)
            filter_apps(search.text())
            # Resolve expensive manifest artwork gradually, keeping the
            # window responsive and displaying names immediately.
            next_icon = [0]
            def fill_icons():
                count = 0
                while next_icon[0] < app_list.count() and count < 5:
                    item = app_list.item(next_icon[0])
                    next_icon[0] += 1
                    if item:
                        item.setIcon(self.windows_app_icon(
                            item.data(Qt.ItemDataRole.UserRole), item.text()))
                    count += 1
                if next_icon[0] < app_list.count():
                    QTimer.singleShot(0, fill_icons)
            QTimer.singleShot(0, fill_icons)

        def apps_loaded(apps, packages, error):
            if error:
                app_list.clear()
                app_list.addItem("Unable to load apps: " + error)
                return
            self._windows_app_packages = packages
            self._windows_apps_cached = apps
            fill_apps(apps)

        if hasattr(self, "_windows_apps_cached"):
            fill_apps(self._windows_apps_cached)
        else:
            loading_item = QListWidgetItem("Loading installed Windows apps...")
            loading_item.setFlags(Qt.ItemFlag.NoItemFlags)
            app_list.addItem(loading_item)
            self._apps_loader = WindowsAppsLoader(self)
            self._apps_loader.loaded.connect(apps_loaded)
            self._apps_loader.start()

        set_native_dialog_theme(dialog)
        self.rename_dialog_open = True
        self.hide_deadline = None
        try:
            accepted = dialog.exec() == QDialog.DialogCode.Accepted
        finally:
            self.rename_dialog_open = False
            if hasattr(self, "_apps_loader") and self._apps_loader.isRunning():
                try:
                    self._apps_loader.loaded.disconnect(apps_loaded)
                except (RuntimeError, TypeError):
                    pass
            self.allow_context_return()
        if not accepted or not app_list.currentItem():
            return
        item = app_list.currentItem()
        self.add_windows_app(item.text(), item.data(Qt.ItemDataRole.UserRole))

    def windows_app_icon_file(self, app_id, name):
        """Write a persistent Windows-compatible PNG-compressed ICO asset.

        The .lnk points at this file, so Explorer can display the actual app
        icon even when the shortcut is dragged out of the dock.
        """
        cache = (Path(os.environ.get("LOCALAPPDATA", str(BASE_DIR))) /
                 "ShortcutsDock" / "AppIcons")
        cache.mkdir(parents=True, exist_ok=True)
        filename = hashlib.sha256(app_id.encode("utf-8")).hexdigest()[:24] + ".ico"
        destination = cache / filename
        if destination.is_file() and destination.stat().st_size > 100:
            return destination
        icon = self.windows_app_icon(app_id, name)
        if icon.isNull():
            return None
        pixmap = icon.pixmap(QSize(256, 256))
        if pixmap.isNull():
            return None
        image = pixmap.toImage()
        image.setDevicePixelRatio(1.0)
        image = image.convertToFormat(QImage.Format.Format_ARGB32)
        # Pad to a square without stretching the source artwork.
        if image.width() != image.height():
            side = max(image.width(), image.height())
            square = QImage(side, side, QImage.Format.Format_ARGB32)
            square.fill(Qt.GlobalColor.transparent)
            painter = QPainter(square)
            painter.drawImage((side-image.width())//2,
                              (side-image.height())//2, image)
            painter.end()
            image = square
        image = image.scaled(256, 256, Qt.AspectRatioMode.IgnoreAspectRatio,
                             Qt.TransformationMode.SmoothTransformation)
        buffer = QBuffer()
        buffer.open(QIODevice.OpenModeFlag.WriteOnly)
        if not image.save(buffer, "PNG"):
            return None
        png = bytes(buffer.data())
        buffer.close()
        # ICO header, 256x256 directory entry, PNG payload.
        ico = (struct.pack("<HHH", 0, 1, 1) +
               struct.pack("<BBBBHHII", 0, 0, 0, 0, 1, 32, len(png), 22) + png)
        destination.write_bytes(ico)
        return destination

    def add_windows_app(self, name, app_id):
        """Create a real Windows shell .lnk, suitable for Explorer drag/drop."""
        # explorer.exe resolves the AppsFolder AUMID, including packaged apps
        # which have no ordinary executable path available to the user.
        # A .lnk is a real filesystem object so Explorer drag-out works as
        # it does for existing dock shortcuts.
        safe_name = re.sub(r'[<>:"/\\\\|?*]', "_", name).strip(" .") or "Windows App"
        path = self.unique_path(SHORTCUTS_DIR, safe_name)
        try:
            link = win32com.client.Dispatch("WScript.Shell").CreateShortcut(str(path))
            link.TargetPath = str(Path(os.environ.get("WINDIR", r"C:\\Windows")) / "explorer.exe")
            link.Arguments = "shell:AppsFolder\\" + app_id
            link.Description = f"Windows app: {name}"
            icon_file = self.windows_app_icon_file(app_id, name)
            if icon_file:
                link.IconLocation = str(icon_file) + ",0"
            link.Save()
            self.save_order(self.shortcut_paths() + [path])
            self.load_shortcuts()
            self.allow_context_return(3.0)
        except Exception as error:
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass
            QMessageBox.warning(self, "Add Windows app",
                                f"Could not add {name}:\\n{error}")

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

        icon_size_action = QWidgetAction(menu)
        icon_size_widget = QWidget(menu)
        icon_size_layout = QVBoxLayout(icon_size_widget)
        icon_size_layout.setContentsMargins(12, 6, 12, 8)
        icon_size_layout.setSpacing(4)
        self.icon_size_label = QLabel(f"Icon size: {self.icon_size} × {self.icon_size} px")
        icon_size_layout.addWidget(self.icon_size_label)
        icon_size_slider = IconSizeSlider(Qt.Orientation.Horizontal)
        icon_size_slider.setRange(ICON_SIZES[0], ICON_SIZES[-1])
        icon_size_slider.setSingleStep(1)
        icon_size_slider.setPageStep(1)
        icon_size_slider.setTickInterval(16)
        icon_size_slider.setValue(self.icon_size)
        icon_size_slider.setMinimumWidth(190)
        icon_size_slider.valueChanged.connect(self.change_icon_size)
        self.style_menu_slider(icon_size_slider)
        icon_size_layout.addWidget(icon_size_slider)
        icon_size_action.setDefaultWidget(icon_size_widget)
        menu.addAction(icon_size_action)

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

        separator_action = QWidgetAction(menu)
        separator_widget = QWidget(menu)
        separator_layout = QVBoxLayout(separator_widget)
        separator_layout.setContentsMargins(12, 6, 12, 8)
        separator_layout.setSpacing(4)
        self.separator_distance_label = QLabel(
            f"Separator distance: {self.separator_distance} px each side")
        separator_layout.addWidget(self.separator_distance_label)
        separator_slider = QSlider(Qt.Orientation.Horizontal)
        separator_slider.setRange(0, 100)
        separator_slider.setValue(self.separator_distance)
        separator_slider.setMinimumWidth(190)
        separator_slider.valueChanged.connect(self.change_separator_distance)
        self.style_menu_slider(separator_slider)
        separator_layout.addWidget(separator_slider)
        separator_action.setDefaultWidget(separator_widget)
        menu.addAction(separator_action)

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
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            # A click is a reveal request, never a toggle or animation restart.
            if self.isVisible() and not self.animation_hiding:
                return
            if self.animating and not self.animation_hiding:
                return
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
