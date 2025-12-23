import os
import sys
import glob
import random
from pathlib import Path

from PyQt5 import QtCore, QtGui, QtWidgets
from PIL import Image

import win32gui
import win32con
import win32api


# -------------------------------
# Win32 desktop attachment helper
# -------------------------------

def attach_to_desktop(hwnd: int):
    """
    Attach window to the desktop "WorkerW" layer so it appears above icons.
    """
    progman = win32gui.FindWindow("Progman", None)

    # Force WorkerW creation (classic technique)
    win32gui.SendMessageTimeout(
        progman,
        0x052C,  # undocumented message
        0,
        0,
        win32con.SMTO_NORMAL,
        1000
    )

    workerw = None

    def enum_windows_callback(top_hwnd, _):
        nonlocal workerw
        shell_def_view = win32gui.FindWindowEx(top_hwnd, 0, "SHELLDLL_DefView", None)
        if shell_def_view != 0:
            # WorkerW is usually a sibling following this one
            workerw = win32gui.FindWindowEx(0, top_hwnd, "WorkerW", None)
            return False
        return True

    win32gui.EnumWindows(enum_windows_callback, None)

    if workerw:
        win32gui.SetParent(hwnd, workerw)


def set_click_through(hwnd: int, enabled: bool):
    """
    Enable/disable click-through mode.
    """
    ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)

    if enabled:
        ex_style |= win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT
    else:
        ex_style &= ~win32con.WS_EX_TRANSPARENT
        ex_style |= win32con.WS_EX_LAYERED

    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style)


def make_tool_window(hwnd: int):
    """
    Remove Alt-Tab + taskbar presence.
    """
    ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    ex_style |= win32con.WS_EX_TOOLWINDOW
    ex_style &= ~win32con.WS_EX_APPWINDOW
    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style)


# -------------------------------
# Frame Widget
# -------------------------------

class PictureFrame(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PictureFrame")
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint |
            QtCore.Qt.WindowStaysOnTopHint |
            QtCore.Qt.Tool
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)

        self.resize(520, 380)

        self.dragging = False
        self.drag_offset = QtCore.QPoint()

        self.is_locked = False
        self.is_topmost = True
        self.click_through = False
        self.attached_to_desktop = False

        self.images = []
        self.index = 0
        # set folder path to C:\Users\tomer.labin\Pictures\PictureFrame
        self.folder = Path("C:/Users/tomer.labin/Pictures/PictureFrame")
        self.folder.mkdir(parents=True, exist_ok=True)

        self.load_images()

        # Slideshow timer
        self.timer = QtCore.QTimer(self)
        self.timer.setInterval(10_000)
        self.timer.timeout.connect(self.next_image)
        self.timer.start()

        # Context menu
        self.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.open_context_menu)

    def showEvent(self, event):
        super().showEvent(event)
        hwnd = int(self.winId())

        if self.attached_to_desktop:
            attach_to_desktop(hwnd)

        make_tool_window(hwnd)
        set_click_through(hwnd, self.click_through)

    # -------------------------------
    # Image handling
    # -------------------------------

    def load_images(self):
        exts = ("*.jpg", "*.jpeg", "*.png", "*.bmp")
        files = []
        for e in exts:
            files.extend(glob.glob(str(self.folder / e)))

        self.images = sorted(files)
        self.index = 0

    def next_image(self):
        if not self.images:
            return
        self.index = (self.index + 1) % len(self.images)
        self.update()

    def current_pixmap(self):
        if not self.images:
            return None

        path = self.images[self.index]
        try:
            img = Image.open(path)
            img = img.convert("RGBA")
            qimg = QtGui.QImage(img.tobytes("raw", "RGBA"), img.width, img.height, QtGui.QImage.Format_RGBA8888)
            return QtGui.QPixmap.fromImage(qimg.copy())
        except Exception:
            return None

    # -------------------------------
    # Painting the picture frame
    # -------------------------------

    def paintEvent(self, event):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing, True)

        rect = self.rect().adjusted(10, 10, -10, -10)

        # Shadow
        shadow_color = QtGui.QColor(0, 0, 0, 120)
        p.setBrush(shadow_color)
        p.setPen(QtCore.Qt.NoPen)
        p.drawRoundedRect(rect.translated(6, 8), 20, 20)

        # Outer "wood" frame gradient
        wood = QtGui.QLinearGradient(rect.topLeft(), rect.bottomRight())
        wood.setColorAt(0.0, QtGui.QColor("#7A4B2A"))
        wood.setColorAt(0.35, QtGui.QColor("#B07A46"))
        wood.setColorAt(1.0, QtGui.QColor("#6A3F22"))

        p.setBrush(wood)
        p.setPen(QtCore.Qt.NoPen)
        p.drawRoundedRect(rect, 20, 20)

        # Inner bevel
        inner = rect.adjusted(16, 16, -16, -16)
        bevel_pen = QtGui.QPen(QtGui.QColor(255, 255, 255, 80), 2)
        p.setPen(bevel_pen)
        p.setBrush(QtCore.Qt.NoBrush)
        p.drawRoundedRect(inner, 14, 14)

        # Matte
        matte = inner.adjusted(10, 10, -10, -10)
        p.setPen(QtCore.Qt.NoPen)
        p.setBrush(QtGui.QColor("#F2EFE6"))
        p.drawRoundedRect(matte, 10, 10)

        # Photo area
        photo_rect = matte.adjusted(10, 10, -10, -10)
        p.setBrush(QtGui.QColor("#111111"))
        p.drawRoundedRect(photo_rect, 8, 8)

        pix = self.current_pixmap()
        if pix:
            # Scale and crop to fill
            scaled = pix.scaled(photo_rect.size(), QtCore.Qt.KeepAspectRatioByExpanding, QtCore.Qt.SmoothTransformation)
            x = (scaled.width() - photo_rect.width()) // 2
            y = (scaled.height() - photo_rect.height()) // 2
            crop = scaled.copy(x, y, photo_rect.width(), photo_rect.height())

            # Clip to rounded rect
            path = QtGui.QPainterPath()
            path.addRoundedRect(QtCore.QRectF(photo_rect), 8, 8)
            p.setClipPath(path)
            p.drawPixmap(photo_rect.topLeft(), crop)

        # Resize grip
        grip_rect = QtCore.QRect(self.width() - 28, self.height() - 28, 18, 18)
        p.setClipping(False)
        p.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 120), 2))
        p.drawLine(grip_rect.bottomLeft(), grip_rect.topRight())
        p.drawLine(grip_rect.bottomLeft() + QtCore.QPoint(6, 0), grip_rect.topRight() + QtCore.QPoint(0, 6))
        p.drawLine(grip_rect.bottomLeft() + QtCore.QPoint(12, 0), grip_rect.topRight() + QtCore.QPoint(0, 12))

    # -------------------------------
    # Dragging & resizing
    # -------------------------------

    def mousePressEvent(self, e):
        if self.is_locked or self.click_through:
            return

        # Resize area bottom-right
        if e.button() == QtCore.Qt.LeftButton:
            if e.pos().x() > self.width() - 40 and e.pos().y() > self.height() - 40:
                self.start_resize(e.globalPos())
            else:
                self.dragging = True
                self.drag_offset = e.globalPos() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self.is_locked or self.click_through:
            return

        if hasattr(self, "_resizing") and self._resizing:
            delta = e.globalPos() - self._resize_origin
            new_w = max(240, self._start_size.width() + delta.x())
            new_h = max(180, self._start_size.height() + delta.y())
            self.resize(new_w, new_h)
            self.update()
        elif self.dragging:
            self.move(e.globalPos() - self.drag_offset)

    def mouseReleaseEvent(self, e):
        self.dragging = False
        self._resizing = False

    def start_resize(self, global_pos):
        self._resizing = True
        self._resize_origin = global_pos
        self._start_size = self.size()

    # -------------------------------
    # Context menu
    # -------------------------------

    def open_context_menu(self, pos):
        menu = QtWidgets.QMenu(self)

        act_next = menu.addAction("Next picture")
        act_folder = menu.addAction("Set pictures folder...")

        menu.addSeparator()

        act_top = menu.addAction("Always on top")
        act_top.setCheckable(True)
        act_top.setChecked(self.is_topmost)

        act_lock = menu.addAction("Lock position")
        act_lock.setCheckable(True)
        act_lock.setChecked(self.is_locked)

        act_ct = menu.addAction("Click-through (poster mode)")
        act_ct.setCheckable(True)
        act_ct.setChecked(self.click_through)

        act_attach = menu.addAction("Attach to desktop layer (above icons)")
        act_attach.setCheckable(True)
        act_attach.setChecked(self.attached_to_desktop)

        menu.addSeparator()
        act_exit = menu.addAction("Exit")

        action = menu.exec_(self.mapToGlobal(pos))
        if action is None:
            return

        if action == act_next:
            self.next_image()

        elif action == act_folder:
            folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Choose picture folder", str(self.folder))
            if folder:
                self.folder = Path(folder)
                self.load_images()
                self.update()

        elif action == act_top:
            self.is_topmost = act_top.isChecked()
            self.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint, self.is_topmost)
            self.show()  # reapply flags

        elif action == act_lock:
            self.is_locked = act_lock.isChecked()

        elif action == act_ct:
            self.click_through = act_ct.isChecked()
            set_click_through(int(self.winId()), self.click_through)

        elif action == act_attach:
            self.attached_to_desktop = act_attach.isChecked()
            if self.attached_to_desktop:
                attach_to_desktop(int(self.winId()))
            else:
                win32gui.SetParent(int(self.winId()), 0)

        elif action == act_exit:
            QtWidgets.QApplication.quit()


def main():
    app = QtWidgets.QApplication(sys.argv)
    w = PictureFrame()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
