import sys
import asyncio
import logging
import ctypes
from winsdk.windows.media.control import GlobalSystemMediaTransportControlsSessionManager as MediaManager
from PyQt5.QtWidgets import QLabel, QApplication, QMenu, QAction, QSystemTrayIcon
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QPainter, QColor, QIcon, QFont, QFontDatabase

logging.basicConfig(filename='app.log', level=logging.ERROR)

# Определяем константы для использования в WinAPI
CSIDL_APPDATA = 0x001a
SHGFP_TYPE_CURRENT = 0
TBPFLAG_NOPROGRESS = 0x0001
TBPFLAG_NOCLOSE = 0x0010
TBPFLAG_RESIZABLE = 0x0040

# Функция для получения пути к папке AppData
def get_appdata_folder():
    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_APPDATA, None, SHGFP_TYPE_CURRENT, buf)
    return buf.value

class Overlay(QLabel):
    MAX_LENGTH = 50  # Максимальная длина текста
    NEXT_ROW = True

    def __init__(self):
        super().__init__()

        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.WindowTransparentForInput)
        self.setAttribute(Qt.WA_TranslucentBackground)

        self.opacity = 0.7
        self.setStyleSheet("QLabel { background-color: rgba(0, 0, 0, 100); color: white; font-size: 10pt; font-weight: bold; }")
        font_path = "./fonts/font.ttf"
        font_id = QFontDatabase.addApplicationFont(font_path)
        font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
        self.setFont(QFont(font_family, 12))

        self.current_title = ""
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_update_and_show)
        self.timer.start(1000)

        self.hide_timer = QTimer(self)
        self.hide_timer.timeout.connect(self.hide_overlay)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setOpacity(self.opacity)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor(0, 0, 0, 250))
        super().paintEvent(event)

    def update_opacity(self, opacity):
        self.opacity = opacity
        self.repaint()

    async def get_target_id(self):
        try:
            sessions = await MediaManager.request_async()
            current_session = sessions.get_current_session()

            if current_session:
                return current_session.source_app_user_model_id
            else:
                raise Exception('No active media session found')
        except Exception as e:
            logging.error(f"Error getting target ID: {e}")
            raise

    async def get_media_info(self):
        try:
            target_id = await self.get_target_id()
            sessions = await MediaManager.request_async()
            current_session = sessions.get_current_session()

            if current_session and current_session.source_app_user_model_id == target_id:
                info = await current_session.try_get_media_properties_async()
                info_dict = {song_attr: info.__getattribute__(song_attr) for song_attr in dir(info) if song_attr[0] != '_'}
                info_dict['genres'] = list(info_dict.get('genres', []))
                return info_dict
            else:
                raise Exception('TARGET_PROGRAM is not the current media session')
        except Exception as e:
            logging.error(f"Error getting media info: {e}")
            raise

    def check_update_and_show(self):
        try:
            current_media_info = asyncio.run(self.get_media_info())
            title = current_media_info.get('title', 'Unknown')

            if title != self.current_title:
                self.current_title = title
                self.show_overlay()

        except Exception as e:
            logging.error(f"Error checking update and show: {e}")
            self.hide()

    def show_overlay(self):
        try:
            current_media_info = asyncio.run(self.get_media_info())
            title = current_media_info.get('title', 'Unknown')
            artist = current_media_info.get('artist', 'Unknown')

            if not title:
                title = "Unknown"
            if not artist:
                artist = "Unknown"

            if Overlay.NEXT_ROW:
                overlay_text = f"  NOW: {title} - {artist}"
            else:
                overlay_text = f"NOW: {artist} - {title[:Overlay.MAX_LENGTH]}{'...' if len(title) > Overlay.MAX_LENGTH else ''}"

            self.setText(overlay_text)
            self.setGeometry(2, 2, self.fontMetrics().width(overlay_text) + 10, self.fontMetrics().height() * (2 if Overlay.NEXT_ROW else 1))
            self.update_opacity(0.75)
            self.show()

            if self.hide_timer.isActive():
                self.hide_timer.stop()

            self.hide_timer.start(15000)

        except Exception as e:
            logging.error(f"Error showing overlay: {e}")
            self.hide()

    def hide_overlay(self):
        self.update_opacity(0.0)
        self.hide()

class TrayIcon(QSystemTrayIcon):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setIcon(QIcon("icons/overlay.png"))

        exit_action = QAction(QIcon("icons/exit.png"), "Exit", self)
        exit_action.triggered.connect(self.exit_app)

        tray_menu = QMenu()
        tray_menu.addAction(exit_action)

        self.setContextMenu(tray_menu)

    def exit_app(self):
        sys.exit()

def add_to_taskbar():
    try:
        taskbar = ctypes.windll.shell32.CoCreateInstance
        taskbar.argtypes = [ctypes.wintypes.LPVOID, ctypes.wintypes.LPVOID, ctypes.wintypes.LPVOID, ctypes.wintypes.LPVOID, ctypes.wintypes.LPVOID]
        taskbar.restype = ctypes.wintypes.LPVOID

        taskbar(None, None, None, None, None)

    except Exception as e:
        logging.error(f"Error adding to taskbar: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)

    overlay = Overlay()
    overlay.show()

    tray_icon = TrayIcon()
    tray_icon.show()

    add_to_taskbar()

    sys.exit(app.exec_())
