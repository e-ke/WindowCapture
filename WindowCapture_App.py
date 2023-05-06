import sys
import os
from PySide2 import QtCore, QtWidgets
from PySide2.QtWidgets import QHBoxLayout, QVBoxLayout, QGridLayout, QGroupBox, QComboBox, QPushButton, QWidget, QSizePolicy
import win32gui
import win32ui
from ctypes import windll

class StartWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.windows = []

    def initUI(self):
        self.setWindowTitle('Start Window')
        self.setGeometry(100, 100, 400, 200)

        # ウィンドウを選択するためのコンボボックス
        self.combobox = QComboBox(self)
        self.combobox.addItem('Select a window')
        self.combobox.addItems(self.get_window_list())

        # キャプチャー実行ボタン
        self.btn_capture = QPushButton('Capture', self)
        self.btn_capture.clicked.connect(self.capture_window)

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addWidget(self.combobox)
        vbox.addWidget(self.btn_capture)
        vbox.addStretch(1)
        self.setLayout(vbox)

    def get_window_list(self):
        # ウィンドウのリストを取得する
        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if window_text:
                    results.append(window_text)

        results = []
        win32gui.EnumWindows(enum_callback, results)
        return results

    def closeEvent(self, event):
        if hasattr(self, 'main_app'):
            self.main_app.close()
        event.accept()

    def capture_window(self):
        selected_window_title = self.combobox.currentText()
        if selected_window_title != 'Select a window':
            hwnd = win32gui.FindWindow(None, selected_window_title)
            if hwnd != 0:
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                width = right - left
                height = bottom - top
                
                # ウィンドウのデバイスコンテキストを取得
                hwnd_dc = win32gui.GetWindowDC(hwnd)
                mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
                save_dc = mfc_dc.CreateCompatibleDC()

                # ビットマップを作成
                save_bitmap = win32ui.CreateBitmap()
                save_bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
                save_dc.SelectObject(save_bitmap)

                # ウィンドウのキャプチャ
                result = windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 0)
                # ファイル名に連番を振るためのカウンター
                i = 0
                while True:
                    filename = f'captured/captured_window_{i}.bmp'
                    if not os.path.isfile(filename):
                        break
                    i += 1
                # ビットマップを保存
                if result == 1:
                    save_bitmap.SaveBitmapFile(save_dc, filename)
                    print(f'Captured and saved the image successfully as "{filename}".')
                else:
                    print('Failed to capture the image.')

                # 後処理
                save_dc.DeleteDC()
                mfc_dc.DeleteDC()
                win32gui.ReleaseDC(hwnd, hwnd_dc)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    start_window = StartWindow()
    start_window.show()
    sys.exit(app.exec_())