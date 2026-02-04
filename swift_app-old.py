
# swift_app.py
import os
import sys
import traceback

from PySide6.QtCore import Qt, QThread, Signal, QRect
from PySide6.QtGui import QIcon, QPixmap, QFont, QPainter, QPainterPath
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QLineEdit,
    QFileDialog, QProgressBar, QMessageBox, QHBoxLayout, QVBoxLayout,
    QGroupBox, QFormLayout
)

import swift_core


# -------------------------
# Logo 圆角正方形处理
# -------------------------
def rounded_square_pixmap(pix: QPixmap, size: int = 56, radius: int = 12) -> QPixmap:
    """
    将图片：中心裁正方形 -> 缩放 -> 圆角裁切
    """
    if pix.isNull():
        return pix

    w, h = pix.width(), pix.height()
    side = min(w, h)
    x = (w - side) // 2
    y = (h - side) // 2

    square = pix.copy(QRect(x, y, side, side)).scaled(
        size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation
    )

    out = QPixmap(size, size)
    out.fill(Qt.transparent)

    painter = QPainter(out)
    painter.setRenderHint(QPainter.Antialiasing, True)
    path = QPainterPath()
    path.addRoundedRect(0, 0, size, size, radius, radius)
    painter.setClipPath(path)
    painter.drawPixmap(0, 0, square)
    painter.end()

    return out


# =========================
# Worker Thread
# =========================
class SwiftWorker(QThread):
    progress = Signal(int, int, str)     # done, total, filename
    status = Signal(str)
    finished_ok = Signal(str)            # output_path
    failed = Signal(str)

    def __init__(self, input_dir, output_dir, mapping_file, mapping_sheet):
        super().__init__()
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.mapping_file = mapping_file
        self.mapping_sheet = mapping_sheet

    def run(self):
        try:
            def progress_cb(done, total, fn):
                self.progress.emit(done, total, fn)

            def status_cb(msg):
                self.status.emit(msg)

            out = swift_core.run_swift_batch(
                input_dir=self.input_dir,
                output_dir=self.output_dir,
                mapping_file=self.mapping_file,
                mapping_sheet=self.mapping_sheet,
                progress_callback=progress_cb,
                status_callback=status_cb
            )
            self.finished_ok.emit(out)
        except Exception as e:
            err = f"{e}\n\n{traceback.format_exc()}"
            self.failed.emit(err)


# =========================
# Main Window
# =========================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("SWIFT Data Collection")
        self.setMinimumWidth(860)
        self.setMinimumHeight(560)

        # ------- icon -------
        icon_path = r"C:\Users\MY43DN\Desktop\app.ico"
        if not os.path.exists(icon_path):
            icon_path = os.path.join(os.path.dirname(__file__), "app.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # MessageBox 深色样式（解决白底白字看不清）
        self.msgbox_style = """
        QMessageBox {
            background-color: #121212;
        }
        QLabel {
            color: #EAEAEA;
            font-size: 12px;
        }
        QPushButton {
            background-color: #1E1E1E;
            color: #EAEAEA;
            border: 1px solid #333333;
            border-radius: 8px;
            padding: 8px 14px;
            min-width: 90px;
        }
        QPushButton:hover {
            border: 1px solid #FFB000;
        }
        """

        # ------- root widget -------
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(14)

        # ------- header with logo -------
        header = QHBoxLayout()
        header.setSpacing(14)

        logo_label = QLabel()
        logo_label.setFixedSize(56, 56)

        logo_path = r"C:\Users\MY43DN\Desktop\ing-logo.png"
        if not os.path.exists(logo_path):
            logo_path = os.path.join(os.path.dirname(__file__), "ing-logo.png")

        if os.path.exists(logo_path):
            pix = QPixmap(logo_path)
            logo_label.setPixmap(rounded_square_pixmap(pix, size=56, radius=12))
        else:
            logo_label.setText("ING")
            logo_label.setStyleSheet("color:#ff9900; font-size:28px; font-weight:700;")

        title_box = QVBoxLayout()
        title = QLabel("SWIFT Data Collection")
        title.setStyleSheet("color:#EAEAEA; font-size:26px; font-weight:800; letter-spacing:0.5px;")

        # 去掉“黑科技/专业软件”字眼
        subtitle = QLabel("一键提取 Step3_Final")
        subtitle.setStyleSheet("color:#B8B8B8; font-size:13px;")

        title_box.addWidget(title)
        title_box.addWidget(subtitle)

        header.addWidget(logo_label, 0, Qt.AlignLeft | Qt.AlignVCenter)
        header.addLayout(title_box, 1)
        layout.addLayout(header)

        # ------- path group -------
        group = QGroupBox("运行配置")
        group.setStyleSheet("""
            QGroupBox {
                color:#EAEAEA;
                border:1px solid #2A2A2A;
                border-radius:12px;
                margin-top:10px;
                padding:12px;
                background-color:#141414;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px 0 6px;
            }
        """)
        form = QFormLayout(group)
        form.setVerticalSpacing(12)
        form.setLabelAlignment(Qt.AlignRight)

        # ✅ 固定默认路径（按你要求）
        default_input = r"Z:/To Jimmy Yu/Swift Data Collection/Swift"
        default_output = r"Z:/To Jimmy Yu/Swift Data Collection"
        default_mapping = r"Z:/To Jimmy Yu/Swift Data Collection/Swift Data Collection.xlsx"

        self.input_edit = QLineEdit(default_input)
        self.output_edit = QLineEdit(default_output)
        self.map_edit = QLineEdit(default_mapping)
        self.sheet_edit = QLineEdit(swift_core.DEFAULT_MAPPING_SHEET)

        for w in (self.input_edit, self.output_edit, self.map_edit, self.sheet_edit):
            w.setStyleSheet("""
                QLineEdit{
                    background:#0F0F0F;
                    color:#EAEAEA;
                    border:1px solid #2E2E2E;
                    border-radius:8px;
                    padding:8px 10px;
                }
                QLineEdit:focus{ border:1px solid #FFB000; }
            """)

        btn_in = QPushButton("选择…")
        btn_out = QPushButton("选择…")
        btn_map = QPushButton("选择…")
        for b in (btn_in, btn_out, btn_map):
            b.setCursor(Qt.PointingHandCursor)
            b.setStyleSheet("""
                QPushButton{
                    background:#1E1E1E;
                    color:#EAEAEA;
                    border:1px solid #333;
                    border-radius:8px;
                    padding:8px 14px;
                }
                QPushButton:hover{ border:1px solid #FFB000; }
            """)

        row_in = QHBoxLayout()
        row_in.addWidget(self.input_edit, 1)
        row_in.addWidget(btn_in)

        row_out = QHBoxLayout()
        row_out.addWidget(self.output_edit, 1)
        row_out.addWidget(btn_out)

        row_map = QHBoxLayout()
        row_map.addWidget(self.map_edit, 1)
        row_map.addWidget(btn_map)

        form.addRow(QLabel("MSG文件夹："), self._wrap(row_in))
        form.addRow(QLabel("输出文件夹："), self._wrap(row_out))
        form.addRow(QLabel("Mapping 文件："), self._wrap(row_map))
        form.addRow(QLabel("Sheet 名称："), self.sheet_edit)

        layout.addWidget(group)

        # ------- run + progress -------
        action_row = QHBoxLayout()
        action_row.setSpacing(12)

        self.run_btn = QPushButton("▶ 运行")
        self.run_btn.setCursor(Qt.PointingHandCursor)
        self.run_btn.setFixedHeight(44)
        self.run_btn.setStyleSheet("""
            QPushButton{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #FFB000, stop:1 #FF7A00);
                color:#111;
                font-size:15px;
                font-weight:800;
                border:none;
                border-radius:10px;
                padding:10px 18px;
            }
            QPushButton:hover{ opacity:0.95; }
            QPushButton:disabled{
                background:#3A3A3A; color:#777;
            }
        """)

        self.progress = QProgressBar()
        self.progress.setFixedHeight(18)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setStyleSheet("""
            QProgressBar{
                background:#0F0F0F;
                border:1px solid #2E2E2E;
                border-radius:9px;
                color:#EAEAEA;
                text-align:center;
            }
            QProgressBar::chunk{
                border-radius:9px;
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #2DD4BF, stop:1 #FFB000);
            }
        """)

        action_row.addWidget(self.run_btn, 0)
        action_row.addWidget(self.progress, 1)
        layout.addLayout(action_row)

        self.status_label = QLabel("就绪。可直接运行或修改路径。")
        self.status_label.setStyleSheet("color:#B8B8B8; font-size:12px;")
        layout.addWidget(self.status_label)

        # ------- footer -------
        footer = QLabel("Designed by 余智秋 in Shanghai")
        footer.setAlignment(Qt.AlignCenter)
        footer_font = QFont()
        footer_font.setPointSize(9)
        footer.setFont(footer_font)
        footer.setStyleSheet("color:#D4AF37;")  # 金色
        layout.addStretch(1)
        layout.addWidget(footer)

        # ------- connections -------
        btn_in.clicked.connect(self.pick_input)
        btn_out.clicked.connect(self.pick_output)
        btn_map.clicked.connect(self.pick_mapping)
        self.run_btn.clicked.connect(self.run_job)

        # ------- dark theme for window background -------
        self.setStyleSheet("""
            QMainWindow { background:#0B0B0B; }
            QLabel { color:#EAEAEA; }
        """)

        self.worker = None

    def _wrap(self, layout: QHBoxLayout) -> QWidget:
        w = QWidget()
        w.setLayout(layout)
        return w

    def _msgbox(self, icon, title, text):
        mb = QMessageBox(self)
        mb.setIcon(icon)
        mb.setWindowTitle(title)
        mb.setText(text)
        mb.setStyleSheet(self.msgbox_style)
        mb.exec()

    def pick_input(self):
        d = QFileDialog.getExistingDirectory(self, "选择 MSG 文件夹", self.input_edit.text().strip() or os.getcwd())
        if d:
            self.input_edit.setText(d)

    def pick_output(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出文件夹", self.output_edit.text().strip() or os.getcwd())
        if d:
            self.output_edit.setText(d)

    def pick_mapping(self):
        f, _ = QFileDialog.getOpenFileName(self, "选择 Mapping Excel", self.map_edit.text().strip() or os.getcwd(), "Excel (*.xlsx *.xls)")
        if f:
            self.map_edit.setText(f)

    def run_job(self):
        input_dir = self.input_edit.text().strip()
        output_dir = self.output_edit.text().strip()
        mapping_file = self.map_edit.text().strip()
        sheet = self.sheet_edit.text().strip() or swift_core.DEFAULT_MAPPING_SHEET

        if not input_dir or not os.path.exists(input_dir):
            self._msgbox(QMessageBox.Warning, "路径错误", "MSG 文件夹不存在，请重新选择。")
            return
        if not output_dir:
            self._msgbox(QMessageBox.Warning, "路径错误", "输出文件夹不能为空。")
            return
        if not mapping_file or not os.path.exists(mapping_file):
            self._msgbox(QMessageBox.Warning, "路径错误", "Mapping 文件不存在，请重新选择。")
            return

        self.progress.setValue(0)
        self.progress.setFormat("0%")
        self.status_label.setText("启动任务中...")
        self.run_btn.setEnabled(False)

        self.worker = SwiftWorker(input_dir, output_dir, mapping_file, sheet)
        self.worker.progress.connect(self.on_progress)
        self.worker.status.connect(self.on_status)
        self.worker.finished_ok.connect(self.on_done)
        self.worker.failed.connect(self.on_failed)
        self.worker.start()

    def on_progress(self, done, total, filename):
        if total <= 0:
            self.progress.setValue(0)
            self.progress.setFormat("0%")
            return
        pct = int(done * 100 / total)
        self.progress.setValue(pct)
        self.progress.setFormat(f"{pct}%  ({done}/{total})")

    def on_status(self, msg):
        self.status_label.setText(msg)

    def on_done(self, output_path):
        self.run_btn.setEnabled(True)
        self.progress.setValue(100)
        self.progress.setFormat("100%  完成")

        self._msgbox(QMessageBox.Information, "完成", f"已完成处理。\n输出文件：\n{output_path}\n\n将自动打开 Excel。")

        # 自动打开 Excel（Windows）
        try:
            if os.path.exists(output_path):
                os.startfile(output_path)
        except Exception as e:
            self._msgbox(QMessageBox.Warning, "打开失败", f"无法自动打开文件：{e}")

    def on_failed(self, err):
        self.run_btn.setEnabled(True)
        self.status_label.setText("运行失败，请查看错误。")
        self._msgbox(QMessageBox.Critical, "运行失败", err)


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
