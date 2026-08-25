import sys
import os
from pathlib import Path

import pandas as pd
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter

from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QFileDialog, QSpinBox, QVBoxLayout, QHBoxLayout, QMessageBox,
    QProgressDialog
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QPixmap, QPainter, QColor, QPalette, QBrush
from match_color import color_match_pindou


# ==================== 主窗口 ====================
class PixelExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('拼豆图纸色号表格生成器')
        self.setFixedSize(620, 480)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint)

        # 生成像素风格背景
        self.set_pixel_background()

        # 创建界面
        self.init_ui()

        # 设置样式表（蓝粉配色）
        self.setStyleSheet(self.get_stylesheet())

        # 记录图片路径，用于自动生成输出路径
        self.last_image_path = ""

    def set_pixel_background(self):
        """动态生成一个像素风格的平铺背景图"""
        pixmap = QPixmap(32, 32)
        pixmap.fill(QColor('#F8E8F8'))  # 浅粉

        painter = QPainter(pixmap)
        # 绘制像素色块
        colors = [
            ('#FFD6E0', 0, 0), ('#B3D9FF', 16, 0),
            ('#FFB3C6', 8, 8), ('#99CCFF', 24, 8),
            ('#F0C0D0', 0, 16), ('#C0E0FF', 16, 16),
            ('#FFC8D6', 8, 24), ('#A8D4FF', 24, 24)
        ]
        for color, x, y in colors:
            painter.fillRect(x, y, 8, 8, QColor(color))
        painter.end()

        # 设置为背景（平铺）
        palette = self.palette()
        palette.setBrush(QPalette.Window, QBrush(pixmap))
        self.setPalette(palette)
        self.setAutoFillBackground(True)

    def init_ui(self):
        # 主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(20)

        # 标题
        title = QLabel('拼豆色号 Excel 生成器')
        title.setAlignment(Qt.AlignCenter)
        title.setObjectName('titleLabel')
        main_layout.addWidget(title)

        # ---------- 第一行：色号表格 ----------
        row1 = QHBoxLayout()
        lbl_csv = QLabel('色号表格:')
        lbl_csv.setFixedWidth(80)
        self.edit_csv = QLineEdit()
        self.edit_csv.setReadOnly(True)
        self.edit_csv.setPlaceholderText('选择色号表文件（CSV/Excel）')
        btn_csv = QPushButton('浏览')
        btn_csv.clicked.connect(self.browse_csv)
        row1.addWidget(lbl_csv)
        row1.addWidget(self.edit_csv, 1)
        row1.addWidget(btn_csv)
        main_layout.addLayout(row1)

        # ---------- 第二行：图片 ----------
        row2 = QHBoxLayout()
        lbl_img = QLabel('图片:')
        lbl_img.setFixedWidth(80)
        self.edit_img = QLineEdit()
        self.edit_img.setReadOnly(True)
        self.edit_img.setPlaceholderText('选择要转换的图片')
        btn_img = QPushButton('浏览')
        btn_img.clicked.connect(self.browse_image)
        row2.addWidget(lbl_img)
        row2.addWidget(self.edit_img, 1)
        row2.addWidget(btn_img)
        main_layout.addLayout(row2)

        # ---------- 第三行：输出路径 ----------
        row3 = QHBoxLayout()
        lbl_out = QLabel('输出路径:')
        lbl_out.setFixedWidth(80)
        self.edit_out = QLineEdit()
        self.edit_out.setPlaceholderText('选择图片后自动生成，也可手动修改')
        btn_out = QPushButton('浏览')
        btn_out.clicked.connect(self.browse_output)
        row3.addWidget(lbl_out)
        row3.addWidget(self.edit_out, 1)
        row3.addWidget(btn_out)
        main_layout.addLayout(row3)

        # ---------- 第四行：单元格大小 ----------
        row4 = QHBoxLayout()
        lbl_size = QLabel('单元格大小:')
        lbl_size.setFixedWidth(80)
        self.spin_size = QSpinBox()
        self.spin_size.setRange(5, 100)
        self.spin_size.setValue(30)
        self.spin_size.setSuffix(' px')
        row4.addWidget(lbl_size)
        row4.addWidget(self.spin_size)
        row4.addStretch(1)
        main_layout.addLayout(row4)

        # ---------- 生成按钮 ----------
        self.btn_generate = QPushButton('生成 Excel')
        self.btn_generate.setFixedHeight(50)
        self.btn_generate.clicked.connect(self.generate_excel)
        main_layout.addWidget(self.btn_generate)

        # 底部空白伸缩
        main_layout.addStretch(1)

    def browse_csv(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            '选择色号表',
            '',
            '色号表 (*.csv *.xlsx *.xls);;所有文件 (*.*)'
        )
        if file_path:
            self.edit_csv.setText(file_path)

    def browse_image(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            '选择图片',
            '',
            '图片文件 (*.png *.jpg *.jpeg *.bmp *.gif *.tiff *.webp);;所有文件 (*.*)'
        )
        if file_path:
            self.edit_img.setText(file_path)
            # 自动生成输出路径（同目录下同名 .xlsx）
            path = Path(file_path)
            default_out = str(path.with_suffix('.xlsx'))
            self.edit_out.setText(default_out)
            self.last_image_path = file_path

    def browse_output(self):
        # 默认文件名从当前输出路径提取，否则使用图片名
        default_name = os.path.basename(self.edit_out.text()) if self.edit_out.text() else 'output.xlsx'
        default_dir = os.path.dirname(self.edit_out.text()) if self.edit_out.text() else os.path.dirname(
            self.edit_img.text()) if self.edit_img.text() else ''
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            '保存 Excel 文件',
            os.path.join(default_dir, default_name),
            'Excel 文件 (*.xlsx);;所有文件 (*.*)'
        )
        if file_path:
            if not file_path.lower().endswith('.xlsx'):
                file_path += '.xlsx'
            self.edit_out.setText(file_path)

    def generate_excel(self):
        # 验证输入
        csv_path = self.edit_csv.text().strip()
        img_path = self.edit_img.text().strip()
        out_path = self.edit_out.text().strip()
        cell_size = self.spin_size.value()

        if not csv_path:
            QMessageBox.warning(self, '提示', '请选择色号表文件！')
            return
        if not img_path:
            QMessageBox.warning(self, '提示', '请选择要转换的图片！')
            return
        if not out_path:
            QMessageBox.warning(self, '提示', '请指定输出 Excel 文件路径！')
            return
        if not os.path.exists(csv_path):
            QMessageBox.critical(self, '错误', '色号表文件不存在！')
            return
        if not os.path.exists(img_path):
            QMessageBox.critical(self, '错误', '图片文件不存在！')
            return

        # 确保输出目录存在
        out_dir = os.path.dirname(out_path)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        # 禁用按钮，避免重复点击
        self.btn_generate.setEnabled(False)
        self.btn_generate.setText('正在生成...')

        try:
            # 这里直接调用处理函数（同步，大图可能会阻塞界面）
            # 如果需要真正异步，可以使用 QThread，但为简单演示先同步处理
            color_match_pindou(csv_path, img_path, out_path, cell_size)
            QMessageBox.information(self, '完成', f'Excel 文件已成功生成！\n{out_path}')
        except Exception as e:
            QMessageBox.critical(self, '错误', f'生成失败：{str(e)}')
        finally:
            self.btn_generate.setEnabled(True)
            self.btn_generate.setText('生成 Excel')

    def get_stylesheet(self):
        """返回蓝粉配色的样式表"""
        return """
        QWidget {
            font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
            font-size: 14px;
        }
        #titleLabel {
            font-size: 22px;
            font-weight: bold;
            color: #4A4A8A;
            background: transparent;
        }
        QLabel {
            color: #4A4A8A;
            background: transparent;
        }
        QLineEdit {
            background-color: #FFFFFF;
            border: 2px solid #B3D9FF;
            border-radius: 8px;
            padding: 6px 10px;
            color: #333333;
        }
        QLineEdit:focus {
            border-color: #FF8FAB;
        }
        QPushButton {
            background-color: #B3D9FF;
            border: none;
            border-radius: 8px;
            padding: 8px 16px;
            color: #FFFFFF;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #99CCFF;
        }
        QPushButton:pressed {
            background-color: #7FB3F0;
        }
        QPushButton:disabled {
            background-color: #D0D0D0;
        }
        QSpinBox {
            background-color: #FFFFFF;
            border: 2px solid #B3D9FF;
            border-radius: 8px;
            padding: 4px 8px;
        }
        #btn_generate {
            background-color: #FF8FAB;
        }
        #btn_generate:hover {
            background-color: #FF6B8A;
        }
        #btn_generate:pressed {
            background-color: #E65C7A;
        }
        """


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PixelExcelApp()
    window.show()
    sys.exit(app.exec())