#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Word與PDF轉換與合併應用程式啟動器
"""

import os
import sys
import subprocess
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                            QVBoxLayout, QHBoxLayout, QWidget)
from PyQt5.QtGui import QFont, QPixmap, QIcon
from PyQt5.QtCore import Qt

class LauncherApp(QMainWindow):
    """應用程式啟動器主視窗"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        """初始化使用者界面"""
        self.setWindowTitle('Word與PDF轉換工具')
        self.setGeometry(300, 300, 500, 300)
        self.setMinimumSize(400, 250)
        
        # 設定字型以支援繁體中文
        self.font = QFont()
        self.font.setFamily('Microsoft JhengHei')  # 微軟正黑體
        self.font.setPointSize(10)
        QApplication.setFont(self.font)
        
        # 主佈局
        main_layout = QVBoxLayout()
        
        # 標題
        title_label = QLabel('Word與PDF轉換與合併工具')
        title_font = QFont()
        title_font.setFamily('Microsoft JhengHei')
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        
        # 說明
        desc_label = QLabel('請選擇要使用的功能:')
        desc_label.setAlignment(Qt.AlignCenter)
        
        # 按鈕區域
        button_layout = QHBoxLayout()
        
        # Word轉PDF按鈕
        self.word_to_pdf_btn = QPushButton('Word轉PDF')
        self.word_to_pdf_btn.setMinimumHeight(80)
        self.word_to_pdf_btn.clicked.connect(self.launch_word_to_pdf)
        
        # 多文件合併按鈕
        self.pdf_merger_btn = QPushButton('多文件合併PDF')
        self.pdf_merger_btn.setMinimumHeight(80)
        self.pdf_merger_btn.clicked.connect(self.launch_pdf_merger)
        
        button_layout.addWidget(self.word_to_pdf_btn)
        button_layout.addWidget(self.pdf_merger_btn)
        
        # 版權資訊
        copyright_label = QLabel('© 2025 Word與PDF轉換工具')
        copyright_label.setAlignment(Qt.AlignCenter)
        
        # 添加所有元件到主佈局
        main_layout.addWidget(title_label)
        main_layout.addSpacing(20)
        main_layout.addWidget(desc_label)
        main_layout.addSpacing(10)
        main_layout.addLayout(button_layout)
        main_layout.addStretch(1)
        main_layout.addWidget(copyright_label)
        
        # 設定主視窗
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
    
    def launch_word_to_pdf(self):
        """啟動Word轉PDF應用程式"""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        word_to_pdf_path = os.path.join(script_dir, 'word_to_pdf.py')
        
        try:
            subprocess.Popen([sys.executable, word_to_pdf_path])
        except Exception as e:
            print(f"啟動Word轉PDF應用程式時發生錯誤: {str(e)}")
    
    def launch_pdf_merger(self):
        """啟動多文件合併PDF應用程式"""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        pdf_merger_path = os.path.join(script_dir, 'pdf_merger.py')
        
        try:
            subprocess.Popen([sys.executable, pdf_merger_path])
        except Exception as e:
            print(f"啟動多文件合併PDF應用程式時發生錯誤: {str(e)}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    launcher = LauncherApp()
    launcher.show()
    sys.exit(app.exec_())
