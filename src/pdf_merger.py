#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
多文件合併PDF應用程式
功能：將多個Word或PDF文件合併為單一PDF，支援調整順序、生成目錄、添加頁碼，支援繁體中文
"""

import os
import sys
import time
import threading
import tempfile
import shutil
import platform
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                            QFileDialog, QProgressBar, QMessageBox, QVBoxLayout, 
                            QHBoxLayout, QWidget, QGroupBox, QListWidget, QCheckBox,
                            QComboBox, QSpinBox, QListWidgetItem, QAbstractItemView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QUrl
from PyQt5.QtGui import QFont, QDesktopServices, QIcon
import docx2pdf
from docx import Document
import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm  # 明確導入 cm 單位

# 根據不同作業系統設定字型
def setup_fonts():
    """設定繁體中文字型，根據不同作業系統選擇適當的字型"""
    system = platform.system()
    
    if system == 'Windows':
        # Windows系統使用內建的微軟正黑體
        try:
            font_path = "C:/Windows/Fonts/msjh.ttc"
            if os.path.exists(font_path):
                pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                return 'ChineseFont'
            else:
                # 備用字型
                font_path = "C:/Windows/Fonts/simsun.ttc"
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                    return 'ChineseFont'
                else:
                    # 如果找不到中文字型，使用預設字型
                    return 'Helvetica'
        except Exception as e:
            print(f"字型載入錯誤: {str(e)}")
            return 'Helvetica'
    
    elif system == 'Linux':
        # Linux系統嘗試使用Noto Sans CJK或文泉驛正黑
        try:
            font_paths = [
                '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
                '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
                '/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc'
            ]
            
            for font_path in font_paths:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                    return 'ChineseFont'
            
            # 如果找不到中文字型，使用預設字型
            return 'Helvetica'
        except Exception as e:
            print(f"字型載入錯誤: {str(e)}")
            return 'Helvetica'
    
    elif system == 'Darwin':  # macOS
        # macOS系統嘗試使用蘋方或黑體
        try:
            font_paths = [
                '/System/Library/Fonts/PingFang.ttc',
                '/System/Library/Fonts/STHeiti Light.ttc',
                '/Library/Fonts/Arial Unicode.ttf'
            ]
            
            for font_path in font_paths:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                    return 'ChineseFont'
            
            # 如果找不到中文字型，使用預設字型
            return 'Helvetica'
        except Exception as e:
            print(f"字型載入錯誤: {str(e)}")
            return 'Helvetica'
    
    else:
        # 其他作業系統使用預設字型
        return 'Helvetica'

# 設定字型
CHINESE_FONT = setup_fonts()

# 替代 docx2pdf 的函數，用於處理 Word 轉 PDF
def convert_word_to_pdf(word_path, pdf_path):
    """使用替代方法將 Word 轉換為 PDF，避免 RPC 錯誤"""
    system = platform.system()
    
    try:
        # 嘗試使用 docx2pdf
        docx2pdf.convert(word_path, pdf_path)
        return True
    except Exception as e:
        error_msg = str(e)
        print(f"docx2pdf 轉換失敗: {error_msg}")
        
        # 如果是 RPC 錯誤，嘗試使用替代方法
        if "遠端程序呼叫失敗" in error_msg or "RPC" in error_msg:
            try:
                # 在 Windows 上，嘗試使用 COM 自動化
                if system == 'Windows':
                    import win32com.client
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    doc = word.Documents.Open(word_path)
                    doc.SaveAs(pdf_path, FileFormat=17)  # 17 代表 PDF 格式
                    doc.Close()
                    word.Quit()
                    return True
                else:
                    # 非 Windows 系統，嘗試使用其他方法
                    return False
            except Exception as com_error:
                print(f"COM 自動化轉換失敗: {str(com_error)}")
                return False
        else:
            return False


class ConversionThread(QThread):
    """處理文件轉換與合併的執行緒，避免UI凍結"""
    progress_signal = pyqtSignal(int)
    status_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    
    def __init__(self, file_list, output_file, options):
        super().__init__()
        self.file_list = file_list
        self.output_file = output_file
        self.options = options
        self.temp_dir = tempfile.mkdtemp()
        
    def run(self):
        try:
            # 計算總步驟數
            total_steps = len(self.file_list) + 3  # 轉換 + 合併 + 目錄 + 頁碼
            current_step = 0
            
            # 步驟1: 將所有Word文件轉換為PDF
            pdf_files = []
            for i, file_info in enumerate(self.file_list):
                file_path = file_info['path']
                file_name = os.path.basename(file_path)
                
                self.status_signal.emit(f"處理檔案 {i+1}/{len(self.file_list)}: {file_name}")
                
                # 檢查是否為Word文件
                if file_path.lower().endswith(('.docx', '.doc')):
                    # 轉換為PDF
                    temp_pdf = os.path.join(self.temp_dir, f"temp_{i}.pdf")
                    self.status_signal.emit(f"轉換Word檔案: {file_name}")
                    
                    # 使用改進的轉換函數
                    success = convert_word_to_pdf(file_path, temp_pdf)
                    
                    if success and os.path.exists(temp_pdf):
                        pdf_files.append({
                            'path': temp_pdf,
                            'title': file_info['title']
                        })
                    else:
                        raise Exception(f"無法轉換 Word 檔案: {file_name}。請確認 Microsoft Word 已正確安裝且可以開啟此檔案。")
                else:
                    # 已經是PDF，直接添加
                    pdf_files.append({
                        'path': file_path,
                        'title': file_info['title']
                    })
                
                current_step += 1
                self.progress_signal.emit(int(current_step * 100 / total_steps))
            
            # 步驟2: 合併PDF文件
            self.status_signal.emit("合併PDF檔案...")
            
            # 如果需要目錄，先創建目錄頁
            toc_pdf = None
            if self.options['generate_toc']:
                toc_pdf = self.create_toc(pdf_files)
            
            # 合併所有PDF
            merger = PyPDF2.PdfMerger()
            
            # 如果有目錄，先添加目錄
            if toc_pdf:
                merger.append(toc_pdf)
            
            # 添加所有PDF文件
            for pdf_info in pdf_files:
                merger.append(pdf_info['path'])
            
            # 暫時保存合併後的PDF
            merged_pdf = os.path.join(self.temp_dir, "merged.pdf")
            merger.write(merged_pdf)
            merger.close()
            
            current_step += 1
            self.progress_signal.emit(int(current_step * 100 / total_steps))
            
            # 步驟3: 如果需要添加頁碼
            final_pdf = merged_pdf
            if self.options['add_page_numbers']:
                self.status_signal.emit("添加頁碼...")
                final_pdf = self.add_page_numbers(merged_pdf)
                
            current_step += 1
            self.progress_signal.emit(int(current_step * 100 / total_steps))
            
            # 步驟4: 複製到最終輸出位置
            shutil.copy2(final_pdf, self.output_file)
            
            current_step += 1
            self.progress_signal.emit(100)
            
            self.status_signal.emit("完成!")
            self.finished_signal.emit(self.output_file)
            
        except Exception as e:
            self.error_signal.emit(str(e))
        finally:
            # 清理臨時文件
            try:
                shutil.rmtree(self.temp_dir)
            except:
                pass
    
    def create_toc(self, pdf_files):
        """創建目錄頁"""
        toc_pdf_path = os.path.join(self.temp_dir, "toc.pdf")
        
        # 使用reportlab創建目錄
        c = canvas.Canvas(toc_pdf_path, pagesize=A4)
        c.setFont(CHINESE_FONT, 24)
        c.drawString(4*cm, 27*cm, "目錄")
        
        # 添加目錄項目
        y_position = 25*cm
        c.setFont(CHINESE_FONT, 12)
        
        # 計算起始頁碼
        start_page = 1  # 目錄頁本身
        
        for i, pdf_info in enumerate(pdf_files):
            title = pdf_info['title']
            
            # 檢查標題長度，過長則截斷
            if len(title) > 50:
                title = title[:47] + "..."
                
            # 繪製標題和頁碼
            c.drawString(2*cm, y_position, f"{i+1}. {title}")
            c.drawString(16*cm, y_position, f"{start_page + 1}")  # +1 因為目錄頁
            
            # 更新位置
            y_position -= 0.8*cm
            
            # 如果頁面空間不足，創建新頁
            if y_position < 2*cm:
                c.showPage()
                c.setFont(CHINESE_FONT, 12)
                y_position = 27*cm
            
            # 更新起始頁碼
            reader = PyPDF2.PdfReader(pdf_info['path'])
            start_page += len(reader.pages)
        
        c.save()
        return toc_pdf_path
    
    def add_page_numbers(self, pdf_path):
        """添加頁碼到PDF"""
        output_path = os.path.join(self.temp_dir, "with_page_numbers.pdf")
        
        # 讀取PDF
        reader = PyPDF2.PdfReader(pdf_path)
        writer = PyPDF2.PdfWriter()
        
        # 頁碼格式和位置
        page_format = self.options['page_number_format']
        start_number = self.options['start_page_number']
        
        # 處理每一頁
        for i, page in enumerate(reader.pages):
            # 獲取頁面尺寸
            page_width = float(page.mediabox.width)
            page_height = float(page.mediabox.height)
            
            # 創建一個臨時的PDF來繪製頁碼
            packet = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
            c = canvas.Canvas(packet.name, pagesize=(page_width, page_height))
            c.setFont(CHINESE_FONT, 10)
            
            # 格式化頁碼
            page_number = start_number + i
            if page_format == "數字":
                page_text = str(page_number)
            elif page_format == "羅馬數字":
                page_text = self.to_roman(page_number)
            else:  # 預設數字
                page_text = str(page_number)
            
            # 繪製頁碼在底部中央
            c.drawString(page_width/2 - 10, 20, page_text)
            c.save()
            
            # 讀取臨時PDF
            watermark = PyPDF2.PdfReader(packet.name)
            watermark_page = watermark.pages[0]
            
            # 合併頁面和頁碼
            page.merge_page(watermark_page)
            writer.add_page(page)
            
            # 清理臨時文件
            packet.close()
            os.unlink(packet.name)
        
        # 寫入結果
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        
        return output_path
    
    def to_roman(self, num):
        """將數字轉換為羅馬數字"""
        val = [
            1000, 900, 500, 400,
            100, 90, 50, 40,
            10, 9, 5, 4,
            1
        ]
        syms = [
            "M", "CM", "D", "CD",
            "C", "XC", "L", "XL",
            "X", "IX", "V", "IV",
            "I"
        ]
        roman_num = ''
        i = 0
        while num > 0:
            for _ in range(num // val[i]):
                roman_num += syms[i]
                num -= val[i]
            i += 1
        return roman_num


class PDFMerger(QMainWindow):
    """多文件合併PDF應用程式主視窗"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        """初始化使用者界面"""
        self.setWindowTitle('多文件合併PDF工具')
        self.setGeometry(300, 300, 700, 600)
        self.setMinimumSize(600, 500)
        
        # 設定字型以支援繁體中文
        self.font = QFont()
        if platform.system() == 'Windows':
            self.font.setFamily('Microsoft JhengHei')  # 微軟正黑體
        elif platform.system() == 'Darwin':  # macOS
            self.font.setFamily('PingFang TC')  # 蘋方
        else:
            self.font.setFamily('Noto Sans CJK TC')  # Noto Sans
        self.font.setPointSize(10)
        QApplication.setFont(self.font)
        
        # 主佈局
        main_layout = QVBoxLayout()
        
        # 檔案列表區域
        file_group = QGroupBox('檔案列表')
        file_layout = QVBoxLayout()
        
        # 檔案列表
        self.file_list = QListWidget()
        self.file_list.setDragDropMode(QAbstractItemView.InternalMove)  # 允許拖放調整順序
        
        # 檔案操作按鈕
        file_buttons_layout = QHBoxLayout()
        self.add_file_btn = QPushButton('添加檔案')
        self.add_file_btn.clicked.connect(self.add_files)
        
        self.remove_file_btn = QPushButton('移除檔案')
        self.remove_file_btn.clicked.connect(self.remove_file)
        
        self.move_up_btn = QPushButton('上移')
        self.move_up_btn.clicked.connect(self.move_file_up)
        
        self.move_down_btn = QPushButton('下移')
        self.move_down_btn.clicked.connect(self.move_file_down)
        
        file_buttons_layout.addWidget(self.add_file_btn)
        file_buttons_layout.addWidget(self.remove_file_btn)
        file_buttons_layout.addWidget(self.move_up_btn)
        file_buttons_layout.addWidget(self.move_down_btn)
        
        file_layout.addWidget(self.file_list)
        file_layout.addLayout(file_buttons_layout)
        file_group.setLayout(file_layout)
        
        # 合併選項區域
        options_group = QGroupBox('合併選項')
        options_layout = QVBoxLayout()
        
        # 目錄選項
        toc_layout = QHBoxLayout()
        self.generate_toc_cb = QCheckBox('生成目錄')
        self.generate_toc_cb.setChecked(True)
        toc_layout.addWidget(self.generate_toc_cb)
        
        # 頁碼選項
        page_number_layout = QHBoxLayout()
        self.add_page_numbers_cb = QCheckBox('添加頁碼')
        self.add_page_numbers_cb.setChecked(True)
        self.add_page_numbers_cb.stateChanged.connect(self.toggle_page_options)
        
        self.page_format_label = QLabel('頁碼格式:')
        self.page_format_combo = QComboBox()
        self.page_format_combo.addItems(['數字', '羅馬數字'])
        
        self.start_page_label = QLabel('起始頁碼:')
        self.start_page_spin = QSpinBox()
        self.start_page_spin.setMinimum(1)
        self.start_page_spin.setValue(1)
        
        page_number_layout.addWidget(self.add_page_numbers_cb)
        page_number_layout.addWidget(self.page_format_label)
        page_number_layout.addWidget(self.page_format_combo)
        page_number_layout.addWidget(self.start_page_label)
        page_number_layout.addWidget(self.start_page_spin)
        
        options_layout.addLayout(toc_layout)
        options_layout.addLayout(page_number_layout)
        options_group.setLayout(options_layout)
        
        # 輸出選項區域
        output_group = QGroupBox('輸出設定')
        output_layout = QHBoxLayout()
        
        self.output_label = QLabel('輸出檔案:')
        self.output_path = QLabel('尚未選擇輸出檔案')
        self.output_btn = QPushButton('選擇...')
        self.output_btn.clicked.connect(self.select_output)
        
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(self.output_path, 1)  # 1表示可伸展
        output_layout.addWidget(self.output_btn)
        
        output_group.setLayout(output_layout)
        
        # 進度區域
        progress_layout = QVBoxLayout()
        self.status_label = QLabel('準備就緒')
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        
        progress_layout.addWidget(self.status_label)
        progress_layout.addWidget(self.progress_bar)
        
        # 操作按鈕區域
        button_layout = QHBoxLayout()
        self.merge_btn = QPushButton('開始合併')
        self.merge_btn.clicked.connect(self.start_merge)
        self.merge_btn.setEnabled(False)
        
        self.open_btn = QPushButton('開啟檔案')
        self.open_btn.clicked.connect(self.open_output_file)
        self.open_btn.setEnabled(False)
        
        button_layout.addWidget(self.merge_btn)
        button_layout.addWidget(self.open_btn)
        
        # 添加所有元件到主佈局
        main_layout.addWidget(file_group, 3)  # 3表示佔比較大
        main_layout.addWidget(options_group, 1)
        main_layout.addWidget(output_group, 1)
        main_layout.addLayout(progress_layout)
        main_layout.addLayout(button_layout)
        
        # 設定主視窗
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        
        # 初始化變數
        self.files = []
        self.output_file = ''
        
        # 初始化頁碼選項狀態
        self.toggle_page_options()
    
    def add_files(self):
        """添加檔案到列表"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, '選擇檔案', '', 'Word檔案 (*.docx *.doc);;PDF檔案 (*.pdf);;所有檔案 (*.*)'
        )
        
        if file_paths:
            for file_path in file_paths:
                # 檢查檔案是否已在列表中
                if any(file['path'] == file_path for file in self.files):
                    continue
                
                # 獲取檔案名稱作為標題
                title = os.path.splitext(os.path.basename(file_path))[0]
                
                # 添加到內部列表
                self.files.append({
                    'path': file_path,
                    'title': title
                })
                
                # 添加到UI列表
                item = QListWidgetItem(f"{title} ({os.path.basename(file_path)})")
                self.file_list.addItem(item)
            
            # 如果有檔案且已選擇輸出路徑，啟用合併按鈕
            if self.files and self.output_file:
                self.merge_btn.setEnabled(True)
    
    def remove_file(self):
        """從列表中移除選中的檔案"""
        current_row = self.file_list.currentRow()
        if current_row >= 0:
            self.file_list.takeItem(current_row)
            self.files.pop(current_row)
            
            # 如果列表為空或未選擇輸出路徑，禁用合併按鈕
            if not self.files or not self.output_file:
                self.merge_btn.setEnabled(False)
    
    def move_file_up(self):
        """將選中的檔案上移"""
        current_row = self.file_list.currentRow()
        if current_row > 0:
            # 移動UI列表項目
            item = self.file_list.takeItem(current_row)
            self.file_list.insertItem(current_row - 1, item)
            self.file_list.setCurrentRow(current_row - 1)
            
            # 移動內部列表項目
            self.files.insert(current_row - 1, self.files.pop(current_row))
    
    def move_file_down(self):
        """將選中的檔案下移"""
        current_row = self.file_list.currentRow()
        if current_row < self.file_list.count() - 1:
            # 移動UI列表項目
            item = self.file_list.takeItem(current_row)
            self.file_list.insertItem(current_row + 1, item)
            self.file_list.setCurrentRow(current_row + 1)
            
            # 移動內部列表項目
            self.files.insert(current_row + 1, self.files.pop(current_row))
    
    def toggle_page_options(self):
        """切換頁碼選項的啟用狀態"""
        enabled = self.add_page_numbers_cb.isChecked()
        self.page_format_label.setEnabled(enabled)
        self.page_format_combo.setEnabled(enabled)
        self.start_page_label.setEnabled(enabled)
        self.start_page_spin.setEnabled(enabled)
    
    def select_output(self):
        """選擇輸出檔案路徑"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, '儲存合併PDF', '', 'PDF檔案 (*.pdf)'
        )
        
        if file_path:
            # 確保有.pdf副檔名
            if not file_path.lower().endswith('.pdf'):
                file_path += '.pdf'
                
            self.output_file = file_path
            self.output_path.setText(os.path.basename(file_path))
            
            # 如果有檔案，啟用合併按鈕
            if self.files:
                self.merge_btn.setEnabled(True)
    
    def start_merge(self):
        """開始合併PDF檔案"""
        if not self.files:
            QMessageBox.warning(self, '警告', '請先添加檔案')
            return
            
        if not self.output_file:
            QMessageBox.warning(self, '警告', '請選擇輸出檔案路徑')
            return
        
        # 獲取合併選項
        options = {
            'generate_toc': self.generate_toc_cb.isChecked(),
            'add_page_numbers': self.add_page_numbers_cb.isChecked(),
            'page_number_format': self.page_format_combo.currentText(),
            'start_page_number': self.start_page_spin.value()
        }
        
        # 禁用按鈕，避免重複操作
        self.merge_btn.setEnabled(False)
        self.add_file_btn.setEnabled(False)
        self.remove_file_btn.setEnabled(False)
        self.move_up_btn.setEnabled(False)
        self.move_down_btn.setEnabled(False)
        self.output_btn.setEnabled(False)
        
        # 開始合併執行緒
        self.conversion_thread = ConversionThread(self.files, self.output_file, options)
        self.conversion_thread.progress_signal.connect(self.update_progress)
        self.conversion_thread.status_signal.connect(self.update_status)
        self.conversion_thread.finished_signal.connect(self.merge_finished)
        self.conversion_thread.error_signal.connect(self.merge_error)
        self.conversion_thread.start()
    
    def update_progress(self, value):
        """更新進度條"""
        self.progress_bar.setValue(value)
    
    def update_status(self, status):
        """更新狀態標籤"""
        self.status_label.setText(status)
    
    def merge_finished(self, output_file):
        """合併完成處理"""
        # 重新啟用按鈕
        self.merge_btn.setEnabled(True)
        self.add_file_btn.setEnabled(True)
        self.remove_file_btn.setEnabled(True)
        self.move_up_btn.setEnabled(True)
        self.move_down_btn.setEnabled(True)
        self.output_btn.setEnabled(True)
        self.open_btn.setEnabled(True)
        
        QMessageBox.information(
            self, '完成', f'檔案已成功合併!\n儲存於: {output_file}'
        )
    
    def merge_error(self, error_msg):
        """合併錯誤處理"""
        self.status_label.setText('合併失敗!')
        self.progress_bar.setValue(0)
        
        # 重新啟用按鈕
        self.merge_btn.setEnabled(True)
        self.add_file_btn.setEnabled(True)
        self.remove_file_btn.setEnabled(True)
        self.move_up_btn.setEnabled(True)
        self.move_down_btn.setEnabled(True)
        self.output_btn.setEnabled(True)
        
        QMessageBox.critical(self, '錯誤', f'合併過程中發生錯誤:\n{error_msg}')
    
    def open_output_file(self):
        """開啟生成的PDF檔案"""
        if os.path.exists(self.output_file):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.output_file))
        else:
            QMessageBox.warning(self, '警告', '找不到輸出檔案')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    merger = PDFMerger()
    merger.show()
    sys.exit(app.exec_())
