#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

"""
Word與PDF轉換與合併應用程式 - 整合版
功能：
1. 將Word文件轉換為PDF格式
2. 將PDF文件轉換為Word格式
3. 將多個Word或PDF文件合併為單一PDF檔案
"""

import os
import sys
import time
import threading
import tempfile
import shutil
import platform
import subprocess
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                            QFileDialog, QProgressBar, QMessageBox, QVBoxLayout, 
                            QHBoxLayout, QWidget, QGroupBox, QListWidget, QCheckBox,
                            QComboBox, QSpinBox, QListWidgetItem, QAbstractItemView,
                            QTabWidget, QTextEdit, QInputDialog, QLineEdit)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QUrl, QMetaObject, pyqtSlot, Q_ARG, Q_RETURN_ARG
from PyQt5.QtGui import QFont, QDesktopServices, QIcon
from src.core_converter import CHINESE_FONT, convert_word_to_pdf, convert_pdf_to_word, to_roman, create_toc, add_page_numbers, prepare_pdf_reader_core, merge_pdfs_core

# --- 常數 ---
USER_CANCELLED_PASSWORD = "USER_CANCELLED_MAGIC_STRING"






class PdfToWordThread(QThread):
    """處理PDF轉Word的執行緒，避免UI凍結，使用智能進度估算"""
    progress_signal = pyqtSignal(int)
    status_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)

    def __init__(self, input_file, output_file, quality='default'):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.quality = quality
        self.converting = False
        self.progress_thread = None

    def _estimate_conversion_time(self):
        """根據PDF頁數和檔案大小估算轉換時間"""
        try:
            # 讀取PDF資訊
            reader = pypdf.PdfReader(self.input_file)
            page_count = len(reader.pages)
            file_size_mb = os.path.getsize(self.input_file) / (1024 * 1024)

            # 估算公式：基礎時間 + 頁數相關 + 檔案大小相關
            # 每頁約0.5秒，每MB約1秒
            estimated_time = 2 + (page_count * 0.5) + (file_size_mb * 1)
            return max(3, min(estimated_time, 120))  # 限制在3-120秒之間
        except:
            return 10  # 預設10秒

    def _update_progress_smooth(self, estimated_time):
        """平滑更新進度條"""
        start_progress = 10
        end_progress = 85
        progress_range = end_progress - start_progress

        # 計算更新間隔（每0.2秒更新一次）
        update_interval = 0.2
        total_updates = int(estimated_time / update_interval)
        progress_per_update = progress_range / total_updates

        current_progress = start_progress

        while self.converting and current_progress < end_progress:
            time.sleep(update_interval)
            if not self.converting:
                break
            current_progress += progress_per_update
            self.progress_signal.emit(int(min(current_progress, end_progress)))

    def run(self):
        try:
            # 階段1：準備轉換
            self.status_signal.emit("準備轉換...")
            self.progress_signal.emit(5)

            # 讀取PDF資訊並估算時間
            try:
                reader = pypdf.PdfReader(self.input_file)
                page_count = len(reader.pages)
                file_size = os.path.getsize(self.input_file) / (1024 * 1024)
                estimated_time = self._estimate_conversion_time()

                self.status_signal.emit(f"分析PDF... ({page_count} 頁, {file_size:.1f} MB)")
            except:
                page_count = 0
                estimated_time = 10
                self.status_signal.emit("分析PDF...")

            self.progress_signal.emit(10)

            # 階段2：開始轉換
            quality_text = {"default": "預設", "high": "高品質", "text": "純文字"}.get(self.quality, "預設")
            self.status_signal.emit(f"轉換中 ({quality_text})... 預計 {int(estimated_time)} 秒")
            self.converting = True

            # 啟動平滑進度更新
            self.progress_thread = threading.Thread(
                target=self._update_progress_smooth,
                args=(estimated_time,),
                daemon=True
            )
            self.progress_thread.start()

            # 執行轉換（這裡可以根據 quality 參數調整轉換方式）
            success = convert_pdf_to_word(self.input_file, self.output_file)

            # 停止進度更新
            self.converting = False
            if self.progress_thread:
                self.progress_thread.join(timeout=1)

            # 階段3：完成
            self.progress_signal.emit(90)
            self.status_signal.emit("完成轉換，正在驗證...")

            if success:
                self.progress_signal.emit(95)
                time.sleep(0.3)

                self.status_signal.emit("轉換完成!")
                self.progress_signal.emit(100)
                self.finished_signal.emit(self.output_file)
            else:
                self.status_signal.emit("轉換失敗!")
                self.error_signal.emit("PDF轉Word失敗，請確認PDF檔案是否受保護或損壞。")
        except Exception as e:
            self.converting = False
            self.status_signal.emit("轉換出錯!")
            self.error_signal.emit(str(e))


class WordToPdfThread(QThread):
    """處理Word轉PDF的執行緒，避免UI凍結，使用智能進度估算"""
    progress_signal = pyqtSignal(int)
    status_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)

    def __init__(self, input_file, output_file):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.converting = False
        self.progress_thread = None

    def _estimate_conversion_time(self):
        """根據檔案大小估算轉換時間（秒）"""
        try:
            file_size_mb = os.path.getsize(self.input_file) / (1024 * 1024)
            # 估算公式：基礎時間 + 檔案大小相關時間
            # 小檔案（<1MB）約3秒，每增加1MB約增加2秒
            estimated_time = 3 + (file_size_mb * 2)
            return max(3, min(estimated_time, 60))  # 限制在3-60秒之間
        except:
            return 10  # 預設10秒

    def _update_progress_smooth(self, estimated_time):
        """平滑更新進度條"""
        start_progress = 10
        end_progress = 85
        progress_range = end_progress - start_progress

        # 計算更新間隔（每0.2秒更新一次）
        update_interval = 0.2
        total_updates = int(estimated_time / update_interval)
        progress_per_update = progress_range / total_updates

        current_progress = start_progress

        while self.converting and current_progress < end_progress:
            time.sleep(update_interval)
            if not self.converting:
                break
            current_progress += progress_per_update
            self.progress_signal.emit(int(min(current_progress, end_progress)))

    def run(self):
        try:
            # 階段1：準備轉換
            self.status_signal.emit("準備轉換...")
            self.progress_signal.emit(5)

            # 估算轉換時間
            estimated_time = self._estimate_conversion_time()
            file_size = os.path.getsize(self.input_file) / (1024 * 1024)

            self.status_signal.emit(f"分析檔案... ({file_size:.1f} MB)")
            self.progress_signal.emit(10)

            # 階段2：開始轉換，啟動進度更新執行緒
            self.status_signal.emit(f"轉換中... (預計 {int(estimated_time)} 秒)")
            self.converting = True

            # 在背景執行緒中平滑更新進度
            self.progress_thread = threading.Thread(
                target=self._update_progress_smooth,
                args=(estimated_time,),
                daemon=True
            )
            self.progress_thread.start()

            # 執行轉換
            success = convert_word_to_pdf(self.input_file, self.output_file)

            # 停止進度更新
            self.converting = False
            if self.progress_thread:
                self.progress_thread.join(timeout=1)

            # 階段3：完成轉換
            self.progress_signal.emit(90)
            self.status_signal.emit("完成轉換，正在驗證...")

            # 檢查結果
            if success and os.path.exists(self.output_file) and os.path.getsize(self.output_file) > 0:
                self.progress_signal.emit(95)
                time.sleep(0.3)  # 短暫延遲讓使用者看到完成前的狀態

                self.status_signal.emit("轉換完成!")
                self.progress_signal.emit(100)
                self.finished_signal.emit(self.output_file)
            else:
                self.status_signal.emit("轉換失敗!")
                self.error_signal.emit("轉換失敗，請確認Microsoft Word已正確安裝且可以開啟此檔案。")
        except Exception as e:
            self.converting = False
            self.status_signal.emit("轉換出錯!")
            self.error_signal.emit(str(e))


class MergePdfThread(QThread):
    """處理文件轉換與合併的執行緒，避免UI凍結"""
    progress_signal = pyqtSignal(int)
    status_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    
    def __init__(self, file_list, output_file, options, main_window):
        super().__init__()
        self.file_list = file_list
        self.output_file = output_file
        self.options = options
        # 確保主窗口對象引用是线程安全的
        self.main_window = main_window
        
    def run(self):
        try:
            # The total_steps calculation might need adjustment depending on how merge_pdfs_core reports progress
            # For now, let's simplify the progress reporting in the GUI thread.
            self.status_signal.emit("準備檔案中...")
            self.progress_signal.emit(5)

            temp_dir = tempfile.mkdtemp() # Create temp_dir here

            def gui_password_callback(filename):
                self.main_window.password_requested.emit(filename)
                return self.main_window.wait_for_password()

            merge_pdfs_core(
                file_list=self.file_list,
                output_file=self.output_file,
                options=self.options,
                temp_dir=temp_dir,
                password_callback=gui_password_callback
            )

            self.progress_signal.emit(90)
            self.status_signal.emit("完成轉換，正在驗證...")

            # Check result (merge_pdfs_core handles the actual merging and copying)
            if os.path.exists(self.output_file) and os.path.getsize(self.output_file) > 0:
                self.progress_signal.emit(95)
                time.sleep(0.3)

                self.status_signal.emit("轉換完成!")
                self.progress_signal.emit(100)
                self.finished_signal.emit(self.output_file)
            else:
                self.status_signal.emit("合併失敗!")
                self.error_signal.emit("PDF合併失敗，請確認檔案是否損壞或內容有誤。")

        except Exception as e:
            self.status_signal.emit("合併出錯!")
            self.error_signal.emit(str(e))
        finally:
            try:
                if 'temp_dir' in locals() and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
            except Exception as e:
                print(f"清理臨時目錄失敗: {e}")
    



class PdfToWordTab(QWidget):
    """PDF轉Word標籤頁"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
        
    def init_ui(self):
        """初始化使用者界面"""
        # 主佈局
        main_layout = QVBoxLayout()
        
        # 檔案選擇區域
        file_group = QGroupBox('檔案選擇')
        file_layout = QVBoxLayout()
        
        input_layout = QHBoxLayout()
        self.input_label = QLabel('PDF檔案:')
        self.input_path = QTextEdit()
        self.input_path.setReadOnly(True)
        self.input_path.setMaximumHeight(60)
        self.browse_btn = QPushButton('瀏覽...')
        self.browse_btn.clicked.connect(self.browse_file)
        
        input_layout.addWidget(self.input_label)
        input_layout.addWidget(self.input_path)
        input_layout.addWidget(self.browse_btn)
        
        output_layout = QHBoxLayout()
        self.output_label = QLabel('儲存位置:')
        self.output_path = QTextEdit()
        self.output_path.setReadOnly(True)
        self.output_path.setMaximumHeight(60)
        self.save_btn = QPushButton('選擇...')
        self.save_btn.clicked.connect(self.save_file)
        
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(self.output_path)
        output_layout.addWidget(self.save_btn)
        
        file_layout.addLayout(input_layout)
        file_layout.addLayout(output_layout)
        file_group.setLayout(file_layout)
        
        # 檔案資訊區域
        info_group = QGroupBox('檔案資訊')
        info_layout = QVBoxLayout()
        self.file_info = QTextEdit()
        self.file_info.setReadOnly(True)
        info_layout.addWidget(self.file_info)
        info_group.setLayout(info_layout)
        
        # 轉換選項區域
        options_group = QGroupBox('轉換選項')
        options_layout = QVBoxLayout()
        self.quality_label = QLabel("轉換品質:")
        self.quality_combo = QComboBox()
        self.quality_combo.addItems(['預設', '高品質 (保留更多格式)', '純文字'])
        options_layout.addWidget(self.quality_label)
        options_layout.addWidget(self.quality_combo)
        options_group.setLayout(options_layout)
        
        # 進度條區域
        progress_layout = QVBoxLayout()
        self.status_label = QLabel('轉換進度:')
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.status_label)
        progress_layout.addWidget(self.progress_bar)
        
        # 操作按鈕區域
        button_layout = QHBoxLayout()
        self.convert_btn = QPushButton('開始轉換')
        self.convert_btn.clicked.connect(self.start_conversion)
        self.convert_btn.setEnabled(False)
        
        self.open_btn = QPushButton('開啟檔案')
        self.open_btn.clicked.connect(self.open_output_file)
        self.open_btn.setEnabled(False)
        
        button_layout.addWidget(self.convert_btn)
        button_layout.addWidget(self.open_btn)
        
        # 添加所有元件到主佈局
        main_layout.addWidget(file_group)
        main_layout.addWidget(info_group)
        main_layout.addWidget(options_group)
        main_layout.addLayout(progress_layout)
        main_layout.addLayout(button_layout)
        
        self.setLayout(main_layout)
        
        # 初始化變數
        self.input_file = ''
        self.output_file = ''
        
    def browse_file(self):
        """選擇PDF檔案"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, '選擇PDF檔案', '', 'PDF檔案 (*.pdf)'
        )
        
        if file_path:
            self.input_file = file_path
            self.input_path.setText(file_path)
            
            # 自動設定輸出檔案路徑
            dir_name = os.path.dirname(file_path)
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            self.output_file = os.path.join(dir_name, f"{base_name}.docx")
            self.output_path.setText(self.output_file)
            
            # 顯示檔案資訊
            self.show_file_info(file_path)
            
            # 啟用轉換按鈕
            self.convert_btn.setEnabled(True)
    
    def save_file(self):
        """選擇Word儲存位置"""
        if not self.input_file:
            QMessageBox.warning(self, '警告', '請先選擇PDF檔案')
            return
            
        base_name = os.path.splitext(os.path.basename(self.input_file))[0]
        file_path, _ = QFileDialog.getSaveFileName(
            self, '儲存Word檔案', f"{base_name}.docx", 'Word檔案 (*.docx)'
        )
        
        if file_path:
            self.output_file = file_path
            self.output_path.setText(file_path)
    
    def show_file_info(self, file_path):
        """顯示PDF檔案資訊"""
        try:
            reader = pypdf.PdfReader(file_path)
            info = f"檔案名稱: {os.path.basename(file_path)}\n"
            info += f"檔案大小: {os.path.getsize(file_path) / 1024:.2f} KB\n"
            info += f"頁數: {len(reader.pages)}\n"
            
            if reader.metadata:
                info += f"標題: {reader.metadata.title}\n"
                info += f"作者: {reader.metadata.author}\n"
            
            self.file_info.setText(info)
        except Exception as e:
            self.file_info.setText(f"無法讀取檔案資訊: {str(e)}")
    
    def start_conversion(self):
        """開始轉換PDF到Word"""
        if not self.input_file or not self.output_file:
            QMessageBox.warning(self, '警告', '請選擇輸入和輸出檔案')
            return

        # 禁用按鈕，避免重複操作
        self.convert_btn.setEnabled(False)
        self.browse_btn.setEnabled(False)
        self.save_btn.setEnabled(False)

        # 取得品質選項
        quality_map = {
            '預設': 'default',
            '高品質 (保留更多格式)': 'high',
            '純文字': 'text'
        }
        quality = quality_map.get(self.quality_combo.currentText(), 'default')

        # 開始轉換執行緒，傳遞品質參數
        self.conversion_thread = PdfToWordThread(self.input_file, self.output_file, quality)
        self.conversion_thread.progress_signal.connect(self.update_progress)
        self.conversion_thread.status_signal.connect(self.update_status)
        self.conversion_thread.finished_signal.connect(self.conversion_finished)
        self.conversion_thread.error_signal.connect(self.conversion_error)
        self.conversion_thread.start()
    
    def update_progress(self, value):
        """更新進度條"""
        self.progress_bar.setValue(value)
    
    def update_status(self, status):
        """更新狀態標籤"""
        self.status_label.setText(status)
    
    def conversion_finished(self, output_file):
        """轉換完成處理"""
        self.status_label.setText('轉換完成!')
        
        # 重新啟用按鈕
        self.convert_btn.setEnabled(True)
        self.browse_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        self.open_btn.setEnabled(True)
        
        QMessageBox.information(
            self, '完成', f'PDF檔案已成功轉換為Word!\n儲存於: {output_file}'
        )
    
    def conversion_error(self, error_msg):
        """轉換錯誤處理"""
        self.status_label.setText('轉換失敗!')
        self.progress_bar.setValue(0)
        
        # 重新啟用按鈕
        self.convert_btn.setEnabled(True)
        self.browse_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        
        QMessageBox.critical(self, '錯誤', f'轉換過程中發生錯誤:\n{error_msg}')
    
    def open_output_file(self):
        """開啟生成的Word檔案"""
        if os.path.exists(self.output_file):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.output_file))
        else:
            QMessageBox.warning(self, '警告', '找不到輸出檔案')


class WordToPdfTab(QWidget):
    """Word轉PDF標籤頁"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
        
    def init_ui(self):
        """初始化使用者界面"""
        # 主佈局
        main_layout = QVBoxLayout()
        
        # 檔案選擇區域
        file_group = QGroupBox('檔案選擇')
        file_layout = QVBoxLayout()
        
        input_layout = QHBoxLayout()
        self.input_label = QLabel('Word檔案:')
        self.input_path = QTextEdit()
        self.input_path.setReadOnly(True)
        self.input_path.setMaximumHeight(60)
        self.browse_btn = QPushButton('瀏覽...')
        self.browse_btn.clicked.connect(self.browse_file)
        
        input_layout.addWidget(self.input_label)
        input_layout.addWidget(self.input_path)
        input_layout.addWidget(self.browse_btn)
        
        output_layout = QHBoxLayout()
        self.output_label = QLabel('儲存位置:')
        self.output_path = QTextEdit()
        self.output_path.setReadOnly(True)
        self.output_path.setMaximumHeight(60)
        self.save_btn = QPushButton('選擇...')
        self.save_btn.clicked.connect(self.save_file)
        
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(self.output_path)
        output_layout.addWidget(self.save_btn)
        
        file_layout.addLayout(input_layout)
        file_layout.addLayout(output_layout)
        file_group.setLayout(file_layout)
        
        # 檔案資訊區域
        info_group = QGroupBox('檔案資訊')
        info_layout = QVBoxLayout()
        self.file_info = QTextEdit()
        self.file_info.setReadOnly(True)
        info_layout.addWidget(self.file_info)
        info_group.setLayout(info_layout)
        
        # 轉換選項區域
        options_group = QGroupBox('轉換選項')
        options_layout = QVBoxLayout()
        self.keep_format_cb = QCheckBox('保留格式、圖片與表格')
        self.keep_format_cb.setChecked(True)
        self.keep_format_cb.setEnabled(False)  # 預設啟用且不可更改
        options_layout.addWidget(self.keep_format_cb)
        options_group.setLayout(options_layout)
        
        # 進度條區域
        progress_layout = QVBoxLayout()
        self.status_label = QLabel('轉換進度:')
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.status_label)
        progress_layout.addWidget(self.progress_bar)
        
        # 操作按鈕區域
        button_layout = QHBoxLayout()
        self.convert_btn = QPushButton('開始轉換')
        self.convert_btn.clicked.connect(self.start_conversion)
        self.convert_btn.setEnabled(False)
        
        self.open_btn = QPushButton('開啟檔案')
        self.open_btn.clicked.connect(self.open_output_file)
        self.open_btn.setEnabled(False)
        
        button_layout.addWidget(self.convert_btn)
        button_layout.addWidget(self.open_btn)
        
        # 添加所有元件到主佈局
        main_layout.addWidget(file_group)
        main_layout.addWidget(info_group)
        main_layout.addWidget(options_group)
        main_layout.addLayout(progress_layout)
        main_layout.addLayout(button_layout)
        
        self.setLayout(main_layout)
        
        # 初始化變數
        self.input_file = ''
        self.output_file = ''
        
    def browse_file(self):
        """選擇Word檔案"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, '選擇Word檔案', '', 'Word檔案 (*.docx *.doc)'
        )
        
        if file_path:
            self.input_file = file_path
            self.input_path.setText(file_path)
            
            # 自動設定輸出檔案路徑
            dir_name = os.path.dirname(file_path)
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            self.output_file = os.path.join(dir_name, f"{base_name}.pdf")
            self.output_path.setText(self.output_file)
            
            # 顯示檔案資訊
            self.show_file_info(file_path)
            
            # 啟用轉換按鈕
            self.convert_btn.setEnabled(True)
    
    def save_file(self):
        """選擇PDF儲存位置"""
        if not self.input_file:
            QMessageBox.warning(self, '警告', '請先選擇Word檔案')
            return
            
        base_name = os.path.splitext(os.path.basename(self.input_file))[0]
        file_path, _ = QFileDialog.getSaveFileName(
            self, '儲存PDF檔案', f"{base_name}.pdf", 'PDF檔案 (*.pdf)'
        )
        
        if file_path:
            self.output_file = file_path
            self.output_path.setText(file_path)
    
    def show_file_info(self, file_path):
        """顯示Word檔案資訊"""
        try:
            doc = Document(file_path)
            
            # 計算頁數（近似值）
            paragraphs = len(doc.paragraphs)
            tables = len(doc.tables)
            sections = len(doc.sections)
            
            # 計算圖片數量
            image_count = 0
            for rel in doc.part.rels.values():
                if "image" in rel.target_ref:
                    image_count += 1
            
            file_size = os.path.getsize(file_path) / 1024  # KB
            
            info = f"檔案名稱: {os.path.basename(file_path)}\n"
            info += f"檔案大小: {file_size:.2f} KB\n"
            info += f"段落數量: {paragraphs}\n"
            info += f"表格數量: {tables}\n"
            info += f"圖片數量: {image_count}\n"
            info += f"章節數量: {sections}\n"
            
            self.file_info.setText(info)
        except Exception as e:
            self.file_info.setText(f"無法讀取檔案資訊: {str(e)}")
    
    def start_conversion(self):
        """開始轉換Word到PDF"""
        if not self.input_file or not self.output_file:
            QMessageBox.warning(self, '警告', '請選擇輸入和輸出檔案')
            return
        
        # 禁用按鈕，避免重複操作
        self.convert_btn.setEnabled(False)
        self.browse_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        
        # 開始轉換執行緒
        self.conversion_thread = WordToPdfThread(self.input_file, self.output_file)
        self.conversion_thread.progress_signal.connect(self.update_progress)
        self.conversion_thread.status_signal.connect(self.update_status)
        self.conversion_thread.finished_signal.connect(self.conversion_finished)
        self.conversion_thread.error_signal.connect(self.conversion_error)
        self.conversion_thread.start()
    
    def update_progress(self, value):
        """更新進度條"""
        self.progress_bar.setValue(value)
    
    def update_status(self, status):
        """更新狀態標籤"""
        self.status_label.setText(status)
    
    def conversion_finished(self, output_file):
        """轉換完成處理"""
        self.status_label.setText('轉換完成!')
        
        # 重新啟用按鈕
        self.convert_btn.setEnabled(True)
        self.browse_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        self.open_btn.setEnabled(True)
        
        QMessageBox.information(
            self, '完成', f'Word檔案已成功轉換為PDF!\n儲存於: {output_file}'
        )
    
    def conversion_error(self, error_msg):
        """轉換錯誤處理"""
        self.status_label.setText('轉換失敗!')
        self.progress_bar.setValue(0)
        
        # 重新啟用按鈕
        self.convert_btn.setEnabled(True)
        self.browse_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        
        QMessageBox.critical(self, '錯誤', f'轉換過程中發生錯誤:\n{error_msg}')
    
    def open_output_file(self):
        """開啟生成的PDF檔案"""
        if os.path.exists(self.output_file):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.output_file))
        else:
            QMessageBox.warning(self, '警告', '找不到輸出檔案')


class PdfMergerTab(QWidget):
    """多文件合併PDF標籤頁"""
    
    def __init__(self, parent=None, main_window=None):
        super().__init__(parent)
        self.main_window = main_window
        self.init_ui()
        
    def init_ui(self):
        """初始化使用者界面"""
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
        
        self.setLayout(main_layout)
        
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
                display_text = f"{title} ({os.path.basename(file_path)})"
                item = QListWidgetItem(display_text)
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
        
        # 開始合併執行緒，並明確傳遞主視窗的參考
        self.conversion_thread = MergePdfThread(self.files, self.output_file, options, self.main_window)
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


class AboutTab(QWidget):
    """關於標籤頁"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
        
    def init_ui(self):
        """初始化使用者界面"""
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
        
        # 版本
        version_label = QLabel('版本: 1.4.0 (穩定版)')
        version_label.setAlignment(Qt.AlignCenter)
        
        # 說明
        desc_text = QTextEdit()
        desc_text.setReadOnly(True)
        desc_text.setHtml("""
        <h2>功能說明</h2>
        <p>本應用程式提供兩個主要功能：</p>
        <ol>
            <li><b>Word轉PDF</b>：將Word文件轉換為PDF格式，保留原始格式、圖片與表格</li>
            <li><b>多文件合併PDF</b>：將多個Word或PDF文件合併為單一PDF檔案，支援調整順序、生成目錄與添加頁碼</li>
        </ol>
        
        <h2>使用須知</h2>
        <ul>
            <li>使用Word轉PDF功能需要安裝Microsoft Word</li>
            <li>合併加密的PDF檔案時，會提示輸入密碼</li>
            <li>合併大型文件可能需要較長時間</li>
            <li>支援繁體中文</li>
            <li>如果轉換失敗，請確認Word可以正常開啟該檔案</li>
            <li>建議以管理員身份執行本程式</li>
        </ul>
        
        <h2>系統需求</h2>
        <ul>
            <li>作業系統：Windows 10或更新版本</li>
            <li>Microsoft Word 2010或更新版本</li>
        </ul>
        
        <h2>常見問題</h2>
        <p><b>問：轉換Word檔案時出現錯誤</b></p>
        <p>答：請確認以下事項：</p>
        <ul>
            <li>Microsoft Word已正確安裝且可以開啟該檔案</li>
            <li>檔案未被其他程式鎖定</li>
            <li>嘗試以管理員身份執行本程式</li>
            <li>關閉所有開啟的Word視窗後再試</li>
        </ul>
        
        <p><b>問：合併PDF時出現錯誤</b></p>
        <p>答：請確認以下事項：</p>
        <ul>
            <li>所有Word檔案都可以正常開啟</li>
            <li>PDF檔案未被損壞或密碼正確</li>
            <li>有足夠的磁碟空間</li>
        </ul>
        """)
        
        # 版權資訊
        copyright_label = QLabel('© 2025 Word與PDF轉換工具')
        copyright_label.setAlignment(Qt.AlignCenter)
        
        # 添加所有元件到主佈局
        main_layout.addWidget(title_label)
        main_layout.addWidget(version_label)
        main_layout.addWidget(desc_text)
        main_layout.addWidget(copyright_label)
        
        self.setLayout(main_layout)


class IntegratedApp(QMainWindow):
    """整合版應用程式主視窗"""

    # 定義信號用於線程安全的密碼輸入
    password_requested = pyqtSignal(str)  # 參數：檔案名稱
    wrong_password = pyqtSignal(str)      # 參數：檔案名稱
    decrypt_error = pyqtSignal(str)       # 參數：錯誤訊息

    def __init__(self):
        super().__init__()
        self._password_result = None
        self._password_ok = False
        self.init_ui()

        # 連接信號到槽函數
        self.password_requested.connect(self.get_password)
        self.wrong_password.connect(self.show_wrong_password_warning)
        self.decrypt_error.connect(self.show_decrypt_error)
        
    def init_ui(self):
        """初始化使用者界面"""
        self.setWindowTitle('Word與PDF轉換工具')
        self.setGeometry(300, 300, 800, 600)
        self.setMinimumSize(700, 500)
        
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
        
        # 創建標籤頁
        self.tabs = QTabWidget()
        
        # 添加Word轉PDF標籤頁
        self.word_to_pdf_tab = WordToPdfTab(parent=self)
        self.tabs.addTab(self.word_to_pdf_tab, 'Word轉PDF')

        # 添加PDF轉Word標籤頁
        self.pdf_to_word_tab = PdfToWordTab(parent=self)
        self.tabs.addTab(self.pdf_to_word_tab, 'PDF轉Word')
        
        # 添加多文件合併PDF標籤頁，並明確傳遞主視窗參考
        self.pdf_merger_tab = PdfMergerTab(parent=self, main_window=self)
        self.tabs.addTab(self.pdf_merger_tab, '多文件合併PDF')
        
        # 添加關於標籤頁
        self.about_tab = AboutTab(parent=self)
        self.tabs.addTab(self.about_tab, '關於')
        
        # 設定主視窗
        self.setCentralWidget(self.tabs)

    @pyqtSlot(str)
    def get_password(self, filename):
        """彈出對話框讓使用者輸入密碼"""
        print(f"get_password called for file: {filename}")
        password, ok = QInputDialog.getText(
            self,
            '需要密碼',
            f'檔案 "{filename}" 已加密，請輸入密碼：',
            QLineEdit.Password
        )
        self._password_result = password if ok else USER_CANCELLED_PASSWORD
        self._password_ok = ok
        print(f"Password dialog result: ok={ok}, password_length={len(password) if ok else 0}")

    def wait_for_password(self):
        """等待並返回密碼輸入結果"""
        # 使用 processEvents 讓事件循環處理密碼對話框
        from PyQt5.QtCore import QEventLoop
        loop = QEventLoop()
        # 設定一個短暫的定時器來檢查結果
        from PyQt5.QtCore import QTimer
        timer = QTimer()
        timer.timeout.connect(loop.quit)
        timer.start(100)  # 每100ms檢查一次

        # 等待密碼輸入完成（最多等待60秒）
        max_wait = 600  # 60秒
        wait_count = 0
        while self._password_result is None and wait_count < max_wait:
            loop.exec_()
            wait_count += 1
            timer.start(100)

        timer.stop()

        # 獲取結果並重置
        result = self._password_result
        ok = self._password_ok
        self._password_result = None
        self._password_ok = False

        print(f"wait_for_password returning: result={result}, ok={ok}")
        return result, ok

    @pyqtSlot(str)
    def show_wrong_password_warning(self, filename):
        """顯示密碼錯誤的警告框"""
        print(f"show_wrong_password_warning called for file: {filename}")
        QMessageBox.warning(self, '密碼錯誤', f"檔案 '{filename}' 的密碼不正確，請重試。")

    @pyqtSlot(str)
    def show_decrypt_error(self, error_message):
        """顯示解密錯誤訊息"""
        print(f"show_decrypt_error called: {error_message}")
        QMessageBox.critical(self, '解密錯誤', error_message)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    integrated_app = IntegratedApp()
    integrated_app.show()
    sys.exit(app.exec_())