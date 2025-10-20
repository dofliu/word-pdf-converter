import os
import sys
import time
import threading
import tempfile
import shutil
import platform
import subprocess
import docx2pdf
from docx import Document
import pypdf
from pypdf.errors import FileNotDecryptedError, WrongPasswordError
from pdf2docx import Converter
import win32com.client
import pythoncom
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm


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


# 增強版 Word 轉 PDF 函數
def convert_word_to_pdf(word_path, pdf_path):
    """使用多種方法嘗試將Word轉換為PDF，提高成功率"""
    system = platform.system()
    
    # 確保路徑是絕對路徑
    word_path = os.path.abspath(word_path)
    pdf_path = os.path.abspath(pdf_path)
    
    print(f"開始轉換: {word_path} -> {pdf_path}")
    
    # 方法1: 使用docx2pdf
    try:
        print("嘗試使用docx2pdf轉換...")
        docx2pdf.convert(word_path, pdf_path)
        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
            print("docx2pdf轉換成功!")
            return True
    except Exception as e:
        print(f"docx2pdf轉換失敗: {str(e)}")
    
    # 方法2: 使用COM自動化 (Windows專用)
    if system == 'Windows':
        try:
            print("嘗試使用COM自動化轉換...")
            
            # 初始化COM
            pythoncom.CoInitialize()
            
            # 創建Word應用實例
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            try:
                # 打開文檔
                doc = word.Documents.Open(word_path)
                # 另存為PDF (17代表PDF格式)
                doc.SaveAs(pdf_path, FileFormat=17)
                doc.Close()
                print("COM自動化轉換成功!")
                return True
            except Exception as e:
                print(f"Word COM轉換過程失敗: {str(e)}")
                return False
            finally:
                # 確保Word應用關閉
                try:
                    word.Quit()
                except:
                    pass
                # 釋放COM資源
                pythoncom.CoUninitialize()
        except Exception as e:
            print(f"Word COM初始化失敗: {str(e)}")
    
    # 方法3: 使用LibreOffice (如果安裝了)
    try:
        print("嘗試使用LibreOffice轉換...")
        # 檢查LibreOffice是否存在
        if system == 'Windows':
            soffice_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            ]
            soffice_path = None
            for path in soffice_paths:
                if os.path.exists(path):
                    soffice_path = path
                    break
                    
            if soffice_path:
                cmd = [soffice_path, '--headless', '--convert-to', 'pdf', '--outdir', 
                       os.path.dirname(pdf_path), word_path]
                process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                stdout, stderr = process.communicate()
                
                # 檢查轉換後的文件名稱
                base_name = os.path.splitext(os.path.basename(word_path))[0]
                temp_pdf = os.path.join(os.path.dirname(pdf_path), f"{base_name}.pdf")
                
                if os.path.exists(temp_pdf):
                    # 如果輸出路徑不同，移動文件
                    if temp_pdf != pdf_path:
                        shutil.move(temp_pdf, pdf_path)
                    print("LibreOffice轉換成功!")
                    return True
        else:
            # Linux/Mac檢查
            try:
                process = subprocess.Popen(['which', 'soffice'], stdout=subprocess.PIPE)
                soffice_path = process.communicate()[0].strip().decode('utf-8')
                
                if soffice_path:
                    cmd = [soffice_path, '--headless', '--convert-to', 'pdf', '--outdir', 
                           os.path.dirname(pdf_path), word_path]
                    process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    stdout, stderr = process.communicate()
                    
                    # 檢查轉換後的文件名稱
                    base_name = os.path.splitext(os.path.basename(word_path))[0]
                    temp_pdf = os.path.join(os.path.dirname(pdf_path), f"{base_name}.pdf")
                    
                    if os.path.exists(temp_pdf):
                        # 如果輸出路徑不同，移動文件
                        if temp_pdf != pdf_path:
                            shutil.move(temp_pdf, pdf_path)
                        print("LibreOffice轉換成功!")
                        return True
            except:
                pass
    except Exception as e:
        print(f"LibreOffice轉換失敗: {str(e)}")
    
    print("所有轉換方法都失敗了")
    return False


def convert_pdf_to_word(pdf_path, word_path):
    """將PDF轉換為Word文件"""
    try:
        print(f"開始轉換: {pdf_path} -> {word_path}")
        
        # 使用pdf2docx進行轉換
        cv = Converter(pdf_path)
        cv.convert(word_path, start=0, end=None)
        cv.close()
        
        # 檢查轉換結果
        if os.path.exists(word_path) and os.path.getsize(word_path) > 0:
            print("PDF轉Word成功!")
            return True
        else:
            print("PDF轉Word失敗: 輸出檔案為空")
            return False
    except Exception as e:
        print(f"PDF轉Word失敗: {str(e)}")
        return False

def to_roman(num):
    """將數字轉換為羅馬數字"""
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syms = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
    roman_num = ''
    i = 0
    while num > 0:
        for _ in range(num // val[i]):
            roman_num += syms[i]
            num -= val[i]
        i += 1
    return roman_num

def create_toc(prepared_pdfs, temp_dir, CHINESE_FONT):
    """使用已解密的PdfReader物件創建目錄頁"""
    toc_pdf_path = os.path.join(temp_dir, "toc.pdf")
    c = canvas.Canvas(toc_pdf_path, pagesize=A4)
    c.setFont(CHINESE_FONT, 24)
    c.drawString(4*cm, 27*cm, "目錄")
    
    y_position = 25*cm
    c.setFont(CHINESE_FONT, 12)
    
    current_page_count = 1  # 目錄頁本身
    
    for i, pdf_info in enumerate(prepared_pdfs):
        title = pdf_info['title']
        if len(title) > 50:
            title = title[:47] + "..."
            
        c.drawString(2*cm, y_position, f"{i+1}. {title}")
        c.drawString(16*cm, y_position, f"{current_page_count + 1}")
        
        y_position -= 0.8*cm
        if y_position < 2*cm:
            c.showPage()
            c.setFont(CHINESE_FONT, 12)
            y_position = 27*cm
        
        current_page_count += len(pdf_info['reader'].pages)
    
    c.save()
    return toc_pdf_path

def add_page_numbers(pdf_path, temp_dir, options, CHINESE_FONT):
    """添加頁碼到PDF"""
    output_path = os.path.join(temp_dir, "with_page_numbers.pdf")
    reader = pypdf.PdfReader(pdf_path)
    writer = pypdf.PdfWriter()
    
    page_format = options['page_number_format']
    start_number = options['start_page_number']
    
    for i, page in enumerate(reader.pages):
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        
        packet = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
        c = canvas.Canvas(packet.name, pagesize=(page_width, page_height))
        c.setFont(CHINESE_FONT, 10)
        
        page_number = start_number + i
        page_text = str(page_number)
        if page_format == "羅馬數字":
            page_text = to_roman(page_number)
        
        c.drawString(page_width/2 - 10, 20, page_text)
        c.save()
        
        watermark = pypdf.PdfReader(packet.name)
        page.merge_page(watermark.pages[0])
        writer.add_page(page)
        
        packet.close()
        os.unlink(packet.name)
    
    with open(output_path, 'wb') as output_file:
        writer.write(output_file)
    
    return output_path

def prepare_pdf_reader_core(pdf_path, password=None):
    """
    健壯地打開一個PDF文件。如果文件已加密，它將嘗試使用提供的密碼解密。
    如果加密且未提供密碼，或密碼錯誤，則拋出異常。
    """
    reader = pypdf.PdfReader(pdf_path)
    if not reader.is_encrypted:
        return reader

    if password:
        try:
            if reader.decrypt(password):
                return reader
            else:
                raise WrongPasswordError(f"檔案 '{os.path.basename(pdf_path)}' 的密碼不正確。")
        except Exception as e:
            raise e
    else:
        raise FileNotDecryptedError(f"檔案 '{os.path.basename(pdf_path)}' 已加密，需要密碼。")

def merge_pdfs_core(file_list, output_file, options, temp_dir, password_callback=None):
    """核心PDF合併邏輯，可由GUI或API呼叫。"""
    prepared_pdfs = [] # 儲存 {'reader': PdfReader物件, 'title': str}

    for i, file_info in enumerate(file_list):
        file_path = file_info['path']
        file_name = os.path.basename(file_path)

        pdf_to_open = None
        if file_path.lower().endswith(('.docx', '.doc')):
            temp_pdf = os.path.join(temp_dir, f"temp_{i}.pdf")
            if not convert_word_to_pdf(file_path, temp_pdf):
                raise Exception(f"無法轉換 Word 檔案: {file_name}。")
            pdf_to_open = temp_pdf
        else:
            pdf_to_open = file_path
        
        reader = None
        try:
            reader = prepare_pdf_reader_core(pdf_to_open)
        except FileNotDecryptedError:
            if password_callback:
                password, ok = password_callback(file_name)
                if not ok:
                    raise Exception(f"已取消操作，因為未提供 '{file_name}' 的密碼。")
                reader = prepare_pdf_reader_core(pdf_to_open, password)
            else:
                raise Exception(f"檔案 '{file_name}' 已加密，需要密碼。")
        except Exception as e:
            raise Exception(f"處理檔案 '{file_name}' 時發生錯誤：{e}")

        prepared_pdfs.append({'reader': reader, 'title': file_info['title']})

    # 建立目錄 (如果需要)
    toc_pdf_path = None
    if options['generate_toc']:
        toc_pdf_path = create_toc(prepared_pdfs, temp_dir, CHINESE_FONT)

    # 合併所有PDF
    merger = pypdf.PdfWriter()
    if toc_pdf_path:
        merger.append(toc_pdf_path)
    for pdf_info in prepared_pdfs:
        merger.append(fileobj=pdf_info['reader'])

    merged_pdf_path = os.path.join(temp_dir, "merged.pdf")
    merger.write(merged_pdf_path)
    merger.close()

    # 添加頁碼 (如果需要)
    final_pdf_path = merged_pdf_path
    if options['add_page_numbers']:
        final_pdf_path = add_page_numbers(merged_pdf_path, temp_dir, options, CHINESE_FONT)
    
    shutil.copy2(final_pdf_path, output_file)
