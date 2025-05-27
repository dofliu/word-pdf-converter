import PyInstaller.__main__
import os
import shutil

# 確保輸出目錄存在
if os.path.exists('dist'):
    shutil.rmtree('dist')
os.makedirs('dist', exist_ok=True)

# 打包整合版應用
PyInstaller.__main__.run([
    'src/integrated_app.py',
    '--name=Word_PDF_Converter',
    '--windowed',  # 創建無控制台窗口的GUI應用
    '--hidden-import=win32com',
    '--hidden-import=win32com.client',
    '--hidden-import=pythoncom',  # 新增pythoncom支援
    '--hidden-import=docx2pdf',
    '--hidden-import=reportlab.lib.units',
    '--exclude-module=PyQt6',  # 排除PyQt6
    '--clean',
    '--onefile',  # 創建單一執行檔
])

# 複製使用說明到輸出目錄
shutil.copy('README.md', 'dist/使用說明.md')

print("打包完成！執行檔位於 dist/Word_PDF_Converter.exe")
