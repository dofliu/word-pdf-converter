#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試加密PDF密碼輸入功能的腳本
"""

import os
import sys
from pypdf import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

def create_test_pdf(filename, encrypted=False, password="test123"):
    """建立測試用的PDF檔案"""
    print(f"建立測試PDF: {filename}")

    # 建立一個簡單的PDF
    c = canvas.Canvas(filename, pagesize=A4)
    c.drawString(100, 750, "測試PDF文件")
    c.drawString(100, 700, "這是用來測試密碼輸入功能的測試檔案")
    c.drawString(100, 650, f"加密狀態: {'是' if encrypted else '否'}")
    if encrypted:
        c.drawString(100, 600, f"密碼: {password}")
    c.save()

    # 如果需要加密，重新讀取並加密
    if encrypted:
        reader = PdfReader(filename)
        writer = PdfWriter()

        # 複製所有頁面
        for page in reader.pages:
            writer.add_page(page)

        # 加密PDF
        writer.encrypt(password)

        # 寫入加密的PDF
        with open(filename, 'wb') as f:
            writer.write(f)

        print(f"[OK] 已建立加密PDF，密碼為: {password}")
    else:
        print(f"[OK] 已建立未加密PDF")

def main():
    """建立測試檔案"""
    test_dir = "test_files"

    # 建立測試目錄
    if not os.path.exists(test_dir):
        os.makedirs(test_dir)
        print(f"建立測試目錄: {test_dir}")

    # 建立各種測試PDF
    test_files = [
        ("test_normal.pdf", False, None),
        ("test_encrypted_simple.pdf", True, "123456"),
        ("test_encrypted_complex.pdf", True, "MyP@ssw0rd!"),
        ("test_encrypted_chinese.pdf", True, "測試密碼"),
    ]

    for filename, encrypted, password in test_files:
        filepath = os.path.join(test_dir, filename)
        create_test_pdf(filepath, encrypted, password)

    print("\n" + "="*60)
    print("測試檔案建立完成！")
    print("="*60)
    print("\n測試檔案清單：")
    print("1. test_normal.pdf - 未加密")
    print("2. test_encrypted_simple.pdf - 密碼: 123456")
    print("3. test_encrypted_complex.pdf - 密碼: MyP@ssw0rd!")
    print("4. test_encrypted_chinese.pdf - 密碼: 測試密碼")
    print("\n請使用這些檔案測試PDF合併功能中的密碼輸入機制。")
    print("="*60)

if __name__ == '__main__':
    main()