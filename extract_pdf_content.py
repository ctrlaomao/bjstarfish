#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import PyPDF2
import sys

def extract_pdf_text(pdf_path):
    """提取PDF文本内容"""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            print(f"\n{'='*70}")
            print(f"文件: {pdf_path}")
            print(f"总页数: {len(pdf_reader.pages)}")
            print(f"{'='*70}\n")
            
            full_text = []
            for i, page in enumerate(pdf_reader.pages, 1):
                text = page.extract_text()
                if text.strip():
                    print(f"--- 第 {i} 页 ---")
                    print(text)
                    print()
                    full_text.append(text)
            
            return '\n'.join(full_text)
    except Exception as e:
        print(f"读取PDF出错: {e}")
        return ""

if __name__ == "__main__":
    pdf1 = "/workspace/20251103150150.pdf"
    pdf2 = "/workspace/20251103150244.pdf"
    
    print("\n" + "="*70)
    print("提取PDF文件内容")
    print("="*70)
    
    text1 = extract_pdf_text(pdf1)
    text2 = extract_pdf_text(pdf2)
    
    # 保存提取的文本
    with open('/workspace/pdf_extracted.txt', 'w', encoding='utf-8') as f:
        f.write("=" * 70 + "\n")
        f.write("PDF 1: 20251103150150.pdf\n")
        f.write("=" * 70 + "\n")
        f.write(text1)
        f.write("\n\n")
        f.write("=" * 70 + "\n")
        f.write("PDF 2: 20251103150244.pdf\n")
        f.write("=" * 70 + "\n")
        f.write(text2)
    
    print("\n✅ PDF内容已提取并保存到 pdf_extracted.txt")
