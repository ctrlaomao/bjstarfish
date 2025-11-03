#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pdfplumber
import json

def extract_pdf_with_pdfplumber(pdf_path):
    """使用pdfplumber提取PDF内容"""
    try:
        content = {
            'filename': pdf_path.split('/')[-1],
            'pages': [],
            'total_pages': 0
        }
        
        with pdfplumber.open(pdf_path) as pdf:
            content['total_pages'] = len(pdf.pages)
            
            for i, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                
                # 提取表格
                tables = page.extract_tables()
                
                page_content = {
                    'page_number': i,
                    'text': text if text else "",
                    'tables': tables if tables else []
                }
                
                content['pages'].append(page_content)
        
        return content
    except Exception as e:
        print(f"提取PDF出错 {pdf_path}: {e}")
        return None

def main():
    print("\n" + "="*70)
    print("使用pdfplumber提取PDF内容")
    print("="*70 + "\n")
    
    pdf1_path = "/workspace/20251103150150.pdf"
    pdf2_path = "/workspace/20251103150244.pdf"
    
    print("正在提取第一个PDF...")
    pdf1_content = extract_pdf_with_pdfplumber(pdf1_path)
    if pdf1_content:
        print(f"  ✓ {pdf1_content['filename']}: {pdf1_content['total_pages']} 页")
        # 显示前几页内容
        for page in pdf1_content['pages'][:3]:
            if page['text']:
                print(f"\n--- 第 {page['page_number']} 页（前200字符）---")
                print(page['text'][:200])
    
    print("\n正在提取第二个PDF...")
    pdf2_content = extract_pdf_with_pdfplumber(pdf2_path)
    if pdf2_content:
        print(f"  ✓ {pdf2_content['filename']}: {pdf2_content['total_pages']} 页")
        # 显示前几页内容
        for page in pdf2_content['pages'][:3]:
            if page['text']:
                print(f"\n--- 第 {page['page_number']} 页（前200字符）---")
                print(page['text'][:200])
    
    # 保存完整内容到JSON
    all_content = {
        'pdf1': pdf1_content,
        'pdf2': pdf2_content
    }
    
    with open('/workspace/pdf_content.json', 'w', encoding='utf-8') as f:
        json.dump(all_content, f, ensure_ascii=False, indent=2)
    
    print("\n" + "="*70)
    print("✅ PDF内容已提取并保存到 pdf_content.json")
    print("="*70)
    
    # 分析内容识别文件类型
    print("\n分析文件类型...")
    
    if pdf1_content and pdf1_content['pages']:
        first_text = pdf1_content['pages'][0]['text'][:500] if pdf1_content['pages'][0]['text'] else ""
        print(f"\nPDF 1 ({pdf1_content['filename']}) 内容特征:")
        if '需求' in first_text or '确认' in first_text:
            print("  可能是：需求确认单")
        if '上线' in first_text or '验收' in first_text:
            print("  可能是：上线确认书/验收文档")
        print(f"  关键词: {', '.join([w for w in ['需求', '确认', '上线', '验收', '系统'] if w in first_text])}")
    
    if pdf2_content and pdf2_content['pages']:
        first_text = pdf2_content['pages'][0]['text'][:500] if pdf2_content['pages'][0]['text'] else ""
        print(f"\nPDF 2 ({pdf2_content['filename']}) 内容特征:")
        if '需求' in first_text or '确认' in first_text:
            print("  可能是：需求确认单")
        if '上线' in first_text or '验收' in first_text:
            print("  可能是：上线确认书/验收文档")
        print(f"  关键词: {', '.join([w for w in ['需求', '确认', '上线', '验收', '系统'] if w in first_text])}")

if __name__ == "__main__":
    main()
