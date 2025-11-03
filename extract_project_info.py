#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
from docx import Document
import glob

def extract_docx_text(docx_path):
    """提取docx文件的文本内容"""
    try:
        doc = Document(docx_path)
        text = []
        for para in doc.paragraphs:
            if para.text.strip():
                text.append(para.text.strip())
        return "\n".join(text)
    except Exception as e:
        return f"读取失败: {str(e)}"

def analyze_project_docs():
    """分析项目文档"""
    print("=" * 60)
    print("项目文档内容提取")
    print("=" * 60)
    
    # 1. 读取需求规格说明书
    print("\n【需求规格说明书】")
    req_doc = "/workspace/需求规格说明书-20230717.docx"
    if os.path.exists(req_doc):
        content = extract_docx_text(req_doc)
        print(content[:1000])  # 打印前1000字符
        print("...\n")
    
    # 2. 读取会议纪要
    print("\n【会议纪要】")
    meeting_dir = "/workspace/项目会议纪要"
    meeting_files = sorted(glob.glob(os.path.join(meeting_dir, "*.docx")))
    for i, meeting_file in enumerate(meeting_files[:3]):  # 只读前3个
        filename = os.path.basename(meeting_file)
        print(f"\n--- {filename} ---")
        content = extract_docx_text(meeting_file)
        print(content[:500])  # 打印前500字符
        print("...")
        if i >= 2:
            break
    
    print(f"\n总共{len(meeting_files)}个会议纪要文件")
    
    # 3. 统计周报信息
    print("\n【项目周报】")
    weekly_dir = "/workspace/项目周报"
    weekly_files = sorted(glob.glob(os.path.join(weekly_dir, "*.xlsx")))
    print(f"周报文件数量: {len(weekly_files)}")
    print(f"时间跨度: {os.path.basename(weekly_files[0])} 至 {os.path.basename(weekly_files[-1])}")
    
    return {
        "meeting_count": len(meeting_files),
        "weekly_count": len(weekly_files),
        "start_date": "2023-04-25",
        "end_date": "2024-03-15"
    }

if __name__ == "__main__":
    analyze_project_docs()
