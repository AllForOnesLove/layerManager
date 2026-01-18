#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel数据汇总工具 - Flask后端服务
"""

import os
import sys
import json
import pandas as pd
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename
import uuid
from datetime import datetime
import traceback
import webbrowser
import threading
import time

app = Flask(__name__)
CORS(app)

# 获取可执行文件所在目录（打包后使用）
if getattr(sys, 'frozen', False):
    # 如果是打包后的exe
    BASE_DIR = os.path.dirname(sys.executable)
    TEMPLATE_DIR = os.path.join(sys._MEIPASS, 'templates')
    app.template_folder = TEMPLATE_DIR
else:
    # 如果是开发环境
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')

# 配置
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'output')
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# 确保目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size


def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def read_excel_file(file_path, sheet_name=None, include_hidden_rows=True, smart_header=True):
    """读取Excel文件（支持.xls和.xlsx，支持隐藏行，支持多行标题）"""
    try:
        # 确定引擎
        if file_path.endswith('.xls'):
            engine = 'xlrd'
        else:
            engine = 'openpyxl'
        
        # 如果启用智能标题处理，先尝试读取原始数据
        if smart_header:
            try:
                # 读取原始数据，不使用header
                if sheet_name:
                    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine, header=None)
                else:
                    if engine == 'xlrd':
                        excel_file = pd.ExcelFile(file_path, engine='xlrd')
                        df_raw = pd.read_excel(excel_file, sheet_name=excel_file.sheet_names[0], engine='xlrd', header=None)
                    else:
                        excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                        df_raw = pd.read_excel(excel_file, sheet_name=excel_file.sheet_names[0], engine='openpyxl', header=None)
                
                if df_raw is not None and not df_raw.empty and len(df_raw) > 2:
                    # 检查是否是复杂的多行标题结构
                    # 第一行可能是标题行（包含"台账"、"合同"等），第二行是列名行，第三行是说明行
                    row0_str = ' '.join([str(v).strip() for v in df_raw.iloc[0].values[:10] if pd.notna(v)])
                    row1_str = ' '.join([str(v).strip() for v in df_raw.iloc[1].values[:10] if pd.notna(v)])
                    row2_str = ' '.join([str(v).strip() for v in df_raw.iloc[2].values[:10] if pd.notna(v)])
                    
                    # 判断结构：
                    # 1. 第一行是标题，第二行包含"序号"、"承租客户名称"等，第三行包含"房产名称"、"起租时间"等
                    # 2. 或者第一行包含"楼层"、"面积"等（收款合同格式）
                    
                    is_complex_structure = False
                    
                    # 台账结构：第一行标题，第二行列名，第三行说明
                    # 汇总sheet结构：第一行标题（如"各项目汇总收入"），第二行列名（如"项目名称"、"1月实际收入"）
                    is_summary_sheet = (('汇总' in row0_str or '收入' in row0_str) and \
                                       ('项目名称' in row1_str or '实际收入' in row1_str or '欠费' in row1_str))
                    
                    is_account_sheet = (('台帐' in row0_str or '台账' in row0_str or '合同' in row0_str) and \
                                       ('序号' in row1_str or '承租客户名称' in row1_str) and \
                                       ('房产名称' in row2_str or '楼层' in row2_str or '起租时间' in row2_str))
                    
                    if is_account_sheet or is_summary_sheet:
                        is_complex_structure = True
                    
                    # 收款合同结构：第一行包含"承租客户名称"等列名，第二行包含"楼层"、"面积"等说明
                    # 或者第一行包含"楼层"、"面积"等（说明行）
                    is_receipt_format = (('承租客户名称' in row0_str or '客户名称' in row0_str) and \
                                        ('楼层' in row1_str or '面积' in row1_str or '房号' in row1_str)) or \
                                       (('楼层' in row0_str or '面积' in row0_str or '房号' in row0_str) and \
                                        not ('序号' in row0_str))
                    
                    if is_receipt_format:
                        is_complex_structure = True
                    
                    if is_complex_structure:
                        # 处理多行标题结构
                        column_names = []
                        row1_values = df_raw.iloc[1].values if len(df_raw) > 1 else []
                        row2_values = df_raw.iloc[2].values if len(df_raw) > 2 else []
                        
                        # 先检查是否是汇总sheet格式（第一行标题，第二行列名）
                        row1_check_for_summary = ' '.join([str(v).strip() for v in row1_values[:15] if pd.notna(v)])
                        is_summary_format = '项目名称' in row1_check_for_summary or '实际收入' in row1_check_for_summary
                        
                        if is_summary_format:
                            # 汇总sheet格式：第一行是标题，第二行是列名
                            new_columns = [str(v).strip() if pd.notna(v) and str(v).strip() and str(v).strip() != 'NaN' else f'Unnamed_{i}' 
                                         for i, v in enumerate(row1_values)]
                            # 确保列名数量匹配
                            max_cols = len(df_raw.columns)
                            while len(new_columns) < max_cols:
                                new_columns.append(f'Unnamed_{len(new_columns)}')
                            
                            # 从第三行（索引2）开始是数据
                            df = df_raw.iloc[2:].copy()
                            df.columns = new_columns[:len(df.columns)]
                            df = df.reset_index(drop=True)
                            
                            # 将NaN转换为None以便JSON序列化
                            df = df.where(pd.notnull(df), None)
                            return df, None
                        # 如果第三行存在且包含说明性列名，优先使用第三行（台账格式）
                        elif len(row2_values) > 0 and any('房产名称' in str(v) or '起租时间' in str(v) or '楼层' in str(v) 
                                                         for v in row2_values if pd.notna(v)):
                            # 台账结构：第2行是部分列名（如"序号"、"承租客户名称"），第3行是更多列名（如"房产名称"、"楼层及房号"）
                            # 需要合并两行：优先使用第3行（说明行），如果为空则使用第2行（列名行）
                            # 遍历所有列，确保每个位置都有列名
                            max_cols = len(df_raw.columns)
                            column_names = []
                            
                            for i in range(max_cols):
                                col_name = None
                                # 优先使用第3行（说明行）的值
                                if i < len(row2_values):
                                    val2 = row2_values[i]
                                    if pd.notna(val2) and str(val2).strip() and str(val2).strip() != 'NaN':
                                        col_name = str(val2).strip()
                                
                                # 如果第3行为空，使用第2行（列名行）的值
                                if not col_name and i < len(row1_values):
                                    val1 = row1_values[i]
                                    if pd.notna(val1) and str(val1).strip() and str(val1).strip() != 'NaN':
                                        col_name = str(val1).strip()
                                
                                # 如果两行都为空，检查数据行是否有值
                                if not col_name:
                                    # 检查从第4行（索引3）开始的数据行，看该列是否有值
                                    has_data = False
                                    if len(df_raw) > 3:
                                        for row_idx in range(3, min(len(df_raw), 10)):  # 检查前几行数据
                                            if i < len(df_raw.iloc[row_idx]) and pd.notna(df_raw.iloc[row_idx, i]):
                                                val = df_raw.iloc[row_idx, i]
                                                # 如果值不是NaN且不是空字符串
                                                if pd.notna(val) and str(val).strip():
                                                    has_data = True
                                                    break
                                    
                                    # 如果该列有数据但没有列名，使用更友好的默认名称
                                    if has_data:
                                        # 尝试从前面的列推断名称
                                        prev_col_name = None
                                        if i > 0 and len(column_names) > 0:
                                            prev_col_name = column_names[-1]
                                        
                                        # 如果前一列是"序号"，当前列可能是序号的一部分或辅助列
                                        if prev_col_name and ('序号' in str(prev_col_name) or prev_col_name.startswith('列')):
                                            column_names.append(f'列{i+1}')
                                        else:
                                            column_names.append(f'列{i+1}')
                                    else:
                                        column_names.append(f'Unnamed_{i}')
                                else:
                                    column_names.append(col_name)
                            
                            # 从第四行（索引3）开始是数据
                            df = df_raw.iloc[3:].copy()
                            df.columns = column_names[:len(df.columns)]
                            df = df.reset_index(drop=True)
                            
                            # 将NaN转换为None以便JSON序列化
                            df = df.where(pd.notnull(df), None)
                            return df, None
                        # 收款合同结构处理
                        # 收款合同的实际情况：第一行（索引0）是列名行（如"承租客户名称"），但某些列为空
                        # 第二行（索引1）是说明行（如"楼层及房号"、"面积"），对应第一行为空的位置
                        # 需要合并：优先使用第一行，如果为空则使用第二行
                        
                        # 检查是否是收款合同格式：第一行包含"承租客户名称"等，第二行包含"楼层"、"面积"等
                        row0_check = ' '.join([str(v).strip() for v in df_raw.iloc[0].values[:15] if pd.notna(v)])
                        row1_check_for_receipt = ' '.join([str(v).strip() for v in df_raw.iloc[1].values[:15] if pd.notna(v)]) if len(df_raw) > 1 else ''
                        
                        if len(df_raw) > 1 and ('承租客户名称' in row0_check or '客户名称' in row0_check) and \
                           ('楼层' in row1_check_for_receipt or '面积' in row1_check_for_receipt or '楼层及房号' in row1_check_for_receipt):
                            row0_check = ' '.join([str(v).strip() for v in df_raw.iloc[0].values[:15] if pd.notna(v)])
                            row1_check = ' '.join([str(v).strip() for v in df_raw.iloc[1].values[:15] if pd.notna(v)])
                            
                            # 如果第一行包含"承租客户名称"等列名，第二行包含"楼层"、"面积"等说明
                            # 这是收款合同文件的格式
                            if ('承租客户名称' in row0_check or '客户名称' in row0_check) and \
                               ('楼层' in row1_check or '面积' in row1_check or '房号' in row1_check or '楼层及房号' in row1_check):
                                # 使用第一行作为列名（但需要处理空值）
                                # 如果第一行某列为空，尝试使用第二行（说明行）的值
                                new_columns = []
                                row0_values = df_raw.iloc[0].values
                                row1_values = df_raw.iloc[1].values if len(df_raw) > 1 else []
                                
                                for i in range(len(df_raw.columns)):
                                    col_name = None
                                    # 优先使用第一行的值（列名行，如"承租客户名称"）
                                    if i < len(row0_values) and pd.notna(row0_values[i]) and str(row0_values[i]).strip() and str(row0_values[i]).strip() != 'NaN':
                                        col_name = str(row0_values[i]).strip()
                                    # 如果第一行为空，使用第二行（说明行）的值作为列名（如"楼层及房号"、"面积"）
                                    if not col_name and i < len(row1_values) and pd.notna(row1_values[i]) and str(row1_values[i]).strip() and str(row1_values[i]).strip() != 'NaN':
                                        col_name = str(row1_values[i]).strip()
                                    
                                    if col_name:
                                        new_columns.append(col_name)
                                    else:
                                        new_columns.append(f'Unnamed_{i}')
                                
                                # 从第二行开始是数据（跳过说明行）
                                df = df_raw.iloc[1:].copy()
                                df.columns = new_columns[:len(df.columns)]
                                df = df.reset_index(drop=True)
                                
                                # 跳过说明行（现在的第一行）
                                if len(df) > 0:
                                    first_row_str = ' '.join([str(v).strip() for v in df.iloc[0].values[:15] if pd.notna(v)])
                                    if '楼层' in first_row_str or '面积' in first_row_str or '房号' in first_row_str:
                                        df = df[1:].copy()
                                        df = df.reset_index(drop=True)
                                
                                # 将NaN转换为None以便JSON序列化
                                df = df.where(pd.notnull(df), None)
                                return df, None
            except Exception as e:
                # 如果智能处理失败，回退到标准方法
                pass
        
        # 标准读取方法
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine)
        else:
            # 读取第一个sheet
            if engine == 'xlrd':
                excel_file = pd.ExcelFile(file_path, engine='xlrd')
            else:
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            df = pd.read_excel(excel_file, sheet_name=excel_file.sheet_names[0])
        
        # 将NaN转换为None以便JSON序列化
        df = df.where(pd.notnull(df), None)
        return df, None
    except Exception as e:
        return None, str(e)


def get_sheet_names(file_path):
    """获取Excel文件的所有sheet名称（支持.xls和.xlsx）"""
    try:
        # 确定引擎
        if file_path.endswith('.xls'):
            try:
                excel_file = pd.ExcelFile(file_path, engine='xlrd')
            except:
                # 如果xlrd失败，尝试openpyxl
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        else:
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        return excel_file.sheet_names, None
    except Exception as e:
        return None, str(e)


def find_column_by_keywords(df, keywords):
    """通过关键词查找列名（支持模糊匹配）"""
    if df is None or df.empty:
        return None
    
    df_columns = [str(col).strip() for col in df.columns]
    
    # 先尝试精确匹配
    for keyword in keywords:
        for col in df_columns:
            if keyword == col:
                return col
    
    # 再尝试包含匹配
    for keyword in keywords:
        for col in df_columns:
            if keyword in col or col in keyword:
                return col
    
    return None


@app.route('/')
def index():
    """主页面"""
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """上传Excel文件"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': '没有文件被上传'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '文件名为空'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': '不支持的文件格式，请上传 .xlsx 或 .xls 文件'}), 400
        
        # 生成唯一文件名
        file_id = str(uuid.uuid4())
        filename = secure_filename(file.filename)
        file_ext = filename.rsplit('.', 1)[1].lower()
        saved_filename = f"{file_id}.{file_ext}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], saved_filename)
        
        file.save(file_path)
        
        # 读取文件信息（使用智能标题处理）
        df, error = read_excel_file(file_path, sheet_name=None, smart_header=True)
        if error:
            os.remove(file_path)
            return jsonify({'error': f'读取文件失败: {error}'}), 400
        
        # 获取sheet名称
        sheet_names, sheet_error = get_sheet_names(file_path)
        if sheet_error:
            sheet_names = ['Sheet1']
        
        # 准备预览数据，确保NaN转换为None
        preview_df = df.head(10)
        # 将NaN替换为None以便JSON序列化
        preview_df = preview_df.where(pd.notnull(preview_df), None)
        preview_data = preview_df.to_dict('records')
        # 再次清理，确保没有NaN值
        for record in preview_data:
            for key, value in record.items():
                if pd.isna(value):
                    record[key] = None
        
        return jsonify({
            'file_id': file_id,
            'filename': filename,
            'saved_filename': saved_filename,
            'columns': list(df.columns),
            'row_count': len(df),
            'sheet_names': sheet_names,
            'preview': preview_data  # 前10行预览
        })
    
    except Exception as e:
        return jsonify({'error': f'上传失败: {str(e)}'}), 500


@app.route('/api/read_file', methods=['POST'])
def read_file():
    """读取已上传的文件（用于切换sheet）"""
    try:
        data = request.json
        file_id = data.get('file_id')
        sheet_name = data.get('sheet_name')
        
        if not file_id:
            return jsonify({'error': '缺少file_id参数'}), 400
        
        # 查找文件
        file_path = None
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            if filename.startswith(file_id):
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                break
        
        if not file_path or not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404
        
        # 读取文件（包括隐藏行，使用智能标题处理）
        df, error = read_excel_file(file_path, sheet_name, include_hidden_rows=True, smart_header=True)
        if error:
            return jsonify({'error': f'读取文件失败: {error}'}), 400
        
        # 准备预览数据，确保NaN转换为None
        preview_df = df.head(10)
        # 将NaN替换为None以便JSON序列化
        preview_df = preview_df.where(pd.notnull(preview_df), None)
        preview_data = preview_df.to_dict('records')
        # 再次清理，确保没有NaN值
        for record in preview_data:
            for key, value in record.items():
                if pd.isna(value):
                    record[key] = None
        
        return jsonify({
            'columns': list(df.columns),
            'row_count': len(df),
            'preview': preview_data,
            'column_suggestions': {}  # 可以添加字段自动建议
        })
    
    except Exception as e:
        return jsonify({'error': f'读取失败: {str(e)}'}), 500


@app.route('/api/merge', methods=['POST'])
def merge_files():
    """合并Excel文件"""
    try:
        data = request.json
        tables_config = data.get('tables', [])
        output_filename = data.get('output_filename', '汇总表.xlsx')
        include_source = data.get('include_source', True)
        
        if not tables_config:
            return jsonify({'error': '没有配置要合并的表'}), 400
        
        # 存储所有提取的数据
        all_data = []
        
        # 处理每个表
        for table_config in tables_config:
            file_id = table_config.get('file_id')
            sheet_name = table_config.get('sheet_name')
            fields = table_config.get('fields', [])
            table_name = table_config.get('table_name', '未知表')
            
            if not file_id or not fields:
                continue
            
            # 查找文件
            file_path = None
            for filename in os.listdir(app.config['UPLOAD_FOLDER']):
                if filename.startswith(file_id):
                    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    break
            
            if not file_path or not os.path.exists(file_path):
                continue
            
            # 读取文件（包括隐藏行，使用智能标题处理）
            df, error = read_excel_file(file_path, sheet_name, include_hidden_rows=True, smart_header=True)
            if error:
                continue
            
            # 提取指定字段
            extracted_data = {}
            if include_source:
                extracted_data['来源表'] = table_name
            
            for field_config in fields:
                source_field = field_config.get('source')
                target_field = field_config.get('target', source_field)
                keywords = field_config.get('keywords', [])  # 支持关键词列表
                
                # 优先使用精确匹配
                if source_field and source_field in df.columns:
                    extracted_data[target_field] = df[source_field]
                elif keywords:
                    # 如果指定了关键词，尝试模糊匹配
                    matched_col = find_column_by_keywords(df, keywords)
                    if matched_col:
                        extracted_data[target_field] = df[matched_col]
                    else:
                        # 字段不存在，填充空值
                        extracted_data[target_field] = [None] * len(df)
                else:
                    # 字段不存在，填充空值
                    extracted_data[target_field] = [None] * len(df)
            
            # 转换为DataFrame
            extracted_df = pd.DataFrame(extracted_data)
            all_data.append(extracted_df)
        
        if not all_data:
            return jsonify({'error': '没有可合并的数据'}), 400
        
        # 合并所有数据
        merged_df = pd.concat(all_data, ignore_index=True)
        
        # 填充NaN值
        merged_df = merged_df.fillna('-')
        
        # 生成输出文件
        output_id = str(uuid.uuid4())
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{output_id}.xlsx")
        
        # 使用openpyxl引擎写入，支持样式
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            merged_df.to_excel(writer, sheet_name='数据汇总', index=False)
            
            # 获取worksheet对象以设置样式
            worksheet = writer.sheets['数据汇总']
            
            # 设置列宽（自动调整）
            from openpyxl.utils import get_column_letter
            for idx, col in enumerate(merged_df.columns, 1):
                max_length = max(
                    merged_df[col].astype(str).map(len).max(),
                    len(str(col))
                )
                column_letter = get_column_letter(idx)
                worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)
            
            # 设置表头样式
            from openpyxl.styles import Font, PatternFill, Alignment
            
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
        
        return jsonify({
            'output_id': output_id,
            'filename': output_filename,
            'row_count': len(merged_df),
            'column_count': len(merged_df.columns),
            'columns': list(merged_df.columns)
        })
    
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'合并失败: {str(e)}'}), 500


@app.route('/api/download/<output_id>')
def download_file(output_id):
    """下载生成的汇总文件"""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{output_id}.xlsx")
        
        if not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name='汇总表.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    except Exception as e:
        return jsonify({'error': f'下载失败: {str(e)}'}), 500


@app.route('/api/cleanup', methods=['POST'])
def cleanup():
    """清理临时文件"""
    try:
        data = request.json
        file_ids = data.get('file_ids', [])
        output_id = data.get('output_id')
        
        # 清理上传的文件
        for file_id in file_ids:
            for filename in os.listdir(app.config['UPLOAD_FOLDER']):
                if filename.startswith(file_id):
                    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    if os.path.exists(file_path):
                        os.remove(file_path)
        
        # 清理输出文件
        if output_id:
            file_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{output_id}.xlsx")
            if os.path.exists(file_path):
                os.remove(file_path)
        
        return jsonify({'success': True})
    
    except Exception as e:
        return jsonify({'error': f'清理失败: {str(e)}'}), 500


def open_browser():
    """延迟打开浏览器"""
    time.sleep(1.5)  # 等待服务器启动
    webbrowser.open('http://localhost:5000')


if __name__ == '__main__':
    print("=" * 60)
    print(" " * 15 + "Excel数据汇总工具")
    print("=" * 60)
    print()
    print("正在启动服务...")
    print(f"访问地址: http://localhost:5000")
    print()
    print("提示: 浏览器将自动打开，如果没有自动打开，请手动访问上述地址")
    print("按 Ctrl+C 停止服务")
    print("=" * 60)
    print()
    
    # 在非调试模式下自动打开浏览器
    if not getattr(sys, 'frozen', False) or True:  # 打包后也打开浏览器
        browser_thread = threading.Thread(target=open_browser)
        browser_thread.daemon = True
        browser_thread.start()
    
    # 打包后使用debug=False，开发时使用debug=True
    app.run(debug=not getattr(sys, 'frozen', False), host='127.0.0.1', port=5000, use_reloader=False)
