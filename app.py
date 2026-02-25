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


def read_excel_file(file_path, sheet_name=None):
    """读取Excel文件"""
    try:
        # 尝试读取指定sheet，如果没有指定则读取第一个
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        else:
            # 读取第一个sheet
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            df = pd.read_excel(excel_file, sheet_name=excel_file.sheet_names[0])
        
        # 将NaN转换为None以便JSON序列化
        df = df.where(pd.notnull(df), None)
        return df, None
    except Exception as e:
        return None, str(e)


def get_sheet_names(file_path):
    """获取Excel文件的所有sheet名称"""
    try:
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        return excel_file.sheet_names, None
    except Exception as e:
        return None, str(e)


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
        
        # 读取文件信息
        df, error = read_excel_file(file_path)
        if error:
            os.remove(file_path)
            return jsonify({'error': f'读取文件失败: {error}'}), 400
        
        # 获取sheet名称
        sheet_names, sheet_error = get_sheet_names(file_path)
        if sheet_error:
            sheet_names = ['Sheet1']
        
        return jsonify({
            'file_id': file_id,
            'filename': filename,
            'saved_filename': saved_filename,
            'columns': list(df.columns),
            'row_count': len(df),
            'sheet_names': sheet_names,
            'preview': df.head(10).to_dict('records')  # 前10行预览
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
        
        # 读取文件
        df, error = read_excel_file(file_path, sheet_name)
        if error:
            return jsonify({'error': f'读取文件失败: {error}'}), 400
        
        return jsonify({
            'columns': list(df.columns),
            'row_count': len(df),
            'preview': df.head(10).to_dict('records')
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
            
            # 读取文件
            df, error = read_excel_file(file_path, sheet_name)
            if error:
                continue
            
            # 提取指定字段
            extracted_data = {}
            if include_source:
                extracted_data['来源表'] = table_name
            
            for field_config in fields:
                source_field = field_config.get('source')
                target_field = field_config.get('target', source_field)
                
                if source_field in df.columns:
                    extracted_data[target_field] = df[source_field]
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
