#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
金华聚火表格处理 - 整合版
将所有功能整合到单个文件中，便于打包为EXE
包含：主GUI + 处理明细表格 + 处理订单表格
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import time
import os
import sys
import shutil
import random
import re
from datetime import datetime, timedelta

# 懒加载 pandas，避免程序启动时阻塞加载所有重型科学计算库
# 只在真正需要处理表格时才导入，GUI 窗口可以立刻出现
class _LazyPandas:
    def __getattr__(self, name):
        import pandas as _pd_mod
        import warnings
        warnings.filterwarnings('ignore')
        globals()['pd'] = _pd_mod
        return getattr(_pd_mod, name)

pd = _LazyPandas()

# ============================================================================
# 处理明细表格功能（原处理明细表格.py）
# ============================================================================

def process_excel_file_detail(input_path, output_path):
    """
    处理单个Excel文件（明细表格处理）
    """
    # 读取Excel文件
    try:
        df = pd.read_excel(input_path, header=None, engine='xlrd')
    except Exception as e:
        print(f"  读取文件失败: {e}")
        return False
    
    # 获取表头行（第1行，索引1）
    if len(df) < 2:
        print(f"  文件行数不足")
        return False
    
    headers = df.iloc[1]
    
    # 确定关键列索引
    col_indices = {
        '商品名称': None,
        '规格': None,
        '单位': None,
        '实际价格': None,
        '订货数量': None,
        '金额': None,
        '备注': None
    }
    
    # 查找列索引
    for col_name in col_indices:
        for i, header in enumerate(headers):
            if isinstance(header, str) and col_name in header:
                col_indices[col_name] = i
                break
    
    # 如果找不到备注列，假设是最后一列
    if col_indices['备注'] is None and len(df.columns) > 0:
        col_indices['备注'] = len(df.columns) - 1
    
    # 检查是否找到所有必要的列
    required_cols = ['商品名称', '规格', '单位', '实际价格', '订货数量', '金额']
    missing_cols = [col for col in required_cols if col_indices[col] is None]
    if missing_cols:
        print(f"  找不到列: {missing_cols}")
        return False
    
    # 处理每一行（从第三行开始，索引2）
    remark_contents = []
    for row_idx in range(2, len(df)):
        # 获取各列的值
        商品名称 = df.iloc[row_idx, col_indices['商品名称']]
        规格 = df.iloc[row_idx, col_indices['规格']]
        单位 = df.iloc[row_idx, col_indices['单位']]
        实际价格 = df.iloc[row_idx, col_indices['实际价格']]
        订货数量 = df.iloc[row_idx, col_indices['订货数量']]
        金额 = df.iloc[row_idx, col_indices['金额']]
        
        # 处理空值
        if pd.isna(商品名称):
            商品名称 = ""
        if pd.isna(规格):
            规格 = ""
        if pd.isna(单位):
            单位 = ""
        
        # 格式化数值
        # 实际价格：如果小数点后超过两位，保留两位；否则保持原样
        if pd.isna(实际价格):
            实际价格_str = ""
        else:
            try:
                # 尝试转换为浮点数
                price_float = float(实际价格)
                # 转换为字符串检查小数位数
                price_str = str(price_float)
                if '.' in price_str:
                    decimal_part = price_str.split('.')[1]
                    # 如果小数部分超过2位，格式化为2位小数
                    if len(decimal_part) > 2:
                        实际价格_str = f"{price_float:.2f}"
                        # 去除不必要的尾随零
                        实际价格_str = 实际价格_str.rstrip('0').rstrip('.') if '.' in 实际价格_str else 实际价格_str
                    else:
                        # 保持原样，去除尾随零
                        实际价格_str = price_str.rstrip('0').rstrip('.') if '.' in price_str else price_str
                else:
                    # 没有小数部分，保持原样
                    实际价格_str = price_str
            except:
                # 如果不是数字，保持原样
                实际价格_str = str(实际价格)
        
        if pd.isna(订货数量):
            订货数量_str = "0"
        else:
            try:
                # 如果是浮点数，转换为整数格式
                num = float(订货数量)
                if num.is_integer():
                    订货数量_str = str(int(num))
                else:
                    订货数量_str = str(num)
            except:
                订货数量_str = "0"
        
        if pd.isna(金额):
            金额_str = "0.00"
        else:
            try:
                金额_str = f"{float(金额):.2f}"
            except:
                金额_str = "0.00"
        
        # 构建格式化字符串
        # 检查是否为赠品（商品名称中包含"赠品"二字）
        if isinstance(商品名称, str) and "赠品" in 商品名称:
            # 赠品格式: "商品名称" + " " + "规格" + " " + "单位" + "：" + "订货数量"+";"
            formatted_str = f"{商品名称} {规格} {单位}：{订货数量_str}；"
        else:
            # 普通商品格式: "商品名称" + " " + "规格" + " " + "单位" + "：" + "实际价格" + "*" + "订货数量" + "=" + " " + "金额" + "；"
            # 注意：冒号和星号的前后都不要空格，等号的前面不要空格，后面保留一个空格
            # 金额已格式化为小数点后两位，空值用"0"填充
            formatted_str = f"{商品名称} {规格} {单位}：{实际价格_str}*{订货数量_str}= {金额_str}；"
        
        # 更新备注列
        df.iloc[row_idx, col_indices['备注']] = formatted_str
        
        # 收集备注内容
        remark_contents.append(formatted_str)
    
    # 添加T列（第20列，索引19）
    if len(df.columns) <= 19:
        # 如果列数不足20，添加空列直到有20列
        while len(df.columns) < 20:
            df[len(df.columns)] = None
    
    # 在T3单元格（行索引2，列索引19）放入所有备注内容
    all_remarks = "\n".join(remark_contents)
    df.iloc[2, 19] = all_remarks  # T3单元格
    
    # 确保输出目录存在
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # 保存到新文件
    try:
        # 使用openpyxl引擎保存以保持格式
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False)
        return True
    except Exception as e:
        print(f"  保存文件失败: {e}")
        return False

def run_detail_processing_integrated():
    """运行明细表格处理（整合版）"""
    input_dir = r'D:\明细表格'
    output_dir = r'D:\明细表格\MXBG2'
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    print(f"输出目录已创建: {output_dir}")
    
    # 获取所有Excel文件
    files = os.listdir(input_dir)
    excel_files = [f for f in files if f.endswith('.xls')]
    
    print(f"找到 {len(excel_files)} 个Excel文件需要处理")
    
    success_count = 0
    for filename in excel_files:
        input_path = os.path.join(input_dir, filename)
        
        # 从文件名中提取数字部分作为新文件名
        # 例如: "销售订单详细 - 2333720569487585.xls" -> "2333720569487585.xls"
        numbers = re.findall(r'\d+', filename)
        if numbers:
            # 使用找到的第一个数字序列作为文件名
            new_filename = f"{numbers[0]}.xls"
        else:
            # 如果没有找到数字，使用原文件名
            new_filename = filename
        
        output_path = os.path.join(output_dir, new_filename)
        
        print(f"\n处理文件: {filename}")
        print(f"  输入: {input_path}")
        print(f"  输出: {output_path} (新文件名: {new_filename})")
        
        if process_excel_file_detail(input_path, output_path):
            print(f"  [OK] 处理成功")
            success_count += 1
        else:
            print(f"  [FAIL] 处理失败")
    
    print(f"\n处理完成: {success_count}/{len(excel_files)} 个文件成功")
    return success_count > 0, f"明细表格处理成功: {success_count}/{len(excel_files)}"

# ============================================================================
# 处理订单表格功能（原处理订单表格.py）
# ============================================================================

def read_excel_file_order(file_path, sheet_name=0, engine=None):
    """读取Excel文件，自动选择合适的引擎"""
    try:
        if engine:
            return pd.read_excel(file_path, sheet_name=sheet_name, engine=engine, header=None)
        else:
            # 尝试自动选择引擎
            if file_path.endswith('.xlsx'):
                return pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
            elif file_path.endswith('.xls'):
                # 对于.xls文件，先尝试xlrd，如果失败则尝试openpyxl
                try:
                    return pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd', header=None)
                except:
                    # 如果xlrd失败，可能是.xlsx格式但扩展名是.xls
                    return pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
            else:
                # 默认使用openpyxl
                return pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
    except Exception as e:
        raise Exception(f"读取文件失败 {file_path}: {e}")

def save_excel_file_order(df, file_path):
    """保存Excel文件"""
    try:
        # 确保输出目录存在
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        # 使用openpyxl引擎保存以保持格式
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False)
        return True
    except Exception as e:
        print(f"保存文件失败 {file_path}: {e}")
        return False

def extract_cell_content_order(content):
    """提取单元格内容，处理空值和特殊字符"""
    if pd.isna(content):
        return ""
    return str(content).strip()

def extract_phone_from_contact_info(text):
    """从联系信息中提取手机号码（最后一个'-'之后的11位数字）"""
    if not text or '-' not in text:
        return ""
    parts = text.rsplit('-', 1)
    phone = parts[-1].strip()
    if phone.isdigit() and len(phone) == 11 and phone.startswith('1'):
        return phone
    return ""


def extract_contact_name_from_contact_info(text):
    """从联系信息中提取联系人姓名（最后一个'-'之前的部分）"""
    if not text:
        return ""
    if '-' not in text:
        return text.strip()
    return text.rsplit('-', 1)[0].strip()


def process_sales_order_integrated(sales_order_path, template_path):
    """
    处理销售订单文件，将数据复制到模板中
    """
    print("步骤1: 处理销售订单数据...")
    
    # 读取销售订单文件
    sales_df = read_excel_file_order(sales_order_path)
    if sales_df is None:
        raise Exception(f"无法读取销售订单文件，请检查文件是否存在且格式正确:\n{sales_order_path}")

    # 读取模板文件
    template_df = read_excel_file_order(template_path, engine='openpyxl')
    if template_df is None:
        raise Exception(f"无法读取优路达导入模板，请检查文件是否存在且格式正确:\n{template_path}")
    
    # 查找销售订单中的列索引
    column_mapping = {}
    header_row = None
    
    # 尝试找到表头行（通常在第一行或第二行）
    for row_idx in range(min(5, len(sales_df))):
        row = sales_df.iloc[row_idx]
        # 检查是否包含关键字段
        key_fields = ['客户单位', '地址', '订单号', '联系信息', '备注']
        found_count = 0
        for cell in row:
            if isinstance(cell, str):
                for field in key_fields:
                    if field in cell:
                        found_count += 1
                        break
        
        if found_count >= 3:  # 如果找到至少3个关键字段，认为是表头行
            header_row = row_idx
            break
    
    if header_row is None:
        print("未找到表头行，使用第一行作为表头")
        header_row = 0
    
    # 构建列索引映射
    for col_idx in range(len(sales_df.columns)):
        cell_value = sales_df.iloc[header_row, col_idx]
        if isinstance(cell_value, str):
            cell_str = str(cell_value).strip()
            if '客户单位' in cell_str:
                column_mapping['客户单位'] = col_idx
            elif '地址' in cell_str:
                column_mapping['地址'] = col_idx
            elif '销售单号' in cell_str:
                column_mapping['订单号'] = col_idx  # 销售单号就是订单号
            elif '联系信息' in cell_str:
                column_mapping['联系信息'] = col_idx
            elif '订单备注' in cell_str:
                column_mapping['订单备注'] = col_idx
    
    print(f"找到的列映射: {column_mapping}")

    # 如果没找到联系信息，尝试其他可能的列名
    if '联系信息' not in column_mapping:
        for col_idx in range(len(sales_df.columns)):
            cell_value = sales_df.iloc[header_row, col_idx]
            if isinstance(cell_value, str):
                cell_str = str(cell_value).strip()
                if '联系' in cell_str or '联系电话' in cell_str or '联系人' in cell_str:
                    column_mapping['联系信息'] = col_idx
                    print(f"  使用替代列名找到联系信息: {cell_str}")
                    break

    # 如果没找到订单备注，尝试查找备注列
    if '订单备注' not in column_mapping:
        for col_idx in range(len(sales_df.columns)):
            cell_value = sales_df.iloc[header_row, col_idx]
            if isinstance(cell_value, str):
                cell_str = str(cell_value).strip()
                if '备注' in cell_str:
                    column_mapping['订单备注'] = col_idx
                    print(f"  使用替代列名找到订单备注: {cell_str}")
                    break
    
    # 如果没找到订单号，尝试其他可能的列名
    if '订单号' not in column_mapping:
        for col_idx in range(len(sales_df.columns)):
            cell_value = sales_df.iloc[header_row, col_idx]
            if isinstance(cell_value, str):
                cell_str = str(cell_value).strip()
                if '单号' in cell_str:
                    column_mapping['订单号'] = col_idx
                    print(f"  使用替代列名找到订单号: {cell_str}")
                    break
    
    # 查找模板中的列索引
    template_columns = {}
    for col_idx in range(len(template_df.columns)):
        cell_value = template_df.iloc[0, col_idx]  # 假设第一行是表头
        if isinstance(cell_value, str):
            cell_str = str(cell_value).strip()
            if '客户名称' in cell_str:
                template_columns['客户名称'] = col_idx
            elif '详细地址' in cell_str:
                template_columns['详细地址'] = col_idx
            elif '订单号' in cell_str:
                template_columns['订单号'] = col_idx
            elif '手机号' in cell_str:
                template_columns['手机号'] = col_idx
            elif '客户经理' in cell_str:
                template_columns['客户经理'] = col_idx
            elif '备注信息' in cell_str:
                template_columns['备注信息'] = col_idx
            elif '地点名称' in cell_str:
                template_columns['地点名称'] = col_idx
            elif '精准坐标' in cell_str:
                template_columns['精准坐标'] = col_idx
    
    print(f"模板列映射: {template_columns}")
    
    # 准备结果数据
    result_data = []
    
    # 从表头行之后开始处理数据
    for row_idx in range(header_row + 1, len(sales_df)):
        row_data = {}
        
        # 提取各字段数据
        for field, col_idx in column_mapping.items():
            if col_idx < len(sales_df.columns):
                value = extract_cell_content_order(sales_df.iloc[row_idx, col_idx])
                row_data[field] = value
            else:
                row_data[field] = ""
        
        # 跳过空行（所有字段都为空）
        if all(value == "" for value in row_data.values()):
            continue
        
        # 处理地点名称：
        # 1. 如果有"★"号，选择"★"号之后的内容
        # 2. 如果没有"★"号，则选择第一个空格之后的内容
        # 3. 如果都没有，则使用原内容
        customer_unit = row_data.get('客户单位', '')
        location_name = customer_unit
        
        if '★' in customer_unit:
            parts = customer_unit.split('★')
            if len(parts) > 1:
                location_name = parts[1].strip()
        elif ' ' in customer_unit:
            parts = customer_unit.split(' ', 1)  # 只分割第一个空格
            if len(parts) > 1:
                location_name = parts[1].strip()
        
        # 处理订单备注：如果不为空则前面加"备注："
        order_remark = row_data.get('订单备注', '')
        if order_remark and order_remark.strip():
            # 处理多行备注：为每一行都加上"备注："前缀
            lines = order_remark.strip().split('\n')
            processed_lines = []
            for line in lines:
                line = line.strip()
                if line:  # 忽略空行
                    processed_lines.append(f"备注：{line}")
            processed_remark = '\n'.join(processed_lines)
        else:
            processed_remark = ""
        
        # 构建结果行
        result_row = {
            '客户名称': row_data.get('客户单位', ''),
            '详细地址': row_data.get('地址', ''),
            '订单号': str(row_data.get('订单号', '')),  # 转换为文字格式
            '手机号': extract_phone_from_contact_info(row_data.get('联系信息', '')),
            '客户经理': extract_contact_name_from_contact_info(row_data.get('联系信息', '')),
            '备注信息': processed_remark,
            '地点名称': location_name,
            '精准坐标': ''  # 留空，后续填充
        }
        
        result_data.append(result_row)
    
    print(f"处理了 {len(result_data)} 行数据")
    return result_data, template_df, template_columns

def match_customer_coordinates_integrated(customer_coords_path, result_data):
    """
    匹配客户经纬度信息
    """
    print("步骤2: 匹配客户经纬度信息...")
    
    # 读取客户经纬度文件
    coords_df = read_excel_file_order(customer_coords_path)
    if coords_df is None:
        print("无法读取客户经纬度文件，跳过此步骤")
        return result_data, []
    
    # 查找客户经纬度文件中的列索引
    coords_columns = {}
    for col_idx in range(len(coords_df.columns)):
        cell_value = coords_df.iloc[0, col_idx]  # 假设第一行是表头
        if isinstance(cell_value, str):
            cell_str = str(cell_value).strip()
            if '客户单位' in cell_str:
                coords_columns['客户单位'] = col_idx
            elif '地图经纬度' in cell_str:
                coords_columns['地图经纬度'] = col_idx
    
    if '客户单位' not in coords_columns or '地图经纬度' not in coords_columns:
        print("客户经纬度文件中未找到必要的列，跳过此步骤")
        return result_data, []
    
    print(f"客户经纬度列映射: {coords_columns}")
    
    # 构建客户单位到经纬度的映射
    customer_coords_map = {}
    for row_idx in range(1, len(coords_df)):  # 从第二行开始
        customer_name = extract_cell_content_order(coords_df.iloc[row_idx, coords_columns['客户单位']])
        coordinates = extract_cell_content_order(coords_df.iloc[row_idx, coords_columns['地图经纬度']])
        
        if customer_name and coordinates:
            customer_coords_map[customer_name] = coordinates
    
    print(f"建立了 {len(customer_coords_map)} 个客户经纬度映射")
    
    # 匹配并填充精准坐标
    matched_count = 0
    unmatched_names = []
    for row in result_data:
        customer_name = row.get('客户名称', '')
        if customer_name in customer_coords_map:
            row['精准坐标'] = customer_coords_map[customer_name]
            matched_count += 1
        elif customer_name.strip():
            unmatched_names.append(customer_name)

    print(f"成功匹配了 {matched_count} 个客户的经纬度信息")
    if unmatched_names:
        print(f"未匹配到经纬度的客户 ({len(unmatched_names)}): {', '.join(unmatched_names[:10])}...")
    return result_data, unmatched_names

def match_detail_files_integrated(result_data, detail_dir):
    """
    匹配明细表格文件
    """
    print("步骤3: 匹配明细表格文件...")
    
    if not os.path.exists(detail_dir):
        print(f"明细表格目录不存在: {detail_dir}，跳过此步骤")
        return result_data, 0, 0
    
    # 获取所有Excel文件，排除临时文件
    excel_files = []
    for file in os.listdir(detail_dir):
        if (file.endswith('.xls') or file.endswith('.xlsx')) and not file.startswith('~$'):
            excel_files.append(file)
    
    print(f"在明细表格目录中找到 {len(excel_files)} 个Excel文件")
    
    # 构建订单号到文件路径的映射
    order_file_map = {}
    for file in excel_files:
        # 从文件名中提取数字
        numbers = re.findall(r'\d+', file)
        for num in numbers:
            order_file_map[num] = os.path.join(detail_dir, file)
    
    # 处理每个订单
    processed_count = 0
    for row in result_data:
        order_no = str(row.get('订单号', '')).strip()
        if not order_no:
            continue
        
        # 查找匹配的文件
        matched_file = None
        for num, file_path in order_file_map.items():
            if num in order_no or order_no in num:
                matched_file = file_path
                break
        
        if matched_file:
            # 读取匹配的文件并获取T3单元格内容
            try:
                detail_df = read_excel_file_order(matched_file)
                    
                if detail_df is not None and len(detail_df.columns) > 19 and len(detail_df) > 2:
                    t3_content = extract_cell_content_order(detail_df.iloc[2, 19])  # T3单元格
                    if t3_content:
                        # 在备注信息最前面插入T3内容并增加两个换行
                        current_remark = row.get('备注信息', '')
                        row['备注信息'] = f"{t3_content}\n\n{current_remark}"
                        processed_count += 1
                        print(f"  订单 {order_no}: 成功匹配并添加明细信息")
            except Exception as e:
                print(f"读取明细文件失败 {matched_file}: {e}")
    
    total_files = len(excel_files)
    print(f"成功处理了 {processed_count} 个订单的明细信息")
    return result_data, total_files, processed_count

def merge_same_customers_integrated(result_data):
    """
    合并相同客户的数据
    """
    print("步骤4: 合并相同客户数据...")
    
    # 按客户名称分组
    customer_groups = {}
    for row in result_data:
        customer_name = row.get('客户名称', '')
        if customer_name not in customer_groups:
            customer_groups[customer_name] = []
        customer_groups[customer_name].append(row)
    
    # 合并每组数据
    merged_data = []
    for customer_name, rows in customer_groups.items():
        if len(rows) == 1:
            # 只有一个客户，直接添加
            merged_data.append(rows[0])
        else:
            # 合并多个客户行
            merged_row = {}
            
            # 获取第一个行的数据作为基础（除了订单号和备注信息）
            first_row = rows[0]
            for key in first_row:
                if key not in ['订单号', '备注信息']:
                    merged_row[key] = first_row[key]
            
            # 合并订单号
            order_nos = []
            for row in rows:
                order_no = row.get('订单号', '')
                if order_no and order_no not in order_nos:
                    order_nos.append(order_no)
            merged_row['订单号'] = ','.join(order_nos)
            
            # 合并备注信息：进行换行后合并操作
            remarks = []
            for row in rows:
                remark = row.get('备注信息', '')
                if remark:
                    remarks.append(remark)
            
            # 合并之后再换行然后插入（"合并订单："合并订单的数量）
            merged_remarks = '\n'.join(remarks)
            if merged_remarks:
                merged_remarks += f'\n合并订单：{len(rows)}'
            
            merged_row['备注信息'] = merged_remarks
            
            merged_data.append(merged_row)
    
    print(f"合并后数据: {len(result_data)} -> {len(merged_data)} 行")
    return merged_data

def process_remarks_integrated(result_data):
    """
    处理备注信息单元格的内容
    """
    print("步骤5: 处理备注信息...")
    
    for row in result_data:
        remark = row.get('备注信息', '')
        if not remark:
            continue
        
        # 分割成行
        lines = remark.split('\n')
        
        # 提取T3内容（包含"="的行）、备注内容和合并行
        t3_lines = []
        remark_lines = []
        merge_lines = []
        other_lines = []
        
        for line in lines:
            if '=' in line and '；' in line:
                t3_lines.append(line)
            elif line.startswith('备注：'):
                remark_lines.append(line)
            elif line.startswith('合并订单：'):
                merge_lines.append(line)
            elif line.strip() and line != '-' * 10:  # 忽略分隔线
                other_lines.append(line)
        
        # 计算T3内容中的金额总和
        total_amount = 0.0
        for line in t3_lines:
            # 查找"="后的数字
            if '=' in line:
                parts = line.split('=')
                if len(parts) > 1:
                    amount_str = parts[1].strip().rstrip('；')
                    try:
                        # 移除可能的空格和特殊字符
                        amount_str_clean = amount_str.replace(' ', '').replace(',', '')
                        amount = float(amount_str_clean)
                        total_amount += amount
                    except:
                        pass
        
        # 对备注行进行编号（只有在有多个备注行时才编号）
        numbered_remark_lines = []
        if len(remark_lines) > 1:
            # 有多个备注行，进行编号
            for i, remark_line in enumerate(remark_lines):
                if i == 0:
                    # 第一行：备注：1、
                    numbered_remark_lines.append(f"备注：{i+1}、{remark_line[3:]}")
                else:
                    # 其他行：前面加6个空格，然后编号
                    numbered_remark_lines.append(f"      {i+1}、{remark_line[3:]}")
        else:
            # 只有一个备注行或没有备注行，保持原样
            numbered_remark_lines = remark_lines
        
        # 重新构建备注内容
        new_remark_parts = []
        
        # 1. 添加T3内容
        if t3_lines:
            new_remark_parts.extend(t3_lines)
        
        # 2. 添加空白行和总金额
        if t3_lines:
            new_remark_parts.append('')  # 空白行
            new_remark_parts.append(f'总金额：{total_amount:.2f}元')
        
        # 3. 添加合并行（如果有） - 注意：合并行和总金额行之间不要有空行
        if merge_lines:
            # 直接添加合并行，不在总金额和合并行之间添加空行
            new_remark_parts.extend(merge_lines)
        
        # 4. 添加其他内容
        if other_lines:
            new_remark_parts.append('')  # 空白行
            new_remark_parts.extend(other_lines)
        
        # 5. 添加编号后的备注内容到最下面
        if numbered_remark_lines:
            new_remark_parts.append('')  # 空白行
            new_remark_parts.extend(numbered_remark_lines)
        
        row['备注信息'] = '\n'.join(new_remark_parts)
    
    return result_data

def add_customer_info_to_remarks_integrated(result_data):
    """
    将每行的客户信息添加到备注信息中
    """
    print("步骤6: 添加客户信息到备注...")
    
    # 获取当前日期
    # 如果运行时间在12点之后24点之前，则使用第二天的日期
    now = datetime.now()
    if now.hour >= 12 and now.hour < 24:
        # 使用明天的日期
        from datetime import timedelta
        tomorrow = now + timedelta(days=1)
        current_date = tomorrow.strftime('%Y年%m月%d日')
    else:
        # 使用今天的日期
        current_date = now.strftime('%Y年%m月%d日')
    
    for row in result_data:
        # 构建客户信息字符串
        customer_info_parts = []
        for field in ['客户名称', '手机号', '详细地址', '客户经理']:
            value = row.get(field, '')
            if value:
                customer_info_parts.append(str(value))
        
        customer_info = '，'.join(customer_info_parts)
        
        # 获取当前备注信息
        current_remark = row.get('备注信息', '')
        
        # 构建新的备注信息
        new_remark_parts = []
        
        # 1. 添加客户信息作为第一行
        if customer_info:
            new_remark_parts.append(customer_info)
            new_remark_parts.append('')  # 空白行
        
        # 2. 添加原有备注内容
        if current_remark:
            new_remark_parts.append(current_remark)
        
        # 3. 在所有内容后添加空白行
        new_remark_parts.append('')
        
        # 4. 添加聚火配送信息
        new_remark_parts.append(f'聚火配送：{current_date}')
        
        row['备注信息'] = '\n'.join(new_remark_parts)
    
    return result_data

def create_final_template_integrated(result_data, template_df, template_columns):
    """
    创建最终的模板文件
    """
    print("步骤7: 创建最终模板...")
    
    # 创建新的DataFrame，保持与模板相同的列数
    result_df = template_df.copy()
    
    # 清空模板中的数据行（保留表头）
    # 假设第一行是表头
    if len(result_df) > 1:
        result_df = result_df.iloc[[0]]  # 只保留第一行（表头）
    
    # 添加数据行
    for row_data in result_data:
        new_row = ['' for _ in range(len(result_df.columns))]
        
        # 填充数据到对应的列
        for field, col_idx in template_columns.items():
            if field in row_data and col_idx < len(new_row):
                new_row[col_idx] = row_data[field]
        
        # 将新行添加到DataFrame
        result_df.loc[len(result_df)] = new_row
    
    return result_df

def save_final_file_integrated(result_df, output_dir, base_filename="优路达导入模板"):
    """
    保存最终文件
    """
    # 获取当前日期
    # 如果运行时间在12点之后24点之前，则使用第二天的日期
    now = datetime.now()
    if now.hour >= 12 and now.hour < 24:
        # 使用明天的日期
        target_date = now + timedelta(days=1)
    else:
        # 使用今天的日期
        target_date = now
    
    month = target_date.month
    day = target_date.day
    
    # 生成基础随机序号（1000-9999）
    base_random = random.randint(1000, 9999)
    
    # 检查同一天已经存在的文件
    existing_files = []
    
    if os.path.exists(output_dir):
        for file in os.listdir(output_dir):
            if file.startswith(f"{base_filename}{month:02d}{day:02d}-") and file.endswith('.xlsx'):
                existing_files.append(file)
    
    # 计算递增序号
    if existing_files:
        # 提取现有文件中的最大序号
        max_suffix = 0
        for file in existing_files:
            try:
                # 从文件名中提取序号部分
                suffix_part = file.split('-')[1].split('.')[0]
                suffix_num = int(suffix_part)
                if suffix_num > max_suffix:
                    max_suffix = suffix_num
            except:
                pass
        
        # 使用最大序号+1作为新序号
        # 如果最大序号小于基础随机数，使用基础随机数
        if max_suffix < base_random:
            suffix = base_random
        else:
            suffix = max_suffix + 1
    else:
        # 没有现有文件，使用基础随机数
        suffix = base_random
    
    # 确保序号在1000-9999范围内
    if suffix > 9999:
        suffix = 1000 + (suffix % 9000)
    
    # 构建文件名
    filename = f"{base_filename}{month:02d}{day:02d}-{suffix}.xlsx"
    file_path = os.path.join(output_dir, filename)
    
    print(f"保存文件: {file_path}")
    print(f"  基础随机数: {base_random}, 最终序号: {suffix}, 现有文件数: {len(existing_files)}")
    
    # 保存文件
    if save_excel_file_order(result_df, file_path):
        print(f"文件保存成功: {filename}")
        return file_path
    else:
        print(f"文件保存失败: {filename}")
        return None

def _dwidth(s):
    """Microsoft YaHei 下的显示宽度，中文≈3，ASCII≈2，零宽字符=0"""
    w = 0
    for c in s:
        o = ord(c)
        if o in (0x200B, 0x200C, 0x200D, 0xFEFF, 0x200E, 0x200F):
            continue
        if o in (0xFF08, 0xFF09):  # 中文括号（）宽度为1
            w += 1
        else:
            w += 4 if o > 0x2E80 else 2
    return w


def _pad(s, target_w):
    """将字符串 s 补空格到显示宽度 target_w"""
    cur = _dwidth(s)
    return s + ' ' * max(0, target_w - cur)


def compute_detail_stats(input_dir):
    """从原始明细表格统计商品，返回 (格式化文本行列表, 总数)"""
    from collections import defaultdict
    stats = defaultdict(int)
    if not os.path.exists(input_dir):
        return [], 0, 46, 36, 10, []
    for filename in os.listdir(input_dir):
        if not (filename.endswith('.xls') or filename.endswith('.xlsx')):
            continue
        if filename.startswith('~$'):
            continue
        filepath = os.path.join(input_dir, filename)
        try:
            df = pd.read_excel(filepath, header=None, engine='xlrd')
        except:
            try:
                df = pd.read_excel(filepath, header=None, engine='openpyxl')
            except:
                continue
        if len(df) < 2:
            continue
        headers = df.iloc[1]
        col_map = {}
        for i, h in enumerate(headers):
            if isinstance(h, str):
                for key in ('商品名称', '规格', '单位', '订货数量'):
                    if key in h and key not in col_map:
                        col_map[key] = i
        if len(col_map) < 4:
            continue
        for r in range(2, len(df)):
            name = str(df.iloc[r, col_map['商品名称']]) if not pd.isna(df.iloc[r, col_map['商品名称']]) else ''
            spec = str(df.iloc[r, col_map['规格']]) if not pd.isna(df.iloc[r, col_map['规格']]) else ''
            unit = str(df.iloc[r, col_map['单位']]) if not pd.isna(df.iloc[r, col_map['单位']]) else ''
            qty_val = df.iloc[r, col_map['订货数量']]
            if pd.isna(qty_val):
                continue
            try:
                qty = int(float(qty_val))
            except:
                continue
            stats[(name.strip(), spec.strip(), unit.strip())] += qty
    if not stats:
        return [], 0, 46, 36, 10, []
    items = sorted(stats.items(), key=lambda x: x[0][0])
    total_qty = sum(v for _, v in items)
    W_NAME, W_SPEC, W_UNIT = 46, 36, 10
    lines = []
    for (name, spec, unit), qty in items:
        # 超过最大宽度的截断
        while _dwidth(name) > W_NAME:
            name = name[:-1]
        while _dwidth(spec) > W_SPEC:
            spec = spec[:-1]
        while _dwidth(unit) > W_UNIT:
            unit = unit[:-1]
        lines.append(
            f"  {_pad(name, W_NAME)}{_pad(spec, W_SPEC)}{_pad(unit, W_UNIT)}{qty:>6}"
        )
    raw_items = [(k[0], k[1], k[2], v) for k, v in items]
    return lines, total_qty, W_NAME, W_SPEC, W_UNIT, raw_items


def count_total_products(detail_dir):
    """统计 MXBG2 目录下所有明细表格的订货数量总和"""
    total = 0
    if not os.path.exists(detail_dir):
        return 0
    for filename in os.listdir(detail_dir):
        if not (filename.endswith('.xls') or filename.endswith('.xlsx')):
            continue
        if filename.startswith('~$'):
            continue
        filepath = os.path.join(detail_dir, filename)
        try:
            df = pd.read_excel(filepath, header=None, engine='xlrd')
        except:
            try:
                df = pd.read_excel(filepath, header=None, engine='openpyxl')
            except:
                continue
        if len(df) < 2:
            continue
        # 找表头行中的"订货数量"列
        headers = df.iloc[1]
        qty_col = None
        for i, h in enumerate(headers):
            if isinstance(h, str) and '订货数量' in h:
                qty_col = i
                break
        if qty_col is None:
            continue
        # 从第3行开始累加
        for r in range(2, len(df)):
            val = df.iloc[r, qty_col]
            if pd.isna(val):
                continue
            try:
                total += int(float(val))
            except:
                pass
    return total


def run_order_processing_integrated():
    """运行订单表格处理（整合版）"""
    print("=" * 60)
    print("开始处理订单表格")
    print("=" * 60)
    
    # 定义文件路径
    sales_order_path = r'D:\订单表格\销售订单.xls'
    customer_coords_path = r'D:\订单表格\客户经纬度.xls'
    template_path = r'D:\订单表格\优路达导入模板.xlsx'
    detail_dir = r'D:\明细表格\MXBG2'
    output_dir = r'D:\订单表格'
    
    # 检查必要文件是否存在
    required_files = [
        (sales_order_path, '销售订单.xls'),
        (customer_coords_path, '客户经纬度.xls'),
        (template_path, '优路达导入模板.xlsx')
    ]
    
    missing_files = []
    for file_path, file_name in required_files:
        if not os.path.exists(file_path):
            missing_files.append(file_name)
    
    if missing_files:
        error_msg = f"错误：缺少必要的文件: {', '.join(missing_files)}\n请确保以下文件存在于 D:\\订单表格 目录中:\n" + "\n".join([f"  - {file_name}" for file_name in missing_files])
        print(error_msg)
        return False, error_msg
    
    print("所有必要文件都存在，开始处理...")
    
    # 步骤1: 处理销售订单数据
    try:
        result = process_sales_order_integrated(sales_order_path, template_path)
    except Exception as e:
        error_msg = f"处理销售订单数据失败:\n{str(e)}"
        print(error_msg)
        return False, error_msg

    if result is None:
        error_msg = "处理销售订单数据失败（未知原因）"
        print(error_msg)
        return False, error_msg
    
    result_data, template_df, template_columns = result
    
    if not result_data:
        error_msg = "没有找到可处理的数据"
        print(error_msg)
        return False, error_msg
    
    # 步骤2: 匹配客户经纬度信息
    result_data, unmatched_coords = match_customer_coordinates_integrated(customer_coords_path, result_data)
    
    # 步骤3: 匹配明细表格文件
    result_data, total_detail_files, detail_matched = match_detail_files_integrated(result_data, detail_dir)
    detail_warning = ""
    if total_detail_files == 0:
        detail_warning = f"⚠ 明细表格目录中没有 Excel 文件（{detail_dir}），订单备注将缺少明细信息。"
    elif detail_matched == 0:
        detail_warning = f"⚠ 明细目录有 {total_detail_files} 个文件但未匹配到任何订单，请检查。"
    if detail_warning:
        print(detail_warning)
    
    # 步骤4: 合并相同客户数据
    total_orders_before = len(result_data)
    result_data = merge_same_customers_integrated(result_data)
    total_customers = len(result_data)

    # 收集未匹配经纬度的客户（精准坐标为空的）
    customers_no_coords = [
        (row.get('地点名称', ''), row.get('客户名称', ''))
        for row in result_data
        if not row.get('精准坐标', '').strip()
    ]
    
    # 步骤5: 处理备注信息
    result_data = process_remarks_integrated(result_data)
    
    # 步骤6: 添加客户信息到备注
    result_data = add_customer_info_to_remarks_integrated(result_data)
    
    # 步骤7: 创建最终模板
    final_df = create_final_template_integrated(result_data, template_df, template_columns)
    
    # 步骤8: 保存最终文件
    output_file = save_final_file_integrated(final_df, output_dir)
    
    if output_file:
        success_msg = (
            f"处理完成！\n"
            f"生成文件: {os.path.basename(output_file)}\n"
            f"──────────────────────\n"
            f"处理订单数: {total_orders_before} 单\n"
            f"匹配明细: {detail_matched} 单\n"
            f"合并后客户: {total_customers} 个\n"
        )
        if detail_matched != total_orders_before:
            success_msg += "\x02"  # 标记匹配明细行需要标红

        total_products = count_total_products(detail_dir)
        if total_products > 0:
            success_msg += f"出库商品总数 {total_products}\n"

        # 未匹配经纬度的客户
        if customers_no_coords:
            success_msg += f"\n⚠ 未匹配经纬度的客户 ({len(customers_no_coords)} 个):\n"
            for loc, name in customers_no_coords[:30]:
                display = loc if loc else name
                success_msg += f"  • {display}\n"
            if len(customers_no_coords) > 30:
                success_msg += f"  ...等共{len(customers_no_coords)}个\n"

        if detail_warning:
            success_msg += f"\n{detail_warning}"
        print("\n" + "=" * 60)
        print("处理完成！")
        print(f"生成的文件: {output_file}")
        print("=" * 60)
        
        # 显示处理统计
        print(f"\n处理统计:")
        print(f"- 处理数据行数: {len(result_data)}")
        print(f"- 输出文件: {os.path.basename(output_file)}")
        print(f"- 输出目录: {os.path.dirname(output_file)}")
        return True, success_msg
    else:
        error_msg = "处理失败：无法保存输出文件"
        print("\n处理失败！")
        return False, error_msg

# ============================================================================
# 主GUI应用程序
# ============================================================================

class PrintToWidget:
    """将 print 输出重定向到 tkinter Text 控件"""
    def __init__(self, text_widget):
        self.text = text_widget
    def write(self, s):
        self.text.after(0, self._write, s)
    def _write(self, s):
        self.text.config(state=tk.NORMAL)
        self.text.insert(tk.END, s)
        self.text.see(tk.END)
    def flush(self):
        pass


class JinhuaJuhuoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("金华聚火表格处理")
        self.root.geometry("580x720")  # 固定窗口大小
        self.root.resizable(False, False)
        
        # 设置窗口图标
        for icon_path in [
            os.path.join(os.path.dirname(os.path.abspath(__file__)), 'favicon.ico'),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), '_internal', 'favicon.ico'),
            r'D:\WFR\D2Y\favicon.ico'
        ]:
            if os.path.exists(icon_path):
                self.root.iconbitmap(default=icon_path)
                break

        # 隐藏控制台窗口（Windows特定）
        if sys.platform == "win32":
            try:
                import ctypes
                ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
            except:
                pass
        
        # 初始化变量
        self.is_running = False
        self.is_loaded = False  # pandas 后台加载状态
        self.result_message = ""  # 成功后的结果消息（含警告）
        self.output_file_path = ""  # 生成的文件路径
        self.detail_stats_lines = None  # 商品明细统计格式化的行
        self.detail_stats_widths = (24, 14, 8)
        self.detail_stats_raw = []
        self.detail_stats_total = 0
        self.delete_files_var = tk.BooleanVar(value=True)
        self.show_detail_stats_var = tk.BooleanVar(value=True)

        # 创建UI
        self.setup_ui()

        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 窗口显示后再启动后台加载（after 让窗口先渲染）
        self.root.after(50, self._preload_pandas)
    
    def setup_ui(self):
        """设置用户界面"""
        # 创建主框架，便于布局控制
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 标题标签
        title_label = tk.Label(
            main_frame, 
            text="金华聚火表格处理", 
            font=("Microsoft YaHei", 18, "bold"),
            foreground="#2c3e50"
        )
        title_label.pack(pady=(0, 2))
        tk.Label(main_frame,
                 text="订小易_转_优路达_表格处理",
                 font=("Microsoft YaHei", 9),
                 foreground="#95a5a6").pack(pady=(0, 15))
        
        # 说明标签
        info_label = tk.Label(
            main_frame,
            text="请确保以下目录中存在相应文件：\n1、D:\\订单表格\\销售订单.xls\n2、D:\\订单表格\\客户经纬度.xls\n3、D:\\订单表格\\优路达导入模板.xlsx\n4、D:\\明细表格\\销售订单详细 - *.xls",
            font=("Microsoft YaHei", 9),
            foreground="#7f8c8d",
            justify=tk.LEFT,
            wraplength=450,
            bg="#f9f9f9",
            padx=10,
            pady=5
        )
        info_label.pack(pady=(0, 10), fill=tk.X)
        
        # 删除文件勾选框
        checkbox_frame = tk.Frame(main_frame)
        checkbox_frame.pack(pady=(0, 15), fill=tk.X)
        
        self.delete_checkbox = tk.Checkbutton(
            checkbox_frame,
            text="执行完毕后删除 D:\\明细表格 目录下的所有文件（包括子目录）",
            variable=self.delete_files_var,
            onvalue=True,
            offvalue=False,
            font=("Microsoft YaHei", 10),
            bg="#f0f0f0",
            activebackground="#f0f0f0",
            wraplength=450,
            justify=tk.LEFT,
            anchor="w"
        )
        self.delete_checkbox.pack(anchor="w")

        # 查看商品明细勾选框
        self.detail_stats_checkbox = tk.Checkbutton(
            checkbox_frame,
            text="查看商品明细数量",
            variable=self.show_detail_stats_var,
            onvalue=True,
            offvalue=False,
            font=("Microsoft YaHei", 10),
            bg="#f0f0f0",
            activebackground="#f0f0f0",
            anchor="w"
        )
        self.detail_stats_checkbox.pack(anchor="w")
        
        # 执行按钮
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=(0, 20))
        
        self.execute_button = tk.Button(
            button_frame,
            text="加载中...",
            command=self.execute_conversion,
            font=("Microsoft YaHei", 14),
            bg="#95a5a6",
            fg="white",
            width=20,
            height=2,
            state=tk.DISABLED
        )
        self.execute_button.pack()
        
        # 进度条框架
        progress_frame = tk.LabelFrame(main_frame, text="进度", font=("Microsoft YaHei", 10))
        progress_frame.pack(pady=(0, 15), fill=tk.X, padx=10)
        
        # 进度标签
        self.progress_label = tk.Label(
            progress_frame, 
            text="就绪", 
            font=("Microsoft YaHei", 11),
            foreground="#2980b9"
        )
        self.progress_label.pack(pady=(10, 5))
        
        # 进度条
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            length=420,
            mode='determinate'
        )
        self.progress_bar.pack(pady=(0, 5), padx=20)
        
        # 进度百分比标签
        self.percentage_label = tk.Label(
            progress_frame,
            text="0%",
            font=("Microsoft YaHei", 10),
            foreground="#7f8c8d"
        )
        self.percentage_label.pack(pady=(0, 10))
        
        # 结果文本区
        self.result_text = scrolledtext.ScrolledText(
            main_frame,
            font=("Microsoft YaHei", 11),
            foreground="#2c3e50",
            bg="#f9f9f9",
            height=14,
            wrap=tk.NONE,
            state=tk.DISABLED
        )
        self.result_text.tag_config("red", foreground="#e74c3c", font=("Microsoft YaHei", 11, "bold"))
        self.result_text.tag_config("head", font=("Microsoft YaHei", 11, "bold"))
        self.result_text.tag_config("stripe", background="#f5f5f5")
        self.result_text.tag_config("link", foreground="#2980b9", underline=True)
        self.result_text.bind("<Key>", lambda e: "break")
        self.result_text.pack(pady=(0, 2), fill=tk.BOTH, expand=True, padx=10)



        # 状态标签
        self.status_label = tk.Label(
            main_frame,
            text="等待执行...",
            font=("Microsoft YaHei", 10),
            foreground="#95a5a6",
            height=2
        )
        self.status_label.pack(pady=(0, 5), fill=tk.X)
    
    def _preload_pandas(self):
        """后台线程预加载 pandas 全家桶，完成后激活按钮"""
        self.status_label.config(text="正在加载组件...")

        def _load():
            import pandas as _pd
            import warnings
            warnings.filterwarnings('ignore')
            globals()['pd'] = _pd  # 替换懒加载代理

        def _on_done():
            self.is_loaded = True
            self.execute_button.config(
                state=tk.NORMAL, text="执行表格转换",
                bg="#3498db"
            )
            self.status_label.config(text="就绪，等待执行...")

        def _run():
            _load()
            self.root.after(0, _on_done)

        thread = threading.Thread(target=_run, daemon=True)
        thread.start()

    def execute_conversion(self):
        """执行表格转换"""
        if self.is_running or not self.is_loaded:
            return
        
        self.is_running = True
        self.execute_button.config(state=tk.DISABLED, text="执行中...")
        self.delete_checkbox.config(state=tk.DISABLED)
        self.detail_stats_checkbox.config(state=tk.DISABLED)
        self.output_file_path = ""
        if hasattr(self, '_btn_row'):
            self._btn_row.destroy()
        self.execute_button.pack()
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete("1.0", tk.END)
        # 重定向 print 到结果窗口
        self._old_stdout = sys.stdout
        sys.stdout = PrintToWidget(self.result_text)
        
        # 在新线程中执行任务
        thread = threading.Thread(target=self.run_conversion_task)
        thread.daemon = True
        thread.start()
    
    def run_conversion_task(self):
        """运行转换任务"""
        try:
            # 重置进度条
            self.update_progress(0, "开始执行...", "0%")
            
            # 步骤1: 创建所需目录
            self.update_progress(5, "检查并创建所需目录...", "5%")
            self.create_required_directories()
            
            # 步骤2: 处理明细表格（使用整合版函数）
            self.update_progress_with_random_speed(10, 30, "正在处理明细表格...")
            success1, message1 = run_detail_processing_integrated()
            
            if not success1:
                self.show_error(f"处理明细表格失败:\n{message1}")
                return
            
            self.update_progress(30, "明细表格处理完成", "30%")
            
            # 倒计时2秒
            self.update_progress_with_random_speed(30, 40, "等待2秒...")
            self.countdown(2)
            
            # 步骤3: 处理订单表格（使用整合版函数）
            self.update_progress_with_random_speed(40, 70, "正在处理订单表格...")
            success2, message2 = run_order_processing_integrated()

            if not success2:
                self.show_error(f"处理订单表格失败:\n{message2}")
                return

            self.result_message = message2  # 保存成功消息（含警告）
            # 提取输出文件路径
            for line in message2.split('\n'):
                if line.startswith('生成文件:'):
                    fname = line.split('生成文件:', 1)[-1].strip()
                    self.output_file_path = os.path.join(r'D:\订单表格', fname)
                    break
            # 商品明细统计（读取原始目录，在删除前）
            if self.show_detail_stats_var.get():
                self.update_progress(72, "正在统计商品明细...", "72%")
                detail_lines, detail_total, w_name, w_spec, w_unit, raw_items = compute_detail_stats(r'D:\明细表格')
                if detail_lines:
                    self.detail_stats_lines = detail_lines
                    self.detail_stats_total = detail_total
                    self.detail_stats_widths = (w_name, w_spec, w_unit)
                    self.detail_stats_raw = raw_items
                else:
                    self.detail_stats_lines = None
            self.update_progress(70, "订单表格处理完成", "70%")
            
            # 倒计时2秒
            self.update_progress_with_random_speed(70, 80, "等待2秒...")
            self.countdown(2)
            
            # 步骤4: 根据勾选状态删除文件
            if self.delete_files_var.get():
                self.update_progress_with_random_speed(80, 95, "正在删除 D:\\明细表格 目录...")
                success = self.delete_directory_contents(r"D:\明细表格")
                if success:
                    self.update_progress(95, "文件删除完成", "95%")
                else:
                    self.update_progress(95, "文件删除失败或目录不存在", "95%")
            else:
                self.update_progress(95, "跳过文件删除", "95%")
            
            # 完成
            self.update_progress(100, "执行完毕", "100%")
            time.sleep(0.5)
            
            # 更新UI显示完成状态
            self.root.after(0, self.show_completion)
            
        except Exception as e:
            error_msg = f"执行过程中发生错误: {str(e)}"
            self.root.after(0, lambda msg=error_msg: self.show_error(msg))
    
    def create_required_directories(self):
        """创建所需的目录结构"""
        for directory in [r'D:\订单表格', r'D:\明细表格', r'D:\明细表格\MXBG2']:
            if not os.path.exists(directory):
                try:
                    os.makedirs(directory, exist_ok=True)
                except Exception as e:
                    print(f"创建目录失败 {directory}: {e}")
    
    def update_progress(self, value, text, percentage):
        """更新进度条"""
        self.root.after(0, lambda: self._update_progress_ui(value, text, percentage))
    
    def _update_progress_ui(self, value, text, percentage):
        """更新进度条UI"""
        self.progress_bar['value'] = value
        self.progress_label.config(text=text)
        self.percentage_label.config(text=percentage)
        self.status_label.config(text=text)
        self.root.update()
    
    def update_progress_with_random_speed(self, start_value, end_value, text):
        """以不均匀的速度更新进度条"""
        current = start_value
        steps = random.randint(5, 15)  # 随机步数
        step_size = (end_value - start_value) / steps
        
        for i in range(steps):
            current += step_size
            percentage = f"{int(current)}%"
            self.update_progress(current, text, percentage)
            
            # 随机等待时间，模拟不均匀速度
            wait_time = random.uniform(0.1, 0.5)
            time.sleep(wait_time)
    
    def countdown(self, seconds):
        """倒计时"""
        for i in range(seconds, 0, -1):
            self.update_progress(
                self.progress_bar['value'],
                f"等待{i}秒...",
                f"{int(self.progress_bar['value'])}%"
            )
            time.sleep(1)
    
    def show_completion(self):
        """显示完成状态"""
        sys.stdout = getattr(self, '_old_stdout', sys.stdout)
        self.root.update()  # 处理挂起的 GUI 事件
        self.status_label.config(text="执行完毕")
        # 左右排列两个按钮
        self.execute_button.pack_forget()
        self._btn_row = tk.Frame(self.execute_button.master)
        self._btn_row.pack(pady=(0, 20))
        tk.Button(self._btn_row, text="关闭", command=self.close_window,
                  font=("Microsoft YaHei", 14), bg="#e74c3c", fg="white",
                  width=16, height=2).pack(side=tk.LEFT, padx=5)
        if self.output_file_path and os.path.exists(self.output_file_path):
            tk.Button(self._btn_row, text="📄 打开并退出", command=self._open_and_close,
                      font=("Microsoft YaHei", 14), bg="#27ae60", fg="white",
                      width=18, height=2).pack(side=tk.LEFT, padx=5)
        self.is_running = False
        # 左侧：结果文本
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert("1.0", self.result_message)
        # 行标红逻辑
        count = int(self.result_text.index('end-1c').split('.')[0])
        for i in range(1, count + 1):
            line_text = self.result_text.get(f"{i}.0", f"{i}.end")
            if "\x02" in line_text:
                # 清除标记字符，标红
                cleaned = line_text.replace("\x02", "")
                self.result_text.delete(f"{i}.0", f"{i}.end")
                self.result_text.insert(f"{i}.0", cleaned)
                self.result_text.tag_add("red", f"{i}.0", f"{i}.end")
            elif line_text.startswith("⚠"):
                self.result_text.tag_add("red", f"{i}.0", f"{i}.end")
            elif line_text.startswith("──"):
                self.result_text.tag_add("head", f"{i}.0", f"{i}.end")

        # 追加商品明细统计（在 lock 之前）
        if self.detail_stats_lines:
            w_name, w_spec, w_unit = self.detail_stats_widths
            self.result_text.insert(tk.END, f"\n── 商品明细统计（共 {len(self.detail_stats_lines)} 种，出库总计 {self.detail_stats_total}）──\n")
            self.result_text.insert(tk.END, f"  {_pad('商品名称', w_name)}{_pad('规格', w_spec)}{_pad('单位', w_unit)}{'数量':>6}\n")
            sep_w = 2 + w_name + w_spec + w_unit + 6
            self.result_text.insert(tk.END, "  " + "─"*sep_w + "\n")
            # 同商品名同背景色
            colors = ["#f8f8f8", "#eef6ff", "#f5fff0", "#fff8f0"]
            prev_name = None
            color_idx = -1
            for line in self.detail_stats_lines:
                name = line.strip().split()[0] if line.strip() else ""
                if name != prev_name:
                    color_idx = (color_idx + 1) % len(colors)
                    prev_name = name
                tag = f"c{color_idx}"
                self.result_text.tag_config(tag, background=colors[color_idx])
                start = self.result_text.index(tk.END + "-1c")
                self.result_text.insert(tk.END, line + "\n")
                self.result_text.tag_add(tag, start, tk.END + "-1c")

    def _open_and_close(self):
        # 导出结果栏内容为 TXT
        detail_dir = r'D:\订单表格\每日详情'
        os.makedirs(detail_dir, exist_ok=True)
        today = datetime.now().strftime('%m%d')
        txt_path = os.path.join(detail_dir, f'今日详情{today}.txt')
        content = self.result_text.get("1.0", tk.END)
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(content)
        # 打开输出文件
        if self.output_file_path and os.path.exists(self.output_file_path):
            os.startfile(self.output_file_path)
        self.close_window()

    def close_window(self):
        """关闭窗口"""
        self.root.destroy()
    
    def show_error(self, message):
        """显示错误信息"""
        sys.stdout = getattr(self, '_old_stdout', sys.stdout)
        self.is_running = False
        if hasattr(self, '_btn_row'):
            self._btn_row.destroy()
        self.execute_button.pack()
        self.execute_button.config(state=tk.NORMAL, text="执行表格转换", bg="#3498db")
        self.delete_checkbox.config(state=tk.NORMAL)
        self.detail_stats_checkbox.config(state=tk.NORMAL)
        messagebox.showerror("错误", message)
    
    def on_closing(self):
        """窗口关闭事件"""
        if self.is_running:
            if not messagebox.askyesno("确认", "程序正在运行，确定要退出吗？"):
                return
        self.root.destroy()
    
    def delete_directory_contents(self, directory):
        """删除目录下的所有文件和子目录"""
        try:
            if os.path.exists(directory):
                for filename in os.listdir(directory):
                    file_path = os.path.join(directory, filename)
                    try:
                        if os.path.isfile(file_path) or os.path.islink(file_path):
                            os.unlink(file_path)
                        elif os.path.isdir(file_path):
                            shutil.rmtree(file_path)
                    except Exception as e:
                        print(f"删除失败 {file_path}: {e}")
                return True
            return False
        except Exception as e:
            print(f"删除目录失败 {directory}: {e}")
            return False

# ============================================================================
# 主程序入口
# ============================================================================

def main():
    """主函数"""
    root = tk.Tk()
    app = JinhuaJuhuoApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()