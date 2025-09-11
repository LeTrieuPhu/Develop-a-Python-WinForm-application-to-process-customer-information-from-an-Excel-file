#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_summary_format.py
Tạo file summary và tự động format (wrap text, auto column width, row height)
Usage:
  python extract_summary_format.py input.xlsx output.xlsx
"""
import os
import re, sys
import pandas as pd
from pathlib import Path
from datetime import datetime
import win32com.client as win32
from openpyxl.styles import Font
from send2trash import send2trash
from openpyxl import load_workbook

# cần openpyxl để định dạng file excel đầu ra
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

shipper_keys = ['shipper']
weight_keys = ['total weight (kg)']
volume_keys = ['total volume (cbm)']
cases_keys = ['cases booked']
opc_keys = ['origin port code']
date_keys = ['date freight available']

# định dạng file output
def auto_format_excel(path, sheet_name=None, max_col_width=80, base_width_factor=1.1, line_height=15):
    """
    Mở file với openpyxl và:
    - bật wrap_text cho tất cả ô
    - set column width theo nội dung (giới hạn max_col_width)
    - set row height dựa trên số dòng trong ô (count '\n' + 1) * line_height
    """
    wb = load_workbook(path)
    ws = wb[sheet_name] if sheet_name else wb.active

    # Tính độ rộng tối đa có thể cần cho mỗi cột (dựa vào nội dung)
    col_max_chars = {}
    for row in ws.iter_rows(values_only=True):
        for idx, cell in enumerate(row, start=1):
            if cell is None:
                length = 0
            else:
                # lấy chuỗi, tính max length trong các dòng nếu có newline
                s = str(cell)
                lines = s.splitlines()
                length = max((len(line) for line in lines), default=0)
            col_max_chars[idx] = max(col_max_chars.get(idx, 0), length)

    # Set column widths
    for idx, max_chars in col_max_chars.items():
        col_letter = get_column_letter(idx)
        # Thông thường 1 char ~ 1.0 đơn vị width; nhân thêm hệ số để vừa nhìn
        width = max_chars * base_width_factor + 2
        if width > max_col_width:
            width = max_col_width
        if width < 8:  # tối thiểu
            width = 8
        ws.column_dimensions[col_letter].width = width

    # Wrap text và set alignment, set row height dựa trên số dòng lớn nhất trong hàng
    for r_idx, row in enumerate(ws.iter_rows(), start=1):
        max_lines, i = 1, 0
        for cell in row:
            if i != 1 or r_idx == 1:
                cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(wrapText=True, vertical='top')
            # bật wrap text
            if cell.value is not None:                
                # đếm số dòng
                lines = str(cell.value).count('\n') + 1
                if lines > max_lines:
                    max_lines = lines
            i += 1
        
        # set row height (nhân theo số dòng)
        ws.row_dimensions[r_idx].height = max_lines * line_height

    wb.save(path)
    wb.close()

# kiểm tra ô dữ liệu có rỗng không
def is_empty(v):
    return pd.isna(v) or (isinstance(v, str) and v.strip() == '')

def normalize_label(s):
    if s is None:
        return ''
    return re.sub(r'[:\s]+$', '', str(s).strip().lower())

def check_existing_file(output_path, ma_don_hang):

    existing_df = ''
    if os.path.exists(output_path):
        existing_df = pd.read_excel(output_path)
    else:
        return 0
    
    if os.path.getsize(output_path) == 0:
        # File có tồn tại nhưng trống rỗng
        return 1

    # Kiểm tra trùng mã đơn hàng
    for _, row in existing_df.iterrows():
        if ma_don_hang in row.iloc[0]:   # hoặc row["Mã đơn hàng"]
            return 2 
        
    return 3

def convert_xls_to_xlsx(xls_path, xlsx_path):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(str(xls_path))
    wb.SaveAs(str(xlsx_path), FileFormat=51)  # 51 = .xlsx
    wb.Close()
    excel.Application.Quit()

def find_group_value(df_ab, key_variants):
    n = len(df_ab)
    for i in range(n):
        a = df_ab.iat[i, 0]
        a_norm = normalize_label(a)
        matched = False
        for kv in key_variants:
            if a_norm.startswith(kv):
                matched = True
                break
        if matched:
            parts = None
            b = df_ab.iat[i, 1]
            if not is_empty(b):
                parts = str(b).strip()
            else:
                continue
            # j = i + 1
            # while j < n and is_empty(df_ab.iat[j, 0]):
            #     bj = df_ab.iat[j, 1]
            #     if not is_empty(bj):
            #         parts.append(str(bj).strip())
            #     j += 1
            return parts # ', '.join(parts)
    return ''

def safe_float(val):
    try:
        float(val)
        return True
    except (TypeError, ValueError):
        return False

def original_cases(val):
    if not isinstance(val, str):
        return None
    match = re.search(r"Original:\s*([\d\.]+)", val)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            return None
    return None

def total_cases(path, sheet_id, key_variants):
    wb = load_workbook(path, data_only=True)
    ws = wb.worksheets[sheet_id]
    fv = None
    total = 0
    found = False
    key_col = None
    key_row = None

    for r_idx, row in enumerate(ws.iter_rows(values_only=False), start=1):
        for c_idx, cell in enumerate(row, start=1):
            val = cell.value
            if is_empty(val):
                continue
            norm = normalize_label(val).replace('\n', ' ')
            # kiểm tra startswith trên từng variant
            for kv in key_variants:
                if norm.startswith(kv):
                    # tìm thấy key
                    found = True
                    key_col = c_idx
                    key_row = r_idx
                    break
            if found:
                break
        if found:
            break

    if not found:
        wb.close()
        return 0  # không thấy key -> trả về 0

    # Duyệt từ hàng tiếp theo (key_row+1) xuống cuối theo cột key_col
    max_row = ws.max_row
    for r in range(key_row + 1, max_row + 1):
        cell = ws.cell(row=r, column=key_col)
        # kiểm tra rỗng
        if is_empty(cell.value):
            continue
        # kiểm tra strikethrough (gạch ngang)
        strike = False
        try:
            strike = bool(getattr(cell.font, "strike", False))
        except:
            strike = False
        if strike:
            continue
        if safe_float(cell.value):
            # thử chuyển thành float
            fv = float(cell.value)
        else:
            fv = original_cases(cell.value)
        if fv is not None:
            total += fv
        else:
            # nếu không phải số, bỏ qua
            continue

    wb.close()

    return total

def Xu_Ly_File_Input(input_path):
    p = Path(input_path)
    if not p.exists():
        raise FileNotFoundError(f"File không tồn tại: {input_path}")
    input_path_final = input_path
    read_kwargs = {}
    if p.suffix.lower() == '.xls':
        input_path_xlsx = p.with_suffix('.xlsx')
        convert_xls_to_xlsx(input_path, input_path_xlsx)
        input_path_final = input_path_xlsx
        os.remove(input_path)
    # Đọc file không lấy header (để A1 là iat[0,0])
    df_1 = pd.read_excel(input_path_final, dtype=object, header=None, sheet_name=0, **read_kwargs)

    if df_1.shape[1] < 2:
        raise ValueError("File phải có ít nhất 2 cột (A và B).")

    df_ab = df_1.iloc[:, :2].copy().reset_index(drop=True)

    # Mã đơn hàng lấy ô A1:
    ma_don_hang = ''
    if not is_empty(df_ab.iat[0, 0]):
        ma_don_hang = str(df_ab.iat[0, 0]).strip().split(" ")[0]

    return df_ab, ma_don_hang, input_path_final

def Xu_Ly_Data_Input(df_ab):
    shipper_val = find_group_value(df_ab, shipper_keys)
    weight_val = find_group_value(df_ab, weight_keys)
    volume_val = find_group_value(df_ab, volume_keys)
    opc_val = find_group_value(df_ab, opc_keys)
    date_val = find_group_value(df_ab, date_keys)

    # Nếu date_val là string dạng "09/16/2025 16:30"
    if isinstance(date_val, str) and date_val.strip():
        try:
            dt = datetime.strptime(date_val.strip(), "%m/%d/%Y %H:%M")
            date_val = dt.strftime("%d/%m/%Y %H:%M")  # Xuất lại theo DD/MM/YYYY HH:MM
        except ValueError:
            # Nếu không parse được thì giữ nguyên
            pass
    
    return shipper_val, weight_val, volume_val, opc_val, date_val
    
def Loc_Thong_Tin(input_path, output_path):
    # df_2 là input sheet 2, df_ab là sheet 1 có 2 cột
    df_ab, ma_don_hang, input_path_xlsx = Xu_Ly_File_Input(input_path)

    # Kiểm tra Mã đơn hàng có trong file output chưa nếu output đã tồn tại
    # Maloi = 0 | File chưa tồn tại
    # Maloi = 1 | File có tồn tại nhưng trống rỗng
    # Maloi = 2 | Trùng mã đơn hàng
    # Maloi = 3 | File đã tồn tại và sẵn sàng thêm dữ liệu
    Maloi = check_existing_file(output_path, ma_don_hang)
    if Maloi == 2:
        return 0
    
    # Các giá trị cần lấy từ file
    shipper_val, weight_val, volume_val, opc_val, date_val = Xu_Ly_Data_Input(df_ab)
    cases_val = total_cases(input_path_xlsx, 1, cases_keys)

    if opc_val == 'DAD':
        ma_don_hang = 'DAD ' + ma_don_hang

    new_row = pd.DataFrame([{
        'Mã đơn hàng': ma_don_hang,
        'Shipper': shipper_val,
        'Origin Port Code': opc_val,
        'Date Freight Available': date_val,
        'Cases Booked' : float(cases_val),
        'Total Weight (KG)': float(weight_val),
        'Total Volume (CBM)': float(volume_val)
        
    }])

    # Nếu file output đã tồn tại thì đọc rồi nối thêm
    if Maloi == 3:
        existing_df = pd.read_excel(output_path)
        out_df = pd.concat([existing_df, new_row], ignore_index=True)
    else:
        out_df = new_row
        
    # lưu file tạm rồi format bằng openpyxl
    out_df.to_excel(output_path, index=False)

    # định dạng tự động
    auto_format_excel(output_path, sheet_name=None, max_col_width=80, base_width_factor=1.1, line_height=18)

    return True



def Cap_Nhat_Thong_Tin(input_path, output_path, changed_cells):
    # df_2 là input sheet 2, df_ab là sheet 1 có 2 cột
    df_ab, ma_don_hang, input_path_xlsx = Xu_Ly_File_Input(input_path)
    # Các giá trị cần lấy từ file
    shipper_val, weight_val, volume_val, opc_val, date_val = Xu_Ly_Data_Input(df_ab)
    cases_val = total_cases(input_path_xlsx, 1, cases_keys)
    existing_df = pd.read_excel(output_path)

    if opc_val == 'DAD':
        ma_don_hang = 'DAD ' + ma_don_hang

    # Tìm và cập nhật
    flag = False
    for idx, row in existing_df.iterrows():
        if ma_don_hang == row.iloc[0]:
            updates = {
                "Cases Booked": float(cases_val),
                "Total Weight (KG)": float(weight_val),
                "Total Volume (CBM)": float(volume_val)                 
            }
            
            for col, new_val in updates.items():
                if safe_float(row.get(col)):
                    old_val = float(row.get(col))
                if str(old_val).strip() != str(new_val).strip():
                    existing_df.at[idx, col] = new_val
                    changed_cells.append((idx, col))  # lưu lại ô thay đổi   
            flag = True
            break  # dừng sau khi tìm thấy để tránh ghi đè nhiều dòng
    if not flag:
        return False

    # Lưu lại file
    existing_df.to_excel(output_path, index=False)

    # mở lại bằng openpyxl để định dạng ô đổi màu
    wb = load_workbook(output_path)
    ws = wb.active

    # ánh xạ tên cột -> số cột
    col_map = {col: i+1 for i, col in enumerate(existing_df.columns)}

    for idx, col in changed_cells:
        excel_row = idx + 2  # +2 vì DataFrame index bắt đầu từ 0, Excel từ 1 và có header
        excel_col = col_map[col]
        cell = ws.cell(row=excel_row, column=excel_col)
        cell.font = Font(color="FF0000")  # chữ đỏ

    wb.save(output_path)
    wb.close()

    # định dạng tự động (giữ wrap text, width, height...)
    auto_format_excel(output_path, sheet_name=None, max_col_width=80, base_width_factor=1.1, line_height=18)

    return True

def Gop_File(list_path_file, output_path):
    ma_don_hang_total = shipper_val = opc_val = date_val = ''
    weight_val_total = volume_val_total = cases_val_total = 0
    for idx, input_path in enumerate(list_path_file):
        # df_2 là input sheet 2, df_ab là sheet 1 có 2 cột
        df_ab, ma_don_hang, input_path_xlsx = Xu_Ly_File_Input(input_path)

        cases_val = total_cases(input_path_xlsx, 1, cases_keys)
        # Các giá trị cần lấy từ file
        shipper_val, weight_val, volume_val, opc_val, date_val = Xu_Ly_Data_Input(df_ab)
        if idx != 0:
            ma_don_hang_total = ma_don_hang_total + ' + ' + ma_don_hang
        else:
            ma_don_hang_total = ma_don_hang
        cases_val_total += float(cases_val)
        weight_val_total += float(weight_val)
        volume_val_total += float(volume_val)   

    Maloi = check_existing_file(output_path, ma_don_hang_total)
    if Maloi == 2:
        return 0
    
    new_row = pd.DataFrame([{
        'Mã đơn hàng': ma_don_hang_total,
        'Shipper': shipper_val,
        'Origin Port Code': opc_val,
        'Date Freight Available': date_val,
        'Cases Booked' : cases_val_total,
        'Total Weight (KG)': float(weight_val_total),
        'Total Volume (CBM)': float(volume_val_total)
        
    }])

    # Nếu file output đã tồn tại thì đọc rồi nối thêm
    if Maloi == 3:
        existing_df = pd.read_excel(output_path)
        out_df = pd.concat([existing_df, new_row], ignore_index=True)
    else:
        out_df = new_row
        
    # lưu file tạm rồi format bằng openpyxl
    out_df.to_excel(output_path, index=False)

    # định dạng tự động
    auto_format_excel(output_path, sheet_name=None, max_col_width=80, base_width_factor=1.1, line_height=18)

    return True