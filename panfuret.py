import os
import tempfile
from PIL import Image as PILImage
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Border, PatternFill, Font
from copy import copy
import streamlit as st
import pandas as pd

def reset_cells(ws, min_row, max_row, min_col, max_col):
    """ セルの内容とスタイルをリセットする """
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.value = None
            cell.border = Border()
            cell.fill = PatternFill(fill_type=None)
            cell.font = Font()

def copy_cells(ws, source_range, dest_start_cell):
    """ セルをコピーする """
    def copy_cell(src_cell, dest_cell):
        dest_cell.value = src_cell.value
        if src_cell.has_style:
            dest_cell.font = copy(src_cell.font)
            dest_cell.border = copy(src_cell.border)
            dest_cell.fill = copy(src_cell.fill)
            dest_cell.number_format = copy(src_cell.number_format)
            dest_cell.protection = copy(src_cell.protection)
            dest_cell.alignment = copy(src_cell.alignment)
    
    row_count = len(source_range)
    col_count = len(source_range[0])
    
    for i in range(row_count):
        for j in range(col_count):
            src_cell = source_range[i][j]
            dest_cell = ws.cell(row=dest_start_cell.row + i, column=dest_start_cell.column + j)
            copy_cell(src_cell, dest_cell)

def set_column_widths(ws, column_widths, start_col):
    """ 指定された列幅を設定する """
    for idx, width in enumerate(column_widths, start=start_col):
        col_letter = openpyxl.utils.get_column_letter(idx)
        ws.column_dimensions[col_letter].width = width

def add_image_to_sheet(ws, img_path, cell_address, size):
    """ Excelシートに画像を追加する """
    if os.path.exists(img_path):
        original_img = PILImage.open(img_path)
        resized_img = original_img.resize(size)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_file:
            temp_path = tmp_file.name
            resized_img.save(temp_path)

        img = OpenpyxlImage(temp_path)
        img.anchor = cell_address
        ws.add_image(img)
        os.remove(temp_path)

def process_fertilizers(ws, df, fertilizers, n_base_row, n_base_column, row_offset_step, col_offset_step, image_size, image_size2):
    """ 肥料データを処理してExcelシートに書き込む """
    for i, fertilizer in enumerate(fertilizers):
        selected_row = df[df['肥料名称'] == fertilizer]
        col_offset = i * col_offset_step
        row_offset = (i % 3) * row_offset_step
        
        # 肥料情報の書き込み
        ws.cell(row=n_base_row, column=n_base_column + col_offset).value = selected_row['肥料名称'].values[0]
        
        for k in range(1, 6):
            ws.cell(row=n_base_row + 2, column=n_base_column + col_offset + 6 + k).value = selected_row[f'成分名{k}'].values[0]
            ws.cell(row=n_base_row + 3, column=n_base_column + col_offset + 6 + k).value = selected_row[f'成分{k}'].values[0]
        
        for k in range(1, 3):
            ws.cell(row=n_base_row + 4 + k, column=n_base_column + col_offset + 7).value = selected_row[f'その他{k}'].values[0]
        
        ws.cell(row=n_base_row + 8, column=n_base_column + col_offset + 8).value = selected_row['容量'].values[0]
        ws.cell(row=n_base_row + 9, column=n_base_column + col_offset + 8).value = selected_row['形状'].values[0]
        ws.cell(row=n_base_row + 10, column=n_base_column + col_offset + 8).value = selected_row['液色'].values[0]
        ws.cell(row=n_base_row + 11, column=n_base_column + col_offset + 8).value = selected_row['散布方法'].values[0]
        
        for k in range(1, 7):
            ws.cell(row=n_base_row + 13 + k, column=n_base_column + col_offset).value = selected_row[f'特徴{k}'].values[0]
            ws.cell(row=n_base_row + 20 + k, column=n_base_column + col_offset).value = selected_row[f'使用方法{k}'].values[0]
        
        # 画像の追加
        img_path = os.path.join(script_dir, '容器', f'{fertilizer}.jpg')
        img_path2 = os.path.join(script_dir, '肥効曲線', f'{fertilizer}.jpg')
        
        add_image_to_sheet(ws, img_path, ws.cell(row=n_base_row + 3, column=n_base_column + col_offset).coordinate, image_size)
        add_image_to_sheet(ws, img_path2, ws.cell(row=n_base_row + 10, column=n_base_column + col_offset - 1).coordinate, image_size2)

def main():
    if st.button('開始する＜液肥＞'):
        wb = openpyxl.load_workbook('ekihi_tem.xlsx')
        ws = wb['液肥_テンプレ']
        
        count_number_ekihi = selected_fertilizer_count_ekihi - 1
        for i in range(count_number_ekihi + 1):
            row_count = 1
            col_count = 14
            col_offset = i * 13
            
            source_range = [[ws.cell(row=r, column=c) for c in range(1, 14)] for r in range(1, 37)]
            dest_start_cell = ws.cell(row=row_count, column=col_count + col_offset)
            
            copy_cells(ws, source_range, dest_start_cell)
            
            column_widths = [2.25, 1.5, 6.92, 8.42, 8.42, 8.42, 1, 8.42, 8.42, 8.42, 8.42, 8.42, 8.42]
            set_column_widths(ws, column_widths, dest_start_cell.column)
        
        process_fertilizers(
            ws,
            df_ekihi,
            selected_fertilizer_ekihi,
            n_base_row=4,
            n_base_column=2,
            row_offset_step=16,
            col_offset_step=13,
            image_size=(190, 290),
            image_size2=(440, 170)
        )
        
        wb.save('ekihi_tem_finish.xlsx')
    
    if st.button('開始する＜化成＞'):
        wb = openpyxl.load_workbook('kasei_tem.xlsx')
        ws = wb['化成_テンプレ']
        
        count_number_kasei = selected_fertilizer_count_kasei - 1
        process_fertilizers(
            ws,
            df_kasei,
            selected_fertilizer_kasei,
            n_base_row=4,
            n_base_column=2,
            row_offset_step=16,
            col_offset_step=13,
            image_size=(190, 257),
            image_size2=(440, 170)
        )
        
        wb.save('kasei_tem_finish.xlsx')

    # ダウンロードボタンの作成
    col4, col5, col6 = st.columns(3)
    with col4:
        with open('bb_tem_finish.xlsx', 'rb') as file:
            st.download_button(
                label="Download Excel File＜BB＞",
                data=file.read(),
                file_name='bb_tem_finish.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    
    with col5:
        with open('kasei_tem_finish.xlsx', 'rb') as file:
            st.download_button(
                label="Download Excel File＜化成＞",
                data=file.read(),
                file_name='kasei_tem_finish.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    
    with col6:
        with open('ekihi_tem_finish.xlsx', 'rb') as file:
            st.download_button(
                label="Download Excel File＜液肥＞",
                data=file.read(),
                file_name='ekihi_tem_finish.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

if __name__ == "__main__":
    main()
