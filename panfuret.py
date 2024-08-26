import openpyxl
import streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from copy import copy
import os
import tempfile
from PIL import Image as PILImage  # PillowのImageクラスをインポート
from openpyxl.drawing.image import Image as OpenpyxlImage

# タイトルを追加
st.title('パンフレット作成')

# ファイルパスを指定してExcelファイルを読み込む
@st.cache_data
def load_data(file_path):
    df = pd.read_excel(file_path)
    # '肥料名称' カラムから NaN を取り除く
    df = df.dropna(subset=['肥料名称'])
    return df

#キャッシュクリア
if st.button('Clear Cache'):
    st.cache_data.clear()

df = load_data('銘柄データ_BB.xlsx')
df_ekihi = load_data('銘柄データ_液肥.xlsx')
df_kasei = load_data('銘柄データ_化成.xlsx')

#肥料名称のリストをつくる。
fertilizer_names = df['肥料名称'].tolist()
fertilizer_names_ekihi = df_ekihi['肥料名称'].tolist()
fertilizer_names_kasei = df_kasei['肥料名称'].tolist()

# 選択されたアイテムのリストを作成
selected_fertilizer = []
selected_fertilizer_ekihi = []
selected_fertilizer_kasei = []


st.markdown(
    """
    <style>
    .main .block-container {
        max-width: 1000px;
#       padding: 1rem 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# 3つのカラムを作成
col1, col2, col3 = st.columns(3)

# 1列目にBBのチェックボックスを作成
with col1:
    st.header("BB")
    for fertilizer_name in fertilizer_names:
        if st.checkbox(fertilizer_name, key=fertilizer_name):
            selected_fertilizer.append(fertilizer_name)


# 2列目に化成のチェックボックスを作成
with col2:
    st.header("化成")
    for fertilizer_name_kasei in fertilizer_names_kasei:
        if st.checkbox(fertilizer_name_kasei, key=fertilizer_name_kasei):
            selected_fertilizer_kasei.append(fertilizer_name_kasei)

# 3列目に液肥のチェックボックスを作成
with col3:
    st.header("液肥")
    for fertilizer_name_ekihi in fertilizer_names_ekihi:
        if st.checkbox(fertilizer_name_ekihi, key=fertilizer_name_ekihi):
            selected_fertilizer_ekihi.append(fertilizer_name_ekihi)

# 選択されたアイテムのリストを表示
#st.write('選択された肥料:', selected_fertilizer)
#st.write('選択された球技:', selected_sports)
#st.write('選択された魚:', selected_fish)

# 選択されたアイテムの数を表示
selected_fertilizer_count = len(selected_fertilizer)
selected_fertilizer_count_ekihi = len(selected_fertilizer_ekihi)
selected_fertilizer_count_kasei = len(selected_fertilizer_kasei)



import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Border, PatternFill, Font
from PIL import Image as PILImage
import tempfile
import os

def process_bb_fertilizers():
    if selected_fertilizer_count > 0:
        wb = openpyxl.load_workbook('bb_tem.xlsx')
        ws = wb['BB_テンプレ']
        
        count_number = selected_fertilizer_count
        m = count_number - 1
        count = (m // 2)

        for i in range(count):
            row_count = 1
            col_count = 14
            col_offset = i * 13
            source_range = [[ws.cell(row=r, column=c) for c in range(1, 14)] for r in range(1, 45)]
            dest_start_cell = ws.cell(row=row_count, column=col_count + col_offset)

            def copy_cell(src_cell, dest_cell):
                dest_cell.value = src_cell.value
                if src_cell.has_style:
                    dest_cell.font = copy(src_cell.font)
                    dest_cell.border = copy(src_cell.border)
                    dest_cell.fill = copy(src_cell.fill)
                    dest_cell.number_format = copy(src_cell.number_format)
                    dest_cell.protection = copy(src_cell.protection)
                    dest_cell.alignment = copy(src_cell.alignment)

            for i in range(len(source_range)):
                for j in range(len(source_range[0])):
                    src_cell = source_range[i][j]
                    dest_cell = ws.cell(row=dest_start_cell.row + i, column=dest_start_cell.column + j)
                    copy_cell(src_cell, dest_cell)

            specified_widths = [1, 5.67, 8.42, 5.67, 4, 0.84, 8.08, 8.08, 8.08, 8.08, 8.08, 8.08, 8.08]
            for idx, width in enumerate(specified_widths, start=source_range[0][0].column):
                col_letter = openpyxl.utils.get_column_letter(idx)
                ws.column_dimensions[col_letter].width = width

            for idx, width in enumerate(specified_widths, start=dest_start_cell.column):
                col_letter = openpyxl.utils.get_column_letter(idx)
                ws.column_dimensions[col_letter].width = width

        number = m + 1
        kesu_offset = (number // 2) * 13

        if number % 2 != 0:
            for row in ws.iter_rows(min_row=24, max_row=42, min_col=1 + kesu_offset, max_col=13 + kesu_offset):
                for cell in row:
                    cell.value = None
                    cell.border = Border()
                    cell.fill = PatternFill(fill_type=None)
                    cell.font = Font()

        i = 0
        for fertilizer in selected_fertilizer:
            selected_row = df[df['肥料名称'] == fertilizer]
            row_offset = (i % 2) * 20
            col_offset = (i // 2) * 13
            n_base_row = 5
            n_base_column = 8
            name = ws.cell(row=n_base_row + row_offset - 1, column=n_base_column + col_offset - 7)
            name.value = selected_row['肥料名称'].values[0]
            # Fill other values...
            i = i + 1
        wb.save('bb_tem_finish.xlsx')

def process_kasei_fertilizers():
    if selected_fertilizer_count_kasei > 0:
        wb = openpyxl.load_workbook('kasei_tem.xlsx')
        ws = wb['化成_テンプレ']
        
        count_number_kasei = selected_fertilizer_count_kasei
        m = count_number_kasei - 1
        count_number_kasei = (m // 3)

        for i in range(count_number_kasei):
            row_count = 1
            col_count = 14
            col_offset = i * 13
            source_range = [[ws.cell(row=r, column=c) for c in range(1, 14)] for r in range(1, 51)]
            dest_start_cell = ws.cell(row=row_count, column=col_count + col_offset)

            def copy_cell(src_cell, dest_cell):
                dest_cell.value = src_cell.value
                if src_cell.has_style:
                    dest_cell.font = copy(src_cell.font)
                    dest_cell.border = copy(src_cell.border)
                    dest_cell.fill = copy(src_cell.fill)
                    dest_cell.number_format = copy(src_cell.number_format)
                    dest_cell.protection = copy(src_cell.protection)
                    dest_cell.alignment = copy(src_cell.alignment)

            for i in range(len(source_range)):
                for j in range(len(source_range[0])):
                    src_cell = source_range[i][j]
                    dest_cell = ws.cell(row=dest_start_cell.row + i, column=dest_start_cell.column + j)
                    copy_cell(src_cell, dest_cell)

            specified_widths = [1, 8.42, 5.67, 5.67, 4, 0.84, 8.42, 8.42, 8.42, 8.42, 8.42, 8.42, 8.42]
            for idx, width in enumerate(specified_widths, start=source_range[0][0].column):
                col_letter = openpyxl.utils.get_column_letter(idx)
                ws.column_dimensions[col_letter].width = width

            for idx, width in enumerate(specified_widths, start=dest_start_cell.column):
                col_letter = openpyxl.utils.get_column_letter(idx)
                ws.column_dimensions[col_letter].width = width

        number = m + 1
        kesu_offset = (number // 3) * 13

        if (number - 1) % 3 == 0:
            for row in ws.iter_rows(min_row=20, max_row=50, min_col=1 + kesu_offset, max_col=13 + kesu_offset):
                for cell in row:
                    cell.value = None
                    cell.border = Border()
                    cell.fill = PatternFill(fill_type=None)
                    cell.font = Font()

        if (number - 2) % 3 == 0:
            for row in ws.iter_rows(min_row=36, max_row=42, min_col=1 + kesu_offset, max_col=13 + kesu_offset):
                for cell in row:
                    cell.value = None
                    cell.border = Border()
                    cell.fill = PatternFill(fill_type=None)
                    cell.font = Font()

        i = 0
        for fertilizer in selected_fertilizer_kasei:
            selected_row = df_kasei[df_kasei['肥料名称'] == fertilizer]
            row_offset = (i % 3) * 16
            col_offset = (i // 3) * 13
            n_base_row = 4
            n_base_column = 2
            name = ws.cell(row=n_base_row + row_offset, column=n_base_column + col_offset)
            name.value = selected_row['肥料名称'].values[0]
            # Fill other values...
            i = i + 1
        wb.save('kasei_tem_finish.xlsx')

def process_ekihi_fertilizers():
    if selected_fertilizer_count_ekihi > 0:
        wb = openpyxl.load_workbook('ekihi_tem.xlsx')
        ws = wb['液肥_テンプレ']
        
        count_number_ekihi = selected_fertilizer_count_ekihi
        count_number_ekihi = count_number_ekihi - 1

        for i in range(count_number_ekihi):
            row_count = 1
            col_count = 14
            col_offset = i * 13
            source_range = [[ws.cell(row=r, column=c) for c in range(1, 14)] for r in range(1, 37)]
            dest_start_cell = ws.cell(row=row_count, column=col_count + col_offset)

            def copy_cell(src_cell, dest_cell):
                dest_cell.value = src_cell.value
                if src_cell.has_style:
                    dest_cell.font = copy(src_cell.font)
                    dest_cell.border = copy(src_cell.border)
                    dest_cell.fill = copy(src_cell.fill)
                    dest_cell.number_format = copy(src_cell.number_format)
                    dest_cell.protection = copy(src_cell.protection)
                    dest_cell.alignment = copy(src_cell.alignment)

            for i in range(len(source_range)):
                for j in range(len(source_range[0])):
                    src_cell = source_range[i][j]
                    dest_cell = ws.cell(row=dest_start_cell.row + i, column=dest_start_cell.column + j)
                    copy_cell(src_cell, dest_cell)

            specified_widths = [1, 8.42, 5.67, 5.67, 4, 0.84, 8.42, 8.42, 8.42, 8.42, 8.42, 8.42, 8.42]
            for idx, width in enumerate(specified_widths, start=source_range[0][0].column):
                col_letter = openpyxl.utils.get_column_letter(idx)
                ws.column_dimensions[col_letter].width = width

            for idx, width in enumerate(specified_widths, start=dest_start_cell.column):
                col_letter = openpyxl.utils.get_column_letter(idx)
                ws.column_dimensions[col_letter].width = width

        i = 0
        for fertilizer in selected_fertilizer_ekihi:
            selected_row = df_ekihi[df_ekihi['肥料名称'] == fertilizer]
            row_offset = i * 12
            col_offset = 0
            n_base_row = 5
            n_base_column = 2
            name = ws.cell(row=n_base_row + row_offset, column=n_base_column + col_offset)
            name.value = selected_row['肥料名称'].values[0]
            # Fill other values...
            i = i + 1
        wb.save('ekihi_tem_finish.xlsx')

if st.button('Generate Excel Files'):
    process_bb_fertilizers()
    process_kasei_fertilizers()
    process_ekihi_fertilizers()



# 3つのカラムを作成
col4, col5, col6 = st.columns(3)

with col4:
    # Excelファイルを読み込む
    with open('bb_tem_finish.xlsx', 'rb') as file:  # ここでファイルを開きます
        excel_data = file.read()  # インデントされていることを確認
    # ダウンロードボタンの作成
    st.download_button(
        label="Download Excel File＜BB＞",  # ボタンのラベル
        data=excel_data,  # ダウンロードするデータ
        file_name='bb_tem_finish.xlsx',  # ダウンロード時のファイル名
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'  # MIMEタイプを指定
    )

# 2列目に球技のチェックボックスを作成
with col5:
#    st.header("化成")
    with open('kasei_tem_finish.xlsx', 'rb') as file:
        excel_data_ekihi = file.read()

# ダウンロードボタンの作成
    st.download_button(
        label="Download Excel File＜化成＞",  # ボタンのラベル
        data=excel_data_ekihi,  # ダウンロードするデータ
        file_name='kasei_tem_finish.xlsx',  # ダウンロード時のファイル名
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'  # MIMEタイプを指定
    )

# 3列目に魚のチェックボックスを作成
with col6:
#    st.header("化成")
# ファイルをバイナリモードで開く
    with open('ekihi_tem_finish.xlsx', 'rb') as file:
        excel_data_ekihi = file.read()

# ダウンロードボタンの作成
    st.download_button(
        label="Download Excel File＜液肥＞",  # ボタンのラベル
        data=excel_data_ekihi,  # ダウンロードするデータ
        file_name='ekihi_tem_finish.xlsx',  # ダウンロード時のファイル名
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'  # MIMEタイプを指定
    )

#========

# ファイルをバイナリモードで開く
#with open('bb_tem_finish.xlsx', 'rb') as file:
#    excel_data = file.read()

# ダウンロードボタンの作成
#st.download_button(
#    label="Download Excel File＜化成＞",  # ボタンのラベル
#    data=excel_data,  # ダウンロードするデータ
#    file_name='bb_tem_finish.xlsx',  # ダウンロード時のファイル名
#    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'  # MIMEタイプを指定
#)


# ファイルをバイナリモードで開く
#with open('ekihi_tem_finish.xlsx', 'rb') as file:
#    excel_data_ekihi = file.read()

# ダウンロードボタンの作成
#st.download_button(
#    label="Download Excel File＜液肥＞",  # ボタンのラベル
#    data=excel_data_ekihi,  # ダウンロードするデータ
#    file_name='ekihi_tem_finish.xlsx',  # ダウンロード時のファイル名
#    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'  # MIMEタイプを指定
#)