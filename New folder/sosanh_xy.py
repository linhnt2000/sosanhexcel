import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

def normalize_value(val):
    if pd.isna(val):
        return ""

    val = str(val).strip()

    # Nếu dạng xxx.0 thì bỏ .0
    if val.endswith(".0"):
        val = val[:-2]

    return val

def compare_excel_x_y(
    file1,
    file2,
    col_x,
    col_y,
    output_file="compare_xy_result.xlsx",
    sheet_name="Result"
):
    
    
    # =============================
    # 1. Đọc dữ liệu
    # =============================
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    for col in [col_x, col_y]:
        if col not in df1.columns or col not in df2.columns:
            raise ValueError(f"❌ Thiếu cột '{col}' trong một trong hai file")

    # Chuẩn hóa dữ liệu
    df1[col_x] = df1[col_x].apply(normalize_value)
    df2[col_x] = df2[col_x].apply(normalize_value)

    df1[col_y] = df1[col_y].apply(normalize_value)
    df2[col_y] = df2[col_y].apply(normalize_value)

    # =============================
    # 2. Map x → SET(y) của file 2
    # =============================
    map_y = (
        df2
        .groupby(col_x)[col_y]
        .apply(set)
        .to_dict()
    )

    # =============================
    # 3. So sánh x – y (ĐÚNG 1–N)
    # =============================
    df1["__xy_status__"] = ""

    for i, row in df1.iterrows():
        x_val = row[col_x]
        y_val = row[col_y]

        if x_val not in map_y:
            df1.loc[i, "__xy_status__"] = "yellow"
        else:
            if y_val in map_y[x_val]:
                df1.loc[i, "__xy_status__"] = "green"
            else:
                df1.loc[i, "__xy_status__"] = "red"


    # =============================
    # 4. Ghi ra Excel
    # =============================
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df1.to_excel(writer, index=False, sheet_name=sheet_name)

    # =============================
    # 5. Tô màu Excel
    # =============================
    wb = load_workbook(output_file)
    ws = wb[sheet_name]

    fills = {
        "green":  PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        "red":    PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
        "yellow": PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid"),
    }

    status_col = ws.max_column

    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center")

    # Tô header
    for c in range(1, ws.max_column + 1):
        ws.cell(row=1, column=c).fill = header_fill
        ws.cell(row=1, column=c).font = header_font
        ws.cell(row=1, column=c).alignment = header_align

    # Tô màu từng dòng
    for r in range(2, ws.max_row + 1):
        color_key = ws.cell(row=r, column=status_col).value
        if color_key in fills:
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).fill = fills[color_key]

    # Ẩn cột trạng thái kỹ thuật
    ws.delete_cols(status_col)

    # Auto width
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(output_file)

    print("✅ Đã xuất file so sánh x–y")
    print(f"📄 File: {output_file}")
    print("🟩 Trùng x & y")
    print("🟥 Trùng x nhưng khác y")
    print("🟨 Không trùng x")

compare_excel_x_y(
    file1="Test2.xlsx",
    file2="Test1.xlsx",
    col_x="MSSV",
    col_y="Trạng thái",
    output_file="doi_soat_xy.xlsx"
)
