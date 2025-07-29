import openpyxl
from openpyxl import Workbook
import excel
import style
import sys

def main(input_path, output_path):
    """
    Main function to read base workbook, generate new formatted sheet,
    copy data and styles, and save as new Excel file.

    Args:
        input_path (str): 輸入 Excel 檔案路徑
        output_path (str): 輸出 Excel 檔案路徑
    """
    wb_base = openpyxl.load_workbook(input_path)
    base_sheet = wb_base['標準成本結構表']
    weight_sheet = wb_base['鐵板重量計算']
    material_sheet = wb_base['鐵板材料費單價']
    mm_sheet = wb_base['鐵板米數計算']

    new_wb = Workbook()
    output_sheet = new_wb.active
    excel.write_excel_header(output_sheet)
    excel.set_column_widths(output_sheet)
    total_row = style.copy_columns_with_style(
        ws_src=base_sheet,
        ws_dest=output_sheet,
        src_cols='A:H',
        src_start_row=5,
        dest_start_row=3,
        dest_start_col=3
    )
    style.copy_columns_with_style(
        ws_src=base_sheet,
        ws_dest=output_sheet,
        src_cols='I:I',
        src_start_row=5,
        dest_start_row=3,
        dest_start_col=17
    )
    labels, label_nums = excel.get_labels_and_numbers(base_sheet)
    label_name = excel.get_main_name(base_sheet)
    excel.fill_query_no(output_sheet, labels, label_nums)
    excel.set_basic_styles(output_sheet, total_row)

    excel.calculate_and_write_output(base_sheet, output_sheet, weight_sheet, material_sheet, mm_sheet, total_row, 6, 4)
    excel.total_result(output_sheet, total_row)

    output_sheet.title = f'{labels[0]}{label_name} (成本計算)'
    new_wb.save(output_path)

if __name__ == "__main__":
    if len(sys.argv) == 3:
        main(sys.argv[1], sys.argv[2])
    else:
        print("用法: python main.py <輸入檔案路徑> <輸出檔案路徑>")
