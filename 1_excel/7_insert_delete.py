from openpyxl import load_workbook

wb = load_workbook("sample.xlsx")
ws = wb.active

# ws.insert_rows(8)
#ws.insert_rows(8, 5)# 8번부터 5줄 삽입

ws.insert_cols(2)

ws.delete_rows(8)  # 삭제

wb.save("sample_inserted_cols.xlsx")