from openpyxl import load_workbook
wb = load_workbook("sample.xlsx")
ws = wb.active

# row 추가 
ws.insert_rows(8)   # 8행이 비워짐
ws.insert_rows(8,5) # 8번째 줄 위치에 5줄 추가 

ws.insert_cols(2) # 새로운 빈 열 추가 
ws.insert_cols(2,3)  # B 번째 열부터 3열 추가 

 
wb.save("sample_insert_rows.xlsx") 