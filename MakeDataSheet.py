import pandas as pd
import os
import openpyxl


Data_path = r"C:\Users\junsa\Desktop\CBA001\CBA001 Data.xlsx"
Data_sheet_paht = r"C:\Users\junsa\Desktop\CBA001\CBA001 Data Sheet.xlsx"


def make_data_sheet():
    if os.path.isfile(Data_path):
        print('file exist')
    else:
        print("no file ")

    df = pd.read_excel(Data_path, index_col=False)
    target_list = [x for x in df.columns.to_list() if x[3:].isdigit()]
    print(target_list)


    title_list = ["method", "condition", "type", "unit"] + target_list

    wb = openpyxl.Workbook()
    ws = wb.active 
    row_start = 4
    col_start = 2
    # ws.cell(row_start, col_start).value = 

    # make titles
    col_title = col_start
    for title in title_list:
        ws.cell(row_start, col_title).value = title
        if col_title >= col_start + 4:
            col_title += 2
        else:
            col_title += 1


    condition_now = "123"
    row_index = row_start + 1
    for method in ["oil", "レオメータ ", "耐油引張り "]:
        method_count = 0
        # print(method)
        for i, value in enumerate(df["method"].str.contains(method)):
            col_index = col_start
            # print(i, value)
            
            if value:
                row = df.iloc[i, :]
                # print(row["method"], row["condition"], row["type"], row["unit"])
                for title in title_list:

                    if method_count > 0 and title == "method":
                        pass
                    elif row[title] == condition_now and title == "condition":
                        pass
                    else:
                        ws.cell(row_index, col_index).value = row[title]
                    
                    if title == "condition":
                        condition_now = row["condition"]


                    method_count += 1
                    if col_index >= col_start + 4:
                        col_index += 2
                    else:
                        col_index += 1
                row_index += 1

    # print(df["method"].str.contains("レオメータ"))
    wb.save(Data_sheet_paht)

if __name__ == "__main__":
    print("make data sheet")
    make_data_sheet()