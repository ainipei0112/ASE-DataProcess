import os
import re
import pandas as pd
import mysql.connector

# 資料庫配置資訊
db_host = '127.0.0.1'
db_user = 'root'
db_password = ''
db_name = 'wb'

# 建立 MySQL 連線
mydb = mysql.connector.connect(
    host=db_host,
    user=db_user,
    password=db_password,
    database=db_name
)

# 建立游標物件
mycursor = mydb.cursor()

# Excel 檔案所在資料夾路徑
excel_folder = r'D:\ASEKH\K18330\資料處理'

# 定義欄位對應關係 Excel:MySQL
field_mapping = {
    'ID': 'ID',
    'Date': 'Date',
    'Date_1': 'Date_1',
    'Lot': 'Lot',
    'AOI_ID': 'AOI_ID',
    'AOI_Scan_Amount': 'AOI_Scan_Amount',
    'AOI_Pass_Amount': 'AOI_Pass_Amount',
    'AOI_Reject_Amount': 'AOI_Reject_Amount',
    'AOI_Yield': 'AOI_Yield',
    'AOI_Yield_Die_Corner': 'AOI_Yield_Die_Corner',
    'AI_Pass_Amount': 'AI_Pass_Amount',
    'AI_Reject_Amount': 'AI_Reject_Amount',
    'AI_Yield': 'AI_Yield',
    'AI_Fail_Corner_Yield': 'AI_Fail_Corner_Yield',
    'Final_Pass_Amount': 'Final_Pass_Amount',
    'Final_Reject_Amount': 'Final_Reject_Amount',
    'Final_Yield': 'Final_Yield',
    'AI_EA_Overkill_Die_Corner': 'AI_EA_Overkill_Die_Corner',
    'AI_EA_Overkill_Die_Surface': 'AI_EA_Overkill_Die_Surface',
    'AI_Image_Overkill_Die_Corner': 'AI_Image_Overkill_Die_Corner',
    'AI_Image_Overkill_Die_Surface': 'AI_Image_Overkill_Die_Surface',
    'EA_over_kill_Die_Corner': 'EA_over_kill_Die_Corner',
    'EA_over_kill_Die_Surface': 'EA_over_kill_Die_Surface',
    'Image_Overkill_Die_Corner': 'Image_Overkill_Die_Corner',
    'Image_Overkill_Die_Surface': 'Image_Overkill_Die_Surface',
    'Total_Images': 'Total_Images',
    'Image_Overkill': 'Image_Overkill',
    'AI_Fail_EA_Die_Corner': 'AI_Fail_EA_Die_Corner',
    'AI_Fail_EA_Die_Surface': 'AI_Fail_EA_Die_Surface',
    'AI_Fail_Image_Die_Corner': 'AI_Fail_Image_Die_Corner',
    'AI_Fail_Image_Die_Surface': 'AI_Fail_Image_Die_Surface',
    'AI_Fail_Total': 'AI_Fail_Total',
    'Total_AOI_Die_Corner_Image': 'Total_AOI_Die_Corner_Image',
    'AI_Pass': 'AI_Pass',
    'AI_Reduction_Die_Corner': 'AI_Reduction_Die_Corner',
    'AI_Reduction_All': 'AI_Reduction_All',
    'True_Fail': 'True_Fail',
    'True_Fail_Crack': 'True_Fail_Crack',
    'True_Fail_Chipout': 'True_Fail_Chipout',
    'True_Fail_Die_Surface': 'True_Fail_Die_Surface',
    'True_Fail_Others': 'True_Fail_Others',
    'EA_True_Fail_Crack': 'EA_True_Fail_Crack',
    'EA_True_Fail_Chipout': 'EA_True_Fail_Chipout',
    'EA_True_Fail_Die_Surface': 'EA_True_Fail_Die_Surface',
    'EA_True_Fail_Others': 'EA_True_Fail_Others',
    'EA_True_Fail_Crack_Chipout': 'EA_True_Fail_Crack_Chipout',
    'Device_ID': 'Device_ID',
    'OP_EA_Die_Corner': 'OP_EA_Die_Corner',
    'OP_EA_Die_Surface': 'OP_EA_Die_Surface',
    'OP_EA_Others': 'OP_EA_Others',
    'Die_Overkill': 'Die_Overkill'
}

# 建立 SQL INSERT 語法
columns = ', '.join(field_mapping.values())
placeholders = ', '.join(['%s' for _ in field_mapping])
sql = "INSERT INTO all_2oaoi ({}) VALUES ({})".format(columns, placeholders)

# 迭代資料夾中的檔案
for filename in os.listdir(excel_folder):
    # 檢查檔案名稱是否符合命名規則
    if re.match(r'^\d{2}\d{2}_All_\(Security C\).xlsx$', filename):
        # 構建完整檔案路徑
        excel_file = os.path.join(excel_folder, filename)

        # 讀取 Excel 資料
        df = pd.read_excel(excel_file)

        # 優化：使用 df.to_records() 轉換為元組列表
        values = [tuple(str(row[column]).replace('&', '&amp;') if column in df.columns and str(row[column]) != 'nan' else None for column in field_mapping) for row in df.to_records(index=False)]

        # 優化：使用executemany一次性插入所有資料
        mycursor.executemany(sql, values)

# 提交交易
mydb.commit()

# 關閉連線
mycursor.close()
mydb.close()

print("資料已成功匯入資料庫！")