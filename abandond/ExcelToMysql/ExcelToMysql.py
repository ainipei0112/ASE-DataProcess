import shutil
import pandas as pd
import mysql.connector
import os
import re
import datetime

# MySQL 連線資訊
db_host = '10.11.33.122'
db_user = 'root'
db_password = ''
db_name = 'wb'

# 資料表名稱
table_name = 'all_2oaoi'

# CSV 檔案所在資料夾路徑
csv_folder = "D:\ASEKH\K18330\DataProcess"
# 目標資料夾路徑
target_folder = "D:\ASEKH\K18330\DataProcess\All Data"

# 建立 MySQL 連線
mydb = mysql.connector.connect(
    host=db_host,
    user=db_user,
    password=db_password,
    database=db_name
)

# 驗證連線
if mydb.is_connected():
    print("Database connection successful")
else:
    print("Database connection failed")

# 建立 Cursor 物件
mycursor = mydb.cursor()

# 建立 INSERT 語法
sql = "INSERT INTO {} (Date, Date_1, Lot, AOI_ID, AOI_Scan_Amount, AOI_Pass_Amount, AOI_Reject_Amount, AOI_Yield, AOI_Yield_Die_Corner, AI_Pass_Amount, AI_Reject_Amount, AI_Yield, AI_Fail_Corner_Yield, Final_Pass_Amount, Final_Reject_Amount, Final_Yield, AI_EA_Overkill_Die_Corner, AI_EA_Overkill_Die_Surface, AI_Image_Overkill_Die_Corner, AI_Image_Overkill_Die_Surface, EA_over_kill_Die_Corner, EA_over_kill_Die_Surface, Image_Overkill_Die_Corner, Image_Overkill_Die_Surface, Total_Images, Image_Overkill, AI_Fail_EA_Die_Corner, AI_Fail_EA_Die_Surface, AI_Fail_Image_Die_Corner, AI_Fail_Image_Die_Surface, AI_Fail_Total, Total_AOI_Die_Corner_Image, AI_Pass, AI_Reduction_Die_Corner, AI_Reduction_All, True_Fail, True_Fail_Crack, True_Fail_Chipout, True_Fail_Die_Surface, True_Fail_Others, EA_True_Fail_Crack, EA_True_Fail_Chipout, EA_True_Fail_Die_Surface, EA_True_Fail_Others, `EA_True_Fail_Crack_Chipout`, Device_ID, OP_EA_Die_Corner, OP_EA_Die_Surface, OP_EA_Others, Die_Overkill) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)".format(table_name)

# 初始化總插入筆數
total_rows_inserted = 0

# 遍歷資料夾中的檔案
for filename in os.listdir(csv_folder):
    # 使用正則表達式檢查檔案名稱是否符合命名規則
    if re.match(r'^\d{2}\d{2}_All_\(Security C\).csv$', filename):
        # 構建完整檔案路徑
        csv_file = os.path.join(csv_folder, filename)

        # 讀取 CSV 檔案
        df = pd.read_csv(csv_file)

        # 將 DataFrame 轉換成資料列表
        data = df.values.tolist()

        # 執行 INSERT 語法
        try:
            mycursor.executemany(sql, data)
            mydb.commit()

            # 取得插入的資料筆數
            rows_inserted = mycursor.rowcount
            total_rows_inserted += rows_inserted
            print(f"File: {filename} IN {rows_inserted}!")

            # 取得檔案月份
            month = filename[:2]

            # 取得當年的年份
            current_year = datetime.datetime.now().year

            # 構建目標資料夾路徑
            month_folder = os.path.join(target_folder, str(current_year), month)

            # 檢查資料夾是否存在，不存在則建立資料夾
            if not os.path.exists(month_folder):
                os.makedirs(month_folder)

            # 移動檔案
            shutil.move(csv_file, month_folder)
        except mysql.connector.Error as error:
            print(f"File: {filename} ERROR: {error}")

# 輸出總共插入的資料筆數
print(f"\nTotal rows inserted: {total_rows_inserted}")

# 關閉 Cursor 和連線
mycursor.close()
mydb.close()