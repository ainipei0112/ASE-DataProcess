import os
import json
import time
import datetime
import mysql.connector

# MySQL 連線資訊
db_host = '127.0.0.1'
db_user = 'root'
db_password = ''
db_name = 'wb'

# 資料表名稱
table_name = 'all_2oaoi'

# 資料庫連線設定
mydb = mysql.connector.connect(
    host=db_host,
    user=db_user,
    password=db_password,
    database=db_name
)

# 建立 Cursor 物件
mycursor = mydb.cursor()

# 定義一個函數 process_json，接受 JSON 資料和資料庫連線作為參數
def process_json(data, mydb):
    """處理 JSON 資料並寫入資料庫。

    Args:
        data: JSON 資料字典。
        mydb: 資料庫連線物件。
    """
    # 建立 INSERT 語法
    sql = "INSERT INTO {} (Date, Date_1, Lot, AOI_ID, AOI_Scan_Amount, AOI_Pass_Amount, AOI_Reject_Amount, AOI_Yield, AOI_Yield_Die_Corner, AI_Pass_Amount, AI_Reject_Amount, AI_Yield, AI_Fail_Corner_Yield, Final_Pass_Amount, Final_Reject_Amount, Final_Yield, AI_EA_Overkill_Die_Corner, AI_EA_Overkill_Die_Surface, AI_Image_Overkill_Die_Corner, AI_Image_Overkill_Die_Surface, EA_over_kill_Die_Corner, EA_over_kill_Die_Surface, Image_Overkill_Die_Corner, Image_Overkill_Die_Surface, Total_Images, Image_Overkill, AI_Fail_EA_Die_Corner, AI_Fail_EA_Die_Surface, AI_Fail_Image_Die_Corner, AI_Fail_Image_Die_Surface, AI_Fail_Total, Total_AOI_Die_Corner_Image, AI_Pass, AI_Reduction_Die_Corner, AI_Reduction_All, True_Fail, True_Fail_Crack, True_Fail_Chipout, True_Fail_Die_Surface, True_Fail_Others, EA_True_Fail_Crack, EA_True_Fail_Chipout, EA_True_Fail_Die_Surface, EA_True_Fail_Others, `EA_True_Fail_Crack_Chipout`, Device_ID, OP_EA_Die_Corner, OP_EA_Die_Surface, OP_EA_Others, Die_Overkill) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)".format(table_name)

    # 嘗試將資料插入資料庫
    try:
        mycursor.execute(sql, (
            data["Date"], data["Date_1"], data["Lot"], data["AOI_ID"], data["AOI_Scan_Amount"],
            data["AOI_Pass_Amount"], data["AOI_Reject_Amount"], data["AOI_Yield"], data["AOI_Yield_Die_Corner"],
            data["AI_Pass_Amount"], data["AI_Reject_Amount"], data["AI_Yield"], data["AI_Fail_Corner_Yield"],
            data["Final_Pass_Amount"], data["Final_Reject_Amount"], data["Final_Yield"],
            data["AI_EA_Overkill_Die_Corner"], data["AI_EA_Overkill_Die_Surface"], data["AI_Image_Overkill_Die_Corner"],
            data["AI_Image_Overkill_Die_Surface"], data["EA_over_kill_Die_Corner"], data["EA_over_kill_Die_Surface"],
            data["Image_Overkill_Die_Corner"], data["Image_Overkill_Die_Surface"], data["Total_Images"],
            data["Image_Overkill"], data["AI_Fail_EA_Die_Corner"], data["AI_Fail_EA_Die_Surface"],
            data["AI_Fail_Image_Die_Corner"], data["AI_Fail_Image_Die_Surface"], data["AI_Fail_Total"],
            data["Total_AOI_Die_Corner_Image"], data["AI_Pass"], data["AI_Reduction_Die_Corner"],
            data["AI_Reduction_All"], data["True_Fail"], data["True_Fail_Crack"], data["True_Fail_Chipout"],
            data["True_Fail_Die_Surface"], data["True_Fail_Others"], data["EA_True_Fail_Crack"],
            data["EA_True_Fail_Chipout"], data["EA_True_Fail_Die_Surface"], data["EA_True_Fail_Others"],
            data["EA_True_Fail_Crack_Chipout"], data["Device_ID"], data["OP_EA_Die_Corner"],
            data["OP_EA_Die_Surface"], data["OP_EA_Others"], data["Die_Overkill"]
        ))
        mydb.commit()
        print(f"資料已寫入資料庫: {data['Date']}, {data['Lot']}, {data['AOI_ID']}")
    except mysql.connector.Error as error:
        print(f"資料寫入失敗: {error}")

# 主程式
if __name__ == '__main__':
    # 設置資料讀取路徑
    settings_path = r"\\khwbpeaiaoi01\2451AOI$\WaferMapTemp\AI_Result\settings\settings.json"
    main_path = r'\\khwbpeaiaoi01\2451AOI$\WaferMapTemp\AI_Result'

    try:
        print("Reading database.")
        with open(settings_path, "r", encoding='utf-8') as r_file:
            databases = json.load(r_file)
        database = [data for data in databases["folder_details"]]
    except:
        print("Failed to read database.")
        exit()

    # 獲取昨天和今天的日期
    now = datetime.datetime.now()
    yesterday = (now + datetime.timedelta(-1)).strftime('%Y-%m-%d')
    today = now.strftime('%Y-%m-%d')

    # 遍歷所有資料夾
    for directory in [f.path for f in os.scandir(main_path) if os.path.isdir(f.path)]:
        # 遍歷所有批號資料夾
        for lot_name in [f.path for f in os.scandir(directory) if os.path.isdir(f.path)]:
            print(lot_name)
            # 遍歷所有 JSON 文件
            for Json_file in [f.path for f in os.scandir(lot_name) if os.path.isfile(f.path)]:
                try:
                    with open(Json_file) as json_file:
                        data = json.load(json_file)
                except:
                    pass

                # 檢查 JSON 文件是否有效
                if "OP_Checked" in data and data["AOI_Scan_Amount"] != 0:
                    # 進行邏輯運算
                    incorrect_mag_counter = 0
                    chipout_counter = 0
                    crack_counter = 0
                    others_counter = 0
                    op_incorrect_mag_counter = 0
                    op_chipout_counter = 0
                    op_crack_counter = 0
                    op_others_counter  = 0
                    corner_duplicate_Checker = {}
                    die_duplicate_Checker = {}
                    duplicate_Checker = {}
                    xdie=[]

                    # 獲取 OP_PassDetails、OP_FailDetails、AI_FailDetails 中的 id
                    for file_data in data["OP_PassDetails"]:
                        xdie.append(file_data["id"])
                    for file_data in data["OP_FailDetails"]:
                        xdie.append(file_data["id"])
                    for file_data in data["AI_FailDetails"]:
                        if not file_data["id"] in xdie:
                            continue
                        
                        # 根據不同的 aoidefecttype 進行計數
                        if(file_data["aoiDefectType"] == "Incorrect_Magnification" or file_data["aoiDefectType"] == "Incorrect_Size" or file_data["aoiDefectType"] =="Scratch" or file_data["aoiDefectType"] =="Passivation_Effect" or file_data["aoiDefectType"] =="OP_Ink"):
                            incorrect_mag_counter += 1

                            for tempdata in data["OP_PassDetails"]:
                                if(file_data["id"] == tempdata["id"]):
                                    op_incorrect_mag_counter +=1
                                    break
                        elif(file_data["aoiDefectType"] == "chipout" or file_data["aoiDefectType"] == "Chipout" or file_data["aoiDefectType"] == "Peeling"):
                            chipout_counter +=1

                            for tempdata in data["OP_PassDetails"]:
                                if(file_data["id"] == tempdata["id"]):
                                    op_chipout_counter +=1
                                    break
                        elif(str(file_data["aoiDefectType"]) == "Crack"):
                            crack_counter +=1

                            for tempdata in data["OP_PassDetails"]:
                                if(file_data["id"] == tempdata["id"]):
                                    op_crack_counter +=1
                                    break
                        else:
                            others_counter +=1

                            for tempdata in data["OP_PassDetails"]:
                                if(file_data["id"] == tempdata["id"]):
                                    op_others_counter +=1
                                    break
                    
                        # 獲取 XY 座標
                        XY = file_data["fileName"].split("_")[5:9]
                        XY = "_".join(XY)

                        # 根據不同的 aoidefecttype 更新 die_duplicate_Checker 或 corner_duplicate_Checker
                        if(file_data["aoiDefectType"] == "Incorrect_Magnification" or file_data["aoiDefectType"] == "Incorrect_Size" or file_data["aoiDefectType"] =="Scratch" or file_data["aoiDefectType"] =="Passivation_Effect" or file_data["aoiDefectType"] =="OP_Ink"):
                            if(XY in die_duplicate_Checker):
                                die_duplicate_Checker[XY] +=1
                            else:
                                die_duplicate_Checker[XY] = 1
                        else:
                            if(XY in corner_duplicate_Checker):
                                corner_duplicate_Checker[XY] +=1
                            else:
                                corner_duplicate_Checker[XY] = 1
                    
                    EA_Fail_die = len(die_duplicate_Checker)
                    EA_Fail_corner = len(corner_duplicate_Checker)
                    OP_ChipOut={}
                    OP_Metal_Scratch={}
                    OP_Others={}

                    for file_data in data["OP_FailDetails"]:
                        XY = file_data["fileName"].split("_")[5:9]
                        XY = "_".join(XY)
                        try:
                            if XY in die_duplicate_Checker:
                                del die_duplicate_Checker[XY]
                            if XY in corner_duplicate_Checker:
                                del corner_duplicate_Checker[XY]
                            if file_data["opRejudgeDefectMode"]=="ChipOut":
                                if(XY in OP_ChipOut):
                                    OP_ChipOut[XY] +=1
                                else:
                                    OP_ChipOut[XY] = 1
                            elif file_data["opRejudgeDefectMode"]=="Metal Scratch":
                                if(XY in OP_Metal_Scratch):
                                    OP_Metal_Scratch[XY] +=1
                                else:
                                    OP_Metal_Scratch[XY] = 1
                            else:
                                if(XY in OP_Others):
                                    OP_Others[XY] +=1
                                else:
                                    OP_Others[XY] = 1
                        except:
                            pass
                    
                    print(Json_file)
                    OG_loc = Json_file.split("\\")
                    date_obj = datetime.datetime.strptime(OG_loc[6], "%Y-%m-%d")
                    date_str = date_obj.strftime("%m%d%Y")
                    OG_Loc='\\\\'+OG_loc[2]+'\\'+OG_loc[3]+'\\'+OG_loc[4]+'\\Image'+'\\'+date_str+'\\'+OG_loc[7]+'\\'+OG_loc[8].replace('.json','')

                    # 輸出成功顯示OK
                    if not os.path.exists(OG_Loc):
                        print('DIE',OG_Loc)
                        break
                    else:
                        print('OK',OG_Loc)
                    
                    for file_data in os.scandir(OG_Loc):
                        if(file_data.name.endswith("jpg")):
                            XY = file_data.name.split("_")[5:9]
                            XY = "_".join(XY)
                            if(XY in duplicate_Checker):
                                duplicate_Checker[XY] +=1
                            else:
                                duplicate_Checker[XY] = 1
                    
                    AOI_Fail = len(duplicate_Checker)
                    ai_total_image = (int(data["AI_Pass"]) + int(data["AI_Fail"]))
                    total_overkill = op_crack_counter + op_chipout_counter + op_incorrect_mag_counter + op_others_counter
                    total_ai_fail = incorrect_mag_counter + chipout_counter + crack_counter + others_counter

                    if(len(data["AI_PassDetails"]) == 0):
                        ai_reduction_percent = 1
                        ai_reduction_Sawpercent = 1
                    else:
                        ai_reduction_percent = float(len(data["AI_PassDetails"]))/float(ai_total_image)
                        ai_reduction_Sawpercent = float(data["AI_Pass"])/float(data["AI_Pass"] + crack_counter + chipout_counter + others_counter)

                    if(int(data["Reject_Amount"]) == 0):
                        AI_Yield = 1
                    else:
                        AI_Yield = 1-(float(data["Reject_Amount"]) / float(data["AOI_Scan_Amount"]))

                    if total_ai_fail-incorrect_mag_counter == 0:
                        Ai_Overkill_Saw = 0
                    else:
                        Ai_Overkill_Saw = float(abs(total_overkill - op_incorrect_mag_counter)) / float(ai_total_image)

                    if incorrect_mag_counter == 0:
                        incorrect_mag_overkill_per = 0
                    else:   
                        incorrect_mag_overkill_per = float(op_incorrect_mag_counter)/float(ai_total_image)
                    
                    data_filename = Json_file.split("\\")[-1]
                    data_filename = data_filename.split(".")[:-1]
                    data_filename = ".".join(data_filename)
                    
                    print(data_filename)
                    bre = False
                    for database_data in database:
                        if not (data_filename in database_data["filename"]):
                            pass
                        else:
                            print("IN!")
                            bre = True
                            break
                    
                    if not (bre):
                        continue
                    
                    for database_data in database:
                        if(database_data["filename"] == data_filename):
                            data_date = database_data["Date"]
                            data_device_ID = database_data["Device_ID"]
                            break

                    EA_OP_Crack = 0
                    EA_OP_Chipout = 0
                    EA_OP_DieSurface = 0
                    EA_OP_Others = 0
                    EA_OP_Duplicate_Checker = {}
                    True_Fail_Crack = 0
                    True_Fail_Chipout = 0
                    True_Fail_Die_Surface = 0
                    True_Fail_Others = 0

                    for file_data in data["AI_FailDetails"]:
                        fail = False
                        XY = file_data["fileName"].split("_")[5:9]
                        XY = "_".join(XY)

                        for OP_file_data in data["OP_FailDetails"]:
                            OP_XY = OP_file_data["fileName"].split("_")[5:9]
                            OP_XY = "_".join(OP_XY)
                            if XY == OP_XY:
                                fail = True
                                break
                        
                        if fail:
                            if XY in EA_OP_Duplicate_Checker:
                                if file_data["aoiDefectType"] in ("Incorrect_Magnification", "Incorrect_Size", "Scratch", "Passivation_Effect", "OP_Ink"):
                                    True_Fail_Die_Surface += 1
                                elif file_data["aoiDefectType"] in ("chipout", "Chipout", "Peeling"):
                                    True_Fail_Chipout += 1
                                elif file_data["aoiDefectType"] == "Crack":
                                    True_Fail_Crack += 1
                                else:
                                    True_Fail_Others += 1
                            else:
                                EA_OP_Duplicate_Checker[XY] = 1

                                if file_data["aoiDefectType"] in ("Incorrect_Magnification", "Incorrect_Size", "Scratch", "Passivation_Effect", "OP_Ink"):
                                    EA_OP_DieSurface += 1
                                    True_Fail_Die_Surface += 1
                                elif file_data["aoiDefectType"] in ("chipout", "Chipout", "Peeling"):
                                    EA_OP_Chipout += 1
                                    True_Fail_Chipout += 1
                                elif file_data["aoiDefectType"] == "Crack":
                                    True_Fail_Crack += 1
                                    EA_OP_Crack += 1
                                else:
                                    EA_OP_Others += 1
                                    True_Fail_Others += 1

                    if int(data["AOI_Scan_Amount"]) == 0:
                        continue

                    date_time = datetime.datetime.strptime(data_date, "%Y-%m-%d %H:%M:%S").time()
                    compare_time = datetime.time(7, 30)

                    if date_time>=compare_time:
                        date_day=datetime.datetime.strptime(data_date, "%Y-%m-%d %H:%M:%S").date()
                    else:
                        date_day=(datetime.datetime.strptime(data_date, "%Y-%m-%d %H:%M:%S")-datetime.timedelta(1)).date()
                    
                    data_dictionary = {
                        "Date" : data_date, 
                        "Date_1" : date_day,
                        "Lot" : lot_name.split("\\")[-1], 
                        "AOI_ID" : Json_file.split(".")[-2], 
                        "AOI_Scan_Amount" : data["AOI_Scan_Amount"],
                        "AOI_Pass_Amount" : data["AOI_Scan_Amount"]-AOI_Fail,
                        "AOI_Reject_Amount" : AOI_Fail, 
                        "AOI_Yield" :   float(data["AOI_Scan_Amount"]-AOI_Fail)/float(data["AOI_Scan_Amount"]),
                        "AOI_Yield_Die_Corner" :   float(data["AOI_Scan_Amount"]-(AOI_Fail-EA_Fail_die))/float(data["AOI_Scan_Amount"]),
                        "AI_Pass_Amount" : int(data["AOI_Scan_Amount"]) - (EA_Fail_corner + EA_Fail_die),
                        "AI_Reject_Amount" : EA_Fail_corner + EA_Fail_die,
                        "AI_Yield" :float(int(data["AOI_Scan_Amount"]) - (EA_Fail_corner + EA_Fail_die)) / float(data["AOI_Scan_Amount"]),
                        "AI_Fail_Corner_Yield" :float(data["AOI_Scan_Amount"] - EA_Fail_corner) / float(data["AOI_Scan_Amount"]),
                        "Final_Pass_Amount" : data["Pass_Amount"],
                        "Final_Reject_Amount" : data["Reject_Amount"], 
                        "Final_Yield" : AI_Yield,
                        "AI_EA_Overkill_Die_Corner" : (float(len(corner_duplicate_Checker))/float(data["AOI_Scan_Amount"])) if len(corner_duplicate_Checker) != 0 else 0,
                        "AI_EA_Overkill_Die_Surface" : (float(len(die_duplicate_Checker))/float(data["AOI_Scan_Amount"]))if len(die_duplicate_Checker) != 0 else 0,
                        "AI_Image_Overkill_Die_Corner" : Ai_Overkill_Saw,
                        "AI_Image_Overkill_Die_Surface" : incorrect_mag_overkill_per,
                        "EA_over_kill_Die_Corner" : len(corner_duplicate_Checker),
                        "EA_over_kill_Die_Surface" : len(die_duplicate_Checker),
                        "Image_Overkill_Die_Corner" : abs(total_overkill - op_incorrect_mag_counter),
                        "Image_Overkill_Die_Surface" : op_incorrect_mag_counter,
                        "Total_Images" : ai_total_image,
                        "Image_Overkill" : abs(total_overkill - op_incorrect_mag_counter)+op_incorrect_mag_counter,
                        "AI_Fail_EA_Die_Corner" : EA_Fail_corner,
                        "AI_Fail_EA_Die_Surface" : EA_Fail_die, 
                        "AI_Fail_Image_Die_Corner" : crack_counter + chipout_counter + others_counter,
                        "AI_Fail_Image_Die_Surface" : incorrect_mag_counter,
                        "AI_Fail_Total" : total_ai_fail,
                        "Total_AOI_Die_Corner_Image" : ai_total_image-incorrect_mag_counter,
                        "AI_Pass" : data["AI_Pass"], 
                        "AI_Reduction_Die_Corner" : ai_reduction_Sawpercent, 
                        "AI_Reduction_All" : ai_reduction_percent, 
                        "True_Fail" : len(data["OP_FailDetails"]),
                        "True_Fail_Crack" : True_Fail_Crack,
                        "True_Fail_Chipout" :True_Fail_Chipout,
                        "True_Fail_Die_Surface" : True_Fail_Die_Surface,
                        "True_Fail_Others" : True_Fail_Others,
                        "EA_True_Fail_Crack" : EA_OP_Crack ,
                        "EA_True_Fail_Chipout" : EA_OP_Chipout,
                        "EA_True_Fail_Die_Surface" : EA_OP_DieSurface,
                        "EA_True_Fail_Others": EA_OP_Others, 
                        "EA_True_Fail_Crack_Chipout" : EA_OP_Crack+EA_OP_Chipout,
                        "Device_ID" : data_device_ID,
                        "OP_EA_Die_Corner" : len(OP_ChipOut),
                        "OP_EA_Die_Surface" : len(OP_Metal_Scratch),
                        "OP_EA_Others": len(OP_Others),
                        "Die_Overkill": EA_Fail_corner + EA_Fail_die - data["Reject_Amount"]
                    }
                    # 使用 process_json 函數將資料寫入資料庫
                    process_json(data_dictionary, mydb)

    # 關閉 Cursor 和連線
    mycursor.close()
    mydb.close()