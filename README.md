★函式說明★
JsonToExcel：處理 khwbpeaiaoi01\2451AOI$\WaferMapTemp\AI_Result 的 JSON 匯出 CSV 或 Excel
CsvToMysql：讀取資料夾內所有 CSV 寫入 MySQL
TransformHistoricalData：將 KHFS2\WBG PE Stage$\AOI 判圖\AOI驗證測試\AI_Result\Excel Results\today\old csv 的歷史 CSV 轉換為可入資料庫的型態

☆參考用法☆
當日資料 -> 由 JsonToExcel 轉換 二光 Json 後再 CsvToMysql 寫入 MySQL
歷史資料 -> 由 TransformHistoricalData 轉換後再 CsvToMysql 寫入 MySQL
