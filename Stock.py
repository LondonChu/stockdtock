import pandas as pd
from datetime import datetime
import requests

# 讀取 HTML 檔案
file_path_has = "C:\\Users\\kyz87\\Downloads\\未實現損益.xls"
dfs_has = pd.read_html(file_path_has, header=0, skiprows=1)
df_has = dfs_has[0]
codes_has = df_has.iloc[:, 2].dropna().astype(str).tolist()
codes_has = codes_has[:-1]
result_has = ", ".join(codes_has)
print(f"庫存未實現 : {len(codes_has)} 筆 : \r\n{result_has}")

# 讀取 .xlsx 檔案
file_path_buy = f"C:\\Users\\kyz87\\Downloads\\股東會匯出_{datetime.now().strftime('%Y%m%d')}.xlsx"
dfs_buy = pd.read_excel(file_path_buy, engine='openpyxl')
codes_buy = dfs_buy['股票代號'].dropna().astype(str).tolist()
# codes_buy = dfs_buy['股票名稱'].dropna().astype(str).tolist()
result_buy = ", ".join(codes_buy)
print(f"股東會匯出 : {len(codes_buy)} 筆 : \r\n{result_buy}")

# 比較 result_has 與 result_buy
set_has = set(codes_has)
set_buy = set(codes_buy)
diff = set_buy - set_has

# 列出在 result_buy 中但沒有在 result_has 中的所有資料
result_diff = ", ".join(diff)
print(f"在股東會匯出中但不在庫存未實現中的股票代號: {len(diff)} 筆 : \r\n{result_diff}")

# 傳送結果到 Line
line_channel_access_token = '11bb326bc269c1f9614e1a7cc2a8e73c'
line_user_id = '2007068661'
line_message_api = 'https://api.line.me/v2/bot/message/push'
message = f"在股東會匯出中但不在庫存未實現中的股票代號: {len(diff)} 筆 : \r\n{result_diff}"

headers = {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + line_channel_access_token
}

data = {
    'to': line_user_id,
    'messages': [{
        'type': 'text',
        'text': message
    }]
}

response = requests.post(line_message_api, headers=headers, json=data)
print(f"Line Message API Response: {response.status_code}")

# Channel資訊
# Channel ID
# 2007068661
# Channel secret
# 11bb326bc269c1f9614e1a7cc2a8e73c

