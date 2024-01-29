import urllib.request as req
import bs4
import pandas as pd

# 建立一個 Request 物件
url = "https://www.cwa.gov.tw/V8/C/W/Observe/MOD/24hr/46705.html" # 資料測站為新屋氣象站(站點代碼:46705)，可自行更改測站ID
request = req.Request(url)
with req.urlopen(request) as response:
    data = response.read().decode("utf-8")

# 解析原始碼
root = bs4.BeautifulSoup(data, "html.parser")
Wdata = root.find_all("tr")

# 提取需要的欄位放入excel
selected_columns = ["日期時間", "溫度", "天氣", "風向", "風力", "陣風(級)", "能見度(公里)", "相對溼度(%)", "海平面氣壓(百帕)", "當日累積雨量(毫米)", "日照時數"]

# 資料整理
weather_data = []
for row in Wdata[0:]:
    date_time = row.find("th", {"class": "is_show"}).text.strip()
    # weather_icon2word = row.find("td", {"headers": "weather", "class": "is_show"}).find("img")["alt"]     # 將圖標轉為文字敘述，如果遇到氣象站資料缺值則會跳錯誤(2024/1/19)
    values = [date_time] + [value.text.strip() for value in row.find_all("td")]
    # values[2] = weather_icon2word
    row_data = dict(zip(selected_columns, values))
    weather_data.append(row_data)

# 轉換為 DataFrame
df = pd.DataFrame(weather_data)

# 將第一筆(最新)資料的時間做為檔案名稱
last_date_time = df["日期時間"].iloc[0]

# 將 DataFrame 寫到 excel
excel_file_name = f"weather_data_{last_date_time.replace('/', '').replace(' ', '').replace(':', '')}.xlsx"
df.to_excel(excel_file_name, index=False)
