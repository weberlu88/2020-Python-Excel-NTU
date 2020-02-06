import xlwings as xw
from bs4 import BeautifulSoup
import requests
import time
import numpy as np
import smtplib
from email.mime.text import MIMEText

def send_gmail(msg, sender_mail, receiver_mail, app_pass):
    sender = sender_mail
    receiver = receiver_mail
    content = msg

    msg = MIMEText(content.encode('utf-8'), _charset='utf-8')
    msg['Subject'] = 'Hello 你好'
    msg['From'] = sender
    msg['To'] = receiver

    conn = smtplib.SMTP('smtp.gmail.com:587')
    conn.ehlo()
    conn.starttls()
    conn.login(sender_mail, app_pass)
    conn.sendmail(sender,
                receiver,
                msg.as_string())

    conn.quit()
    return

def yahoo_stock_crawler(stock_id):
    doc = requests.get(f"https://tw.stock.yahoo.com/q/q?s={stock_id}")
    html = BeautifulSoup(doc.text, 'html.parser')
    # 搜尋整個網頁裡，內容為 '個股資料' 的 html 標籤, 關聯到 table 最外層
    table = html.findAll(text="個股資料")[0].parent.parent.parent
    # 找尋 table 裡第二個 tr 標籤內所有的 td 標籤
    data_row = table.select("tr")[1].select("td")

    # 回傳一個字典
    return {
        "open": data_row[8].text,
        "high": data_row[9].text,
        "low": data_row[10].text,
        "close": data_row[2].text,
        "lastClose": data_row[7].text
    }

def run_back_test(tsmc_sheet):
    # 從 B1 儲存格開始，往下查找到最後一個有值的儲存格
    last_cell = tsmc_sheet.range('A1').end('down')
    # 截取該儲存格的 row 值
    last_row = last_cell.row

    # 設定我們的範例試算表上的名稱
    tsmc_sheet.range('K2:K11').name = 'weight10d'
    tsmc_sheet.range('K2:K6').name = 'weight5d'

    # 5日加權移動平均計算
    for i in range(6, last_row+1):
        # 由於我們需要把兩個陣列相乘，因此這是一個 Excel 的陣列運算
        formula = f"=SUM(B{i-4}:B{i}*weight5d)/SUM(weight5d)"
        # 若一個 Excel 的公式使用到陣列運算，需要用 .formula_array 設定
        tsmc_sheet.range(f"C{i}").formula_array = formula

    # 10日加權移動平均計算
    for i in range(11, last_row+1):
        # 由於我們需要把兩個陣列相乘，因此這是一個 Excel 的陣列運算
        formula = f"=SUM(B{i-9}:B{i}*weight10d)/SUM(weight10d)"
        # 若一個 Excel 的公式使用到陣列運算，需要用 .formula_array 設定
        tsmc_sheet.range(f"D{i}").formula_array = formula

    # 計算第一天的交易 2017/10/20
    tsmc_sheet.range("E11").value = 1000
    tsmc_sheet.range("F11").value = 0
    tsmc_sheet.range("G11").value = 1000
    tsmc_sheet.range("H11").value = tsmc_sheet.range("L18").value - tsmc_sheet.range("B11").value * 1000
    tsmc_sheet.range("I11").value = tsmc_sheet.range("H11").value +  tsmc_sheet.range("B11").value * tsmc_sheet.range("G11").value

    # 實作交易策略
    for i in range(12, last_row+1):
        # 截取當天的 5日加權移動平均
        short_term_ma = tsmc_sheet.range(f"C{i}").value
        # 截取當天的 10日加權移動平均
        long_term_ma = tsmc_sheet.range(f"D{i}").value
        # 截取當天收盤價
        price_today = tsmc_sheet.range(f"B{i}").value
        # 若 5日 > 10日，而且我有足夠買入以今日收盤價計價的 1000 股的現金，就買入 1000 股（在 E 欄顯示 1000）
        if (short_term_ma > long_term_ma) and (tsmc_sheet.range(f"H{i-1}").value > price_today * 1000):
            tsmc_sheet.range(f"E{i}").value = 1000
        else:
            # 若上述條件不符和，就買入 0 股，（在 E 欄顯示 0）
            tsmc_sheet.range(f"E{i}").value = 0
        # 若 10日 > 5日，而且昨天的持有股數大於 1000 股，就賣出 1000 股
        if (long_term_ma > short_term_ma) and (tsmc_sheet.range(f"G{i-1}").value >= 1000):
            tsmc_sheet.range(f"F{i}").value = 1000
        else:
            tsmc_sheet.range(f"F{i}").value = 0
        # 持有股數，算法是前一天的持有股數 + 今天的買入股數 - 今天的賣出股數
        tsmc_sheet.range(f"G{i}").value = tsmc_sheet.range(f"G{i-1}").value + tsmc_sheet.range(f"E{i}").value - tsmc_sheet.range(f"F{i}").value
        # 持有資金，算法是前一天的持有資金 + 今日收盤價 x (今天的賣出股數 - 今天的買入股數)
        tsmc_sheet.range(f"H{i}").value = tsmc_sheet.range(f"H{i-1}").value + price_today * (tsmc_sheet.range(f"F{i}").value - tsmc_sheet.range(f"E{i}").value)
        # 總資產則是持有股數 x 今日收盤價 + 今日持有資金
        tsmc_sheet.range(f"I{i}").value = tsmc_sheet.range(f"H{i}").value + tsmc_sheet.range(f"G{i}").value * price_today

    # 計算并且將總收益顯示在 L20
    tsmc_sheet.range("L20").value = tsmc_sheet.range(f"I{last_row}").value - tsmc_sheet.range("L18").value

# 開啓 Excel
wb = xw.Book(r"C:\Users\yuyue\OneDrive\桌面\30hr_pyxl\python_excel_lesson5\stock_portfolio_backtest.xlsx")

# 截取所有記錄在 Portfolio 試算表内的股票代號
port_sheet = wb.sheets["Portfolio"]
last_row = port_sheet.range("A3").end("down").row
stock_data = port_sheet.range(f"A2:A{last_row}").value
stocks = [str(int(stock)) if type(stock) == float else str(stock) for stock in stock_data]

# 產生當日日期
date = time.strftime("%Y/%m/%d")

for stock in stocks:
    data = yahoo_stock_crawler(stock)
    
    try:
        # 嘗試建立一個新的試算表，若該試算表已經存在，就會切換至 except
        sheet = wb.sheets.add(name="TW{}".format(stock), after=wb.sheets[-1])
        # 將範例工作表内的表頭、權重等資料複製到新的工作表
        sheet.range("A1:L20").value = wb.sheets["範例"].range("A1:L20").value
        # 寫入資料
        sheet.range("B2").value = data["close"]
        sheet.range("A2").value = date
    except: 
        # 若該試算表已經存在，就開啓它，並且寫入資料
        sheet = wb.sheets["TW{}".format(stock)]
        sheet.activate()
        last_row = sheet.range("B1").end("down").row
        sheet.range(f"B{last_row+1}").value = data["close"]
        sheet.range(f"A{last_row+1}").value = date

    time.sleep(3)

# 記錄總收益    
balance = 0
# 迭代試算表，針對每一個試算表進行回測，將每一個股票的總收益 (L20儲存格) 加總起來
for sheet in wb.sheets:
    sheet.activate()
    last_row = sheet.range("B1").end("down").row
    print(f"{sheet.name} last row: {last_row}")
    # 工作表名稱不是 "Portfolio" 或 "範例"，而且最後一行不是 1048576，也不能小於 11
    if sheet.name not in ["Portfolio", "範例"] and (last_row != 1048576 and last_row > 11):
        print(f"{sheet.name}")
        run_back_test(sheet)
        balance += sheet.range("L20").value

msg = f"{date} 投資組合收益： ${balance}"
# 發送 Gmail，通知使用者回測結果
send_gmail(msg, 
    port_sheet.range("sender_mail").value, 
    port_sheet.range("receiver_mail").value, 
    port_sheet.range("app_pass").value)
