# 引用 xlwings 套件
import xlwings as xw

# 開啓名爲 stock_price_data.xlsx 的檔案
wb = xw.Book(r"stock_price_data.xlsx")
tsmc_sheet = wb.sheets["2330"]

def calc_daily_return(today_price, yesterday_price):
    daily_return = (today_price - yesterday_price)*100 / yesterday_price
    return daily_return

def calc_moving_avg():
    return

# 算出所有的報酬率
for i in range(3, 97): # 3~96

    # Get current price and previous price
    today_price = tsmc_sheet.range(f"B{i}").value
    yesterday_price = tsmc_sheet.range(f"B{i-1}").value

    # Calc return value and write in excel
    daily_return = calc_daily_return(today_price, yesterday_price)
    tsmc_sheet.range(f"C{i}").value = daily_return
    print(str(i) + ": " + str(daily_return))

wb.save()