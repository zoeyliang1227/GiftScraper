import time

import bs4
import openpyxl
from selenium import webdriver


def get_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')                 # 瀏覽器不提供可視化頁面
    options.add_argument('-no-sandbox')               # 以最高權限運行
    options.add_argument('--start-maximized')        # 縮放縮放（全屏窗口）設置元素比較準確
    options.add_argument('--disable-gpu')            # 谷歌文檔說明需要加上這個屬性來規避bug
    options.add_argument('--window-size=1920,1080')  # 設置瀏覽器按鈕（窗口大小）
    options.add_argument('--incognito')               # 啟動無痕

    driver = webdriver.Chrome(chrome_options=options)
    driver.get(
        'https://giftshop-tw.line.me/search?useType=VOUCHER&searchValue=%E9%9B%BB%E5%AD%90')

    return driver


def search():
    driver = get_driver()
    wb = openpyxl.Workbook()
    sheet = wb.create_sheet("line", 0)
    # 先填入第一列的欄位名稱
    sheet['A1'] = 'text'
    sheet['B1'] = 'name'
    sheet['C1'] = 'price'

    for j in range(300):
        driver.execute_script(
            'window.scrollTo(0, document.body.scrollHeight);')
        time.sleep(1)

    soup = bs4.BeautifulSoup(driver.page_source, 'html.parser')
    points = soup.find_all('li', class_='product_item')
    print(len(points))
    for i in points:
        data = i.find('div', class_='info_area')
        # print(data)
        a = data.find('span', class_='text')
        b = data.find('p', class_='name').text
        c = data.find('span', class_='price').text
        if not a:
            pass
        else:
            # print(c.text, a, b)

            # 實際將資料寫入每一列
            sheet.append([a.text, b, c])

    wb.save("line.xlsx")


search()
