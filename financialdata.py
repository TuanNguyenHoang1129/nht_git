import pandas as pd
# import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType

def get_elements_value(items, convert_numeric=True):
    _ls = []
    if convert_numeric == False:
        for x in items:
            _ls.append(x.text)
    else:
        for x in items:
            num = x.text.strip().replace(',', '')
            if num.isnumeric():
                _ls.append(int(num))
            else:
                _ls.append(num)
    return _ls

def year_col_process(driver, col):
    x_path = f"//*[@id='tableContent']/tbody/tr/td[{col}]"
    rows = driver.find_elements(By.XPATH, x_path)
    return rows

def get_report_url(ticker, year, report_type):
    if report_type.upper() == 'BS':
        return f"https://s.cafef.vn/bao-cao-tai-chinh/{ticker}/BSheet/{year}/0/0/0/bao-cao-tai-chinh-cong-ty.chn"
    if report_type.upper() == 'IS':
        return f"https://s.cafef.vn/bao-cao-tai-chinh/{ticker}/IncSta/{year}/0/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty.chn"
    if report_type.upper() == 'CF':
        return f"https://s.cafef.vn/bao-cao-tai-chinh/{ticker}/CashFlow/{year}/0/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty.chn"
    
    return ""

def get_excel_data(driver, ticker, from_year, to_year, report_type):
    year = to_year    
    data = {}
    criteria_names = {}
    runable = True
    while(runable):
        url = get_report_url(ticker, year, report_type)
        print(url)
        driver.get(url)
        driver.implicitly_wait(5)


        if 'criteria' not in criteria_names:
            name_elements = year_col_process(driver, 1)
            criteria_names['criteria'] = get_elements_value(name_elements, False)


        year_elements = driver.find_elements(By.XPATH, "//*[@id='tblGridData']/tbody/tr/td")
        index_cols = {}
        i = 1
        for item in year_elements:
            str_item = item.text.strip()
            if str_item.isnumeric():
                index_cols[str_item] = i
            i += 1

        index_cols = dict(sorted(index_cols.items(), reverse=True))

        y = 0
        for key in index_cols:
            col = index_cols[key]
            items = year_col_process(driver, col)
            data[key] = get_elements_value(items)
            print(key)

            y = int(key)
            if y == from_year:
                runable = False
                break

        if runable == True:
            year = y - 1

    data = dict(sorted(data.items())) # Short dict
    data = criteria_names | data # Merge two dicts into one
    df = pd.DataFrame(data)

    # save_as = f"{ticker}_{report_type}.xlsx"
    # writer = pd.ExcelWriter(save_as,
    #                        engine='xlsxwriter',
    #                        engine_kwargs={'options': {'strings_to_numbers': True}})
    # df.to_excel(writer, sheet_name=report_type, index=False)
    # writer.close()
    
    return df

TICKER = 'FPT'
TO_YEAR = 2022
LOOK_BACK = 5
FROM_YEAR = TO_YEAR - LOOK_BACK + 1
report_types = 'BS'

service = ChromeService(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install())
chrome_options = ChromeOptions()
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--headless=new")
driver_auto = webdriver.Chrome(service=service, options=chrome_options)

get_excel_data(driver_auto, TICKER, FROM_YEAR, TO_YEAR, report_types)