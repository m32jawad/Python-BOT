from __future__ import print_function
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import os.path
import pyodbc
import pandas as pd
from selenium.common.exceptions import NoSuchElementException
import numpy as np
from re import sub
import shutil
import glob
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from tqdm import tqdm

conn = pyodbc.connect(Driver='{ODBC Driver 18 for SQL Server}',
                      Server='tcp:test-server-jd.database.windows.net,1433',
                      Database='TEST',
                      Uid='test',
                      Pwd='@Admin$$')

cursor = conn.cursor()


class Browser:
    browser, service = None, None

    # Initialise the webdriver with the path to chromedriver.exe
    def __init__(self, driver: str):
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\Users\DELL\Downloads\downloads\\",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing_for_trusted_sources_enabled": False,
    "safebrowsing.enabled": False
})
        # options.add_argument("prefs={\"download.default_directory\":\"C:/Users/DELL/Downloads/python scripts/\"}")
        options.add_argument("--headless")
        self.service = Service(driver)
        self.browser = webdriver.Chrome(options=options, service=self.service)

    def open_page(self, url: str):
        self.browser.get(url)

    def close_browser(self):
        self.browser.close()

    def add_input(self, by: By, value: str, text: str):
        field = self.browser.find_element(by=by, value=value)
        field.send_keys(text)
        time.sleep(1)

    def click_button(self, by: By, value: str):
        button = WebDriverWait(self.browser, 10).until(
            EC.element_to_be_clickable((by, value))
        )
        # button = self.browser.find_element(by=by, value=value)
        button.click()
        time.sleep(1)

    def scroll_window(self):
        # Get the height of one row (adjust as needed)
        row_height = 50  # Replace with the actual height of one row

        # Scroll down by the height of 5 rows
        scroll_distance = 5 * row_height
        self.browser.execute_script(f"window.scrollBy(0, {scroll_distance});")

    def login_888lots(self, username: str, password: str):
        self.add_input(by=By.ID, value='input-email', text=username)
        self.add_input(by=By.ID, value='input-password', text=password)

    def login_selleramps(self, username: str, password: str):
        self.add_input(by=By.ID, value='loginform-email', text=username)
        self.add_input(by=By.ID, value='loginform-password', text=password)
        self.click_button(by=By.NAME, value='login-button')

    def fetch_selleramp_field_data(self, by: By, value: str):
        return self.browser.find_element(by=by, value=value)


def DownloadIndividualTableData():
    # Remove files from the 'python scripts' directory
    for f in glob.glob(r'C:/Users/DELL/Downloads/downloads/*'):
        try:
            print("Removing {}".format(f))
            os.remove(f)
        except Exception as e:
            print(f"Error removing file: {e}")


    browser.open_page('https://888lots.com/items')
    time.sleep(1)
    browser.click_button(By.CSS_SELECTOR, ".btn.btn-xs")
    time.sleep(1)
    browser.login_888lots(username='houstongoodsnetwork@gmail.com', password='Wesomarmo123!')
    time.sleep(1)

    browser.scroll_window()
    browser.click_button(By.CSS_SELECTOR, ".btn.btn-default.pull-right")
    time.sleep(1)
    browser.click_button(By.CSS_SELECTOR, ".btn.btn-info.dropdown-toggle.btn-sm")
    time.sleep(15)

    paths = sorted(glob.glob(r'C:/Users/DELL/Downloads/downloads/*.xlsx'), key=os.path.getmtime)
    if paths:
        print(paths)
        shutil.copy(paths[-1], 'C:/Users/DELL/Documents/Python Scripts/DELL/individual items download/recent download')

        # Rename the copied file
        filepath = os.listdir('C:/Users/DELL/Documents/Python Scripts/DELL/individual items download/recent download/')
        # filepath = sorted(filepath, key=os.path.getmtime)
        if filepath:
            os.rename(
                'C:/Users/DELL/Documents/Python Scripts/DELL/individual items download/recent download/' + filepath[-1],
                'C:/Users/DELL/Documents/Python Scripts/DELL/individual items download/recent download/' + 'IndividualItemsDownload.xlsx'
            )
    else:
        print("No Excel files found in 'downloads' directory.")

    # paths = sorted(glob.glob(r'C:/Users/DELL/Downloads/downloads/*.xlsx'), key=os.path.getmtime)
    # shutil.copy(paths[-1], 'C:/Users/DELL/Documents/Python Scripts/DELL/individual items download/recent download/')
    # filepath = os.listdir('C:/Users/DELL/Documents/Python Scripts/DELL/individual items download/recent download/')
    # os.rename('C:/Users/DELL/Documents/Python Scripts/DELL/individual items download/recent download/' + filepath[-1],
    #           'C:/Users/DELL/Documents/Python Scripts/DELL/individual items download/recent download/' + 'IndividualItemsDownload.xlsx')


def DeleteRecordsIndividualData():
    # conn = pyodbc.connect(Driver='{ODBC Driver 18 for SQL Server}',
    #                   Server='tcp:houstongoodnetwork.database.windows.net,1433',
    #                   Database='HoustonGoodsNetwork',
    #                   Uid = 'houstongoodsnetwork',
    #                   Pwd='')

    # cursor = conn.cursor()

    cursor.execute('''
        DELETE FROM Individual_Items
        '''
                   )
    conn.commit()


def PopulateIndividualTableData():
    # This will populate Individual items table in sql with asins that haven't been ran

    # conn = pyodbc.connect(Driver='{ODBC Driver 18 for SQL Server}',
    #                   Server='tcp:houstongoodnetwork.database.windows.net,1433',
    #                   Database='HoustonGoodsNetwork',
    #                   Uid = 'houstongoodsnetwork',
    #                   Pwd='')

    # cursor = conn.cursor()

    IndividualFile = pd.read_excel(
        "C:/Users/DELL/Documents/Python Scripts/DELL/individual items download/recent download/IndividualItemsDownload.xlsx")
    IndividualFile = IndividualFile.rename({'Unit Price': 'UnitPrice', 'Est. selling price on Az': 'EstSellPrice'},
                                           axis=1)

    IndividualFile.dropna(inplace=True)

    # Your existing SQL queries
    query_dead = "SELECT ASIN FROM [TEST].[dbo].[DeadedAsin];"
    query_sellerAmpGoodROI = "SELECT ASIN FROM [TEST].[dbo].[SellerAmpsBadROI];"
    query_sellerAmpBadROI = "SELECT ASIN FROM [TEST].[dbo].[SellerAmpsGoodROI];"
    query_Individual_Items = '''WITH CTE AS (
        SELECT
            *,
            ROW_NUMBER() OVER (PARTITION BY ASIN,SKU ORDER BY (SELECT NULL)) AS RowNum
        FROM
            [TEST].[dbo].[Individual_Items]
    )
    SELECT ASIN FROM CTE WHERE RowNum > 1;
'''
    # Fetch ASIN values from SQL and create a combined list
    asin_ran_list = pd.concat([
        pd.read_sql(query_dead, conn)['ASIN'],
        pd.read_sql(query_sellerAmpBadROI, conn)['ASIN'],
        pd.read_sql(query_sellerAmpGoodROI, conn)['ASIN'],
        pd.read_sql(query_Individual_Items, conn)['ASIN'],
    ]).tolist()

    print(asin_ran_list)
    print("pass")

    # Filter out rows with ASINs already present in the combined list
    IndividualFile = IndividualFile[~IndividualFile['ASIN'].isin(asin_ran_list)]
    print(len(IndividualFile))
    for row in IndividualFile.itertuples():
        cursor.execute('''
                    INSERT INTO Individual_Items (SKU, ASIN, Condition, Item, Category, Qty, UnitPrice, EstSellPrice)
                    VALUES (?,?,?,?,?,?,?,?)
                    ''',
                       row.SKU,
                       row.ASIN,
                       row.Condition,
                       row.Item,
                       row.Category,
                       row.Qty,
                       row.UnitPrice,
                       row.EstSellPrice
                       )

    conn.commit()
    print("end")


def GetIndividualTableData():
    # IndividualItemsQuery = f"SELECT SKU, ASIN, Condition, Item, Category, Qty, UnitPrice, EstSellPrice FROM [TEST].[dbo].[Individual_Items];"
    IndividualItemsQuery = '''SELECT
    SKU, ASIN, Condition, Item, Category, Qty, UnitPrice, EstSellPrice
    FROM[TEST].[dbo].[Individual_Items]
    WHERE
    ASIN
    NOT
    IN(
        SELECT
    ASIN
    FROM[TEST].[dbo].[SellerAmpsGoodROI]
    UNION
    SELECT
    ASIN
    FROM[TEST].[dbo].[SellerAmpsBadROI]
    UNION
    SELECT
    ASIN
    FROM[TEST].[dbo].[DeadedASIN]
    );
    '''
    return pd.read_sql(IndividualItemsQuery, conn)


def clean_value(datatype, value):
    if value != None:
        return datatype(sub(r'[^\d.]', '', value))
    else:
        return -1


def clean_value_buybox(datatype, value):
    if value != None:
        return datatype(sub(r'[^\d.]', '', value))
    else:
        return 0


def send_discord_alert(alert_msg):
    import requests
    webhook_url = "https://discord.com/api/webhooks/1189842899316260864/2T8RYWPaluWPF_v4QSgPwZwTymfloj3qpU137MB9txDTJ9HIbq5C4QOzp1PAvY4NXKbN"
    payload = {"content":f"{alert_msg}"}
    response = requests.post(webhook_url, json=payload)
    print(response.status_code, response.text)
    pass


if __name__ == '__main__':
  while True:
    browser = Browser(r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')

    # DeleteRecordsIndividualData()
    DownloadIndividualTableData()
    PopulateIndividualTableData()

    browser.close_browser()
    browser = Browser(r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')

    browser.open_page('https://sas.selleramp.com/')
    time.sleep(1)
    browser.login_selleramps(username='wesamh96@gmail.com', password='Wesomarmo123!')
    time.sleep(1)

    for i in tqdm(range(len(GetIndividualTableData()))):
    # for i in range(100):
        try:
            asin = str(GetIndividualTableData().iloc[i]['ASIN'])
            sku = str(GetIndividualTableData().iloc[i]['SKU'])
            condition = str(GetIndividualTableData().iloc[i]['Condition'])
            item = str(GetIndividualTableData().iloc[i]['Item'])
            category = str(GetIndividualTableData().iloc[i]['Category'])
            qty = int(GetIndividualTableData().iloc[i]['Qty'])
            unitprice = float(GetIndividualTableData().iloc[i]['UnitPrice'])
            estsellprice = float(GetIndividualTableData().iloc[i]['EstSellPrice'])

            browser.open_page(f'https://sas.selleramp.com/sas/lookup?SasLookup%5Bsearch_term%5D={asin}')
            time.sleep(1)
            browser.fetch_selleramp_field_data(By.ID, 'avg30').click()
            time.sleep(1)
            bsr30 = clean_value(int,
                                browser.fetch_selleramp_field_data(by=By.CLASS_NAME, value='rap_row ').get_attribute(
                                    "value"))
            buybox30 = clean_value_buybox(float, browser.fetch_selleramp_field_data(By.ID, "w1").find_element(By.ID,
                                                                                                              "keepa_csv_type_18").get_attribute(
                "value"))
            brand = str(browser.fetch_selleramp_field_data(by=By.XPATH,
                                                           value="//span[contains(@class,'pdb-manufacturer')]").get_attribute(
                "innerHTML"))
            LastRunDate = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if bsr30 == -1:
                # conn = pyodbc.connect(Driver='{ODBC Driver 18 for SQL Server}',

                #       Server='tcp:houstongoodnetwork.database.windows.net,1433',
                #       Database='HoustonGoodsNetwork',
                #       Uid = 'houstongoodsnetwork',
                #       Pwd='')

                # cursor = conn.cursor()

                cursor.execute('''
                    INSERT INTO DeadedAsin (ASIN)
                    VALUES (?)
                    ''',
                               asin
                               )
                conn.commit()
                # print(asin, " This Asset has been placed in the Deaded Asin Table")

            elif bsr30 >= 350000:
                # conn = pyodbc.connect(Driver='{ODBC Driver 18 for SQL Server}',
                #       Server='tcp:houstongoodnetwork.database.windows.net,1433',
                #       Database='HoustonGoodsNetwork',
                #       Uid = 'houstongoodsnetwork',
                #       Pwd='')
                # cursor = conn.cursor()

                cursor.execute('''
                    INSERT INTO DeadedAsin (ASIN)
                    VALUES (?)
                    ''',
                               asin
                               )
                conn.commit()
                # print(asin, " This Asset has been placed in the Deaded Asin Table")

            elif bsr30 < 350000 and bsr30 >= 0 and buybox30 == 0:  # Missing BuyBox30, classifying as BadROI
                # conn = pyodbc.connect(Driver='{ODBC Driver 18 for SQL Server}',
                #       Server='tcp:houstongoodnetwork.database.windows.net,1433',
                #       Database='HoustonGoodsNetwork',
                #       Uid = 'houstongoodsnetwork',
                #       Pwd='')
                # cursor = conn.cursor()

                cursor.execute('''
                    INSERT INTO SellerAmpsBadROI (SKU, ASIN, Condition, Item, Category, Qty, UnitPrice, EstSellPrice, brand, bsr30, buybox30, profit, totalfees, profitmargin, Notes, LastRunDate)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    ''',
                               sku,
                               asin,
                               condition,
                               item,
                               category,
                               qty,
                               unitprice,
                               estsellprice,
                               brand,
                               bsr30,
                               0,
                               0,
                               0,
                               0,
                               "No Buybox30 price found",
                               LastRunDate
                               )
                conn.commit()
                # print(asin, " This Asset has been placed in the Bad ROI Table")

            else:
                browser.fetch_selleramp_field_data(By.ID, 'cost').send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
                browser.add_input(By.ID, "cost", text=unitprice)
                browser.fetch_selleramp_field_data(By.ID, 'sale_price').send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
                browser.add_input(By.ID, "sale_price", text=buybox30)
                profit = clean_value(float, browser.fetch_selleramp_field_data(By.ID, value="w3").find_element(By.ID,
                                                                                                               "saslookup-profit").text)
                totalfees = clean_value(float, browser.fetch_selleramp_field_data(By.ID, value="w3").find_element(By.ID,
                                                                                                                  "saslookup-total_fee").text)
                profit_margin = round(((buybox30 - (totalfees + unitprice)) / (unitprice)) * 100, 2)
                LastRunDate = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                # conn = pyodbc.connect(Driver='{ODBC Driver 18 for SQL Server}',
                #       Server='tcp:houstongoodnetwork.database.windows.net,1433',
                #       Database='HoustonGoodsNetwork',
                #       Uid = 'houstongoodsnetwork',
                #       Pwd='')
                # cursor = conn.cursor()

                if (profit_margin >= 25 or (profit_margin >= 20 and qty >= 15 )):
                    cursor.execute('''
                        INSERT INTO SellerAmpsGoodROI (SKU, ASIN, Condition, Item, Category, Qty, UnitPrice, EstSellPrice, brand, bsr30, buybox30, profit, totalfees, profitmargin, LastRunDate)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        ''',
                                   sku,
                                   asin,
                                   condition,
                                   item,
                                   category,
                                   qty,
                                   unitprice,
                                   estsellprice,
                                   brand,
                                   bsr30,
                                   buybox30,
                                   profit,
                                   totalfees,
                                   profit_margin,
                                   LastRunDate
                                   )
                    conn.commit()
                    alert_msg = f"Product: {sku}, ASIN: {asin}, Category: {category} Profit: {profit}"
                    send_discord_alert(alert_msg)
                    # print(asin, " This Asset has been placed in the Good ROI Table")
                else:
                    cursor.execute('''
                        INSERT INTO SellerAmpsBadROI (SKU, ASIN, Condition, Item, Category, Qty, UnitPrice, EstSellPrice, brand, bsr30, buybox30, profit, totalfees, profitmargin, Notes, LastRunDate)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        ''',
                                   sku,
                                   asin,
                                   condition,
                                   item,
                                   category,
                                   qty,
                                   unitprice,
                                   estsellprice,
                                   brand,
                                   bsr30,
                                   buybox30,
                                   profit,
                                   totalfees,
                                   profit_margin,
                                   'Profit Margin too low',
                                   LastRunDate
                                   )
                    conn.commit()
                    # print(asin, " This Asset has been placed in the Bad ROI Table")

        except NoSuchElementException:

            # conn = pyodbc.connect(Driver='{ODBC Driver 18 for SQL Server}',
            #           Server='tcp:houstongoodnetwork.database.windows.net,1433',
            #           Database='HoustonGoodsNetwork',
            #           Uid = 'houstongoodsnetwork',
            #           Pwd='')
            # cursor = conn.cursor()

            cursor.execute('''
                    INSERT INTO DeadedAsin (ASIN)
                    VALUES (?)
                    ''',
                           asin
                           )
            conn.commit()
            # print(asin, " This Asset has been placed in the Deaded Asin Table")




