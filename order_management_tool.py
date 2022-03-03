# -*- coding: utf-8 -*-
import atexit
import os
# fix issue: urlopen error [SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed.
import ssl
ssl._create_default_https_context = ssl._create_unverified_context
import time
import warnings
warnings.filterwarnings("ignore")
import xlwt
import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

_Workbook = object()


def order_management_tool():
    options = webdriver.ChromeOptions()
    prefs = {"profile.content_settings.exceptions.clipboard": 1}
    options.add_experimental_option("prefs", prefs)
    driver = uc.Chrome(version_main=98)
    driver.set_page_load_timeout(60)
    driver.get("https://mms.pinduoduo.com/orders/list")
    driver.maximize_window()

    input_cmd = ""
    while input_cmd.lower() != "y" and input_cmd.lower() != "n":
        input_cmd = input("是否已经完成登录? 输入y/Y确认登录; 输入n/N停止程序.\n>> ")
        if input_cmd.lower() != "y" and input_cmd.lower() != "n":
            print("错误输入, 请重新输入 ...")
    
    if input_cmd.lower() == "y":
        global _Workbook
        _Workbook = xlwt.Workbook()
        sheet = _Workbook.add_sheet("待发货信息")
        header_font = xlwt.Font()
        header_font.name = "Arial"
        header_font.bold = True
        header_style = xlwt.XFStyle()
        header_style.font = header_font
        sheet.write(0, 0, "订单编号", header_style)
        sheet.write(0, 1, "收件人", header_style)
        sheet.write(0, 2, "手机", header_style)
        sheet.write(0, 3, "地址", header_style)
        sheet.write(0, 4, "发货信息", header_style)
        sheet.write(0, 5, "商品ID", header_style)
        sheet.write(0, 6, "发货数量", header_style)
        row_num = 1

        while input_cmd.lower() == "y":
            print("当前网页: {}".format(driver.current_url))
            order_table = WebDriverWait(driver, 600).until(
                EC.presence_of_element_located((By.ID, "order-content"))
            )
            order_info_list = WebDriverWait(order_table, 600).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, "tbody"))
            )
            for order_info in order_info_list:
                info_blocks = WebDriverWait(order_info, 600).until(
                    EC.presence_of_all_elements_located((By.TAG_NAME, "td"))
                )
                sku = info_blocks[1].text.split("\n")[2]
                sheet.write(row_num, 4, sku.strip())
                id = info_blocks[1].text.split("\n")[1].lstrip("ID:")
                sheet.write(row_num, 5, id.strip())
                qty = info_blocks[3].text
                sheet.write(row_num, 6, qty.strip())

                copy_order_code_btn = WebDriverWait(order_info, 600).until(
                    EC.presence_of_element_located((By.LINK_TEXT, "复制"))
                )
                driver.execute_script("arguments[0].scrollIntoView();", copy_order_code_btn)
                driver.execute_script("arguments[0].click();", copy_order_code_btn)
                time.sleep(0.4)
                order_code = copy_order_code_btn.get_attribute("data-clipboard-text")
                sheet.write(row_num, 0, order_code.strip())
                check_user_info_btn = WebDriverWait(order_info, 600).until(
                    EC.presence_of_element_located((By.LINK_TEXT, "查看"))
                )
                driver.execute_script("arguments[0].scrollIntoView();", check_user_info_btn)
                driver.execute_script("arguments[0].click();", check_user_info_btn)
                time.sleep(0.4)
                check_phone_number_btn = WebDriverWait(order_info, 600).until(
                    EC.presence_of_element_located((By.LINK_TEXT, "查看手机号"))
                )
                driver.execute_script("arguments[0].scrollIntoView();", check_phone_number_btn)
                driver.execute_script("arguments[0].click();", check_phone_number_btn)
                copy_user_info_btn = WebDriverWait(order_info, 600).until(
                    EC.presence_of_element_located((By.LINK_TEXT, "复制完整信息"))
                )
                driver.execute_script("arguments[0].scrollIntoView();", copy_user_info_btn)
                driver.execute_script("arguments[0].click();", copy_user_info_btn)
                time.sleep(1.2)
                user_info = copy_user_info_btn.get_attribute("data-clipboard-text")
                username = user_info.split("\n")[0]
                sheet.write(row_num, 1, username.strip())
                phone = user_info.split("\n")[1]
                sheet.write(row_num, 2, phone.strip())
                address = user_info.split("\n")[2]
                sheet.write(row_num, 3, address.strip())

                row_num += 1
            input_cmd = input("是否已经跳转下一页? 输入y/Y确认跳转; 输入n/N停止程序.\n>> ")
    
    driver.close()
    driver.quit()


def save_workbook():
    if type(_Workbook) is not object:
        _Workbook.save("{}/Desktop/待发货表-{}.xls".format(
            os.path.expanduser("~"),
            time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        ))


if __name__ == "__main__":
    atexit.register(save_workbook)
    order_management_tool()
