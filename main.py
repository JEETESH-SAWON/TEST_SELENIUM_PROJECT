import time
import xlsxwriter
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


# ---------------------------------THIS IS MY FIRST SELENIUM PROJECT ----------------------------------#

def fetch_data(name, quantity, arg):
    item_no = 1
    page_no = 1
    per_page_item = 12
    max_item_to_fetch = quantity
    item_name = []
    item_price = []
    item_link = []
    item_mod = []

    if arg == 1:
        url = "https://www.amazon.com/s?k=" + name + "&page=" + str(page_no)
        site = "AMAZON"
    else:
        url = "https://www.flipkart.com/search?q=" + name + "&page=" + str(page_no)
        site = "FLIPCART"

    options = Options()
    # Put headless = True for Back-end Browser Run #
    options.headless = False
    browser = webdriver.Chrome(executable_path="driver/chromedriver.exe", options=options)
    browser.maximize_window()
    browser.get(url)
    browser.set_page_load_timeout(10)

    # Logic for fetching data form Amazon AND FLIPCART

    while True:
        try:
            if per_page_item * (page_no - 1) > max_item_to_fetch or item_no > max_item_to_fetch:
                break

            if item_no > per_page_item:
                item_no = 1
                page_no += 1

            if arg == 1:

                # Get Item Element to Get data from

                xpath_of_element = '//*[@id="search"]/div[1]/div[2]/div/span[3]/div[2]/div[' + str(
                    item_no) + ']/div/span/div/div/div[2]/div[2]/div/div[1]/div/div/div[1]/h2/a/span'

                browser.find_element_by_xpath(xpath_of_element).click()

                browser.set_page_load_timeout(10)

                # Get product Name

                product_name = browser.find_element_by_xpath('//span[@id="productTitle"]').text.strip()

                time.sleep(2)

                # Get Product Price

                product_price = browser.find_element_by_xpath("(//span[contains(@class,'a-color-price')])[1]").text

                # model no
                tempmod = browser.find_element_by_xpath("//*[contains(text(), 'Item model number')]")
                modpath = tempmod.find_element_by_xpath("..")
                modno = modpath.find_element_by_class_name('prodDetAttrValue').text

                # print(modno)---------Testing Purpose-------#

            else:
                # Get Item Element to Get data from

                xpath_of_element = '//*[@id="container"]/div/div[3]/div[1]/div[2]/div[' + str(
                    item_no) + ']/div/div/div/a'

                element = browser.find_element_by_xpath(xpath_of_element)
                newurl = element.get_attribute('href')

                # print(newurl)------Testing Purpose-------#
                browser.get(newurl)

                browser.set_page_load_timeout(10)

                # Get product Name

                product_name = browser.find_element_by_xpath(
                    '//*[@id="container"]/div/div[3]/div[1]/div[2]/div[2]/div/div[1]/h1').text
                # print(product_name)----------Testing Purpose-------#

                # Get Product Price

                product_price = browser.find_element_by_class_name("_30jeq3").text

                # Get MODEL NO
                tempmod = browser.find_element_by_xpath("//*[contains(text(), 'Model Number')]")
                modpath = tempmod.find_element_by_xpath("..")
                modno = modpath.find_element_by_class_name('_21lJbe').text

                # print(modno)---------Testing Purpose---------#

            # Get Product Url and Append
            item_link.append(browser.current_url)
            item_name.append(product_name)
            item_price.append(product_price)
            item_mod.append(modno)
            time.sleep(3)
            browser.back()

            browser.set_page_load_timeout(10)

            item_no += 1

        except Exception as e:
            # print("Exception On Count", item_no, e)
            # ---for testing Exceptions---
            item_no += 1

            if per_page_item * (page_no - 1) > max_item_to_fetch or item_no > max_item_to_fetch:
                break

            if item_no > per_page_item:
                item_no = 1
                page_no += 1

            if arg == 1:
                url = "https://www.amazon.com/s?k=" + name + "&page=" + str(page_no)
            else:
                url = "https://www.flipkart.com/search?q=" + name + "&page=" + str(page_no)

            browser.get(url)
            browser.set_page_load_timeout(10)

    browser.close()
    file_create(item_name, item_price, item_link, item_mod, site, arg)


def file_create(name, price, url, model, site, arg):
    dic = {'Name': name, 'Price': price, 'URL': url, 'Model-No': model, 'SITE': site}
    data = pd.DataFrame(dic)
    if arg == 1:
        writer = pd.ExcelWriter('sheet1.xlsx', engine='xlsxwriter')
        data.to_excel(writer, sheet_name="sheet1")
        format_sheet = writer.sheets['sheet1']
    else:
        writer = pd.ExcelWriter('sheet2.xlsx', engine='xlsxwriter')
        data.to_excel(writer, sheet_name="sheet2")
        format_sheet = writer.sheets['sheet2']

    format_sheet.set_column('B:B', 40)
    format_sheet.set_column('C:C', 10)
    format_sheet.set_column('D:D', 100)

    writer.save()


# ---------------HERE THE MAIN FUNCTION TO CALL FROM------------------#
def main():
    print("Welcome To Application")

    product_name = input("Enter product to search = ")
    product_quantity = int(input("input 1 for less than 10 items or 2 for item quantity b/w 10-50 = "))
    if product_quantity == 1:
        t = 10
    else:
        t = 50
    x = int(t / 2)

    fetch_data(str(product_name), x, 1)

    fetch_data(str(product_name), x, 2)


if __name__ == "__main__":
    main()
