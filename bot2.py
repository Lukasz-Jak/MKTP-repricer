from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
from openpyxl.styles import Font, Color, colors
import time
from datetime import datetime
import pandas as pd
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException 


PATH = "C:\Program Files (x86)\chromedriver.exe"
# driver = webdriver.Chrome(PATH)
driver = webdriver.Chrome(ChromeDriverManager().install())


def check_other_offers(xpath):
    try:
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True


# def filter_fn(row):
#     if row["Brand"] == "Revlon" and row["Category name"] == "Tratamente si masca de par":
#         return False
#     else:
#         return True

# calc_df = pd.DataFrame(d, columns=['Name', 'Age', 'Score'])
# m = calc_df.apply(filter_fn, axis=1)
# print(m)

base_file_df = pd.read_excel(r"C:\Users\jakub\Documents\Trena - emag\Repricing_eMAG_FBE_test.xlsx")
#print(base_file_df)
main_df = base_file_df[["Brand", "Category", "EAN", "Product code", "Part number key (PNK)", "Status", "Actual stock", "Position", "Av. Sale price", "R (5)", "Minimum price", "Maximum price", "eMAG URL"]].copy()
main_df = main_df.loc[(main_df["Status"] == 1)]
main_df["Price"] = 0


#options = ChromeOptions()
#options.add_argument("--start-maximized")
#python bot.pydriver = webdriver.Chrome(options=options)

sellers_ratings = {}
    


for index in main_df.index:
    driver.get(main_df["eMAG URL"][index])
    time.sleep(5)
    try:
        price = driver.find_element_by_xpath('//*[@id="main-container"]/section[1]/div/div[2]/div[2]/div/div/div[2]/form/div[1]/div[1]/div/div/p[2]').text.split(" ")[0]
        price = float(price[:-2] + "." + price[-2:])
        print(price)
        main_df.loc[main_df.index[index], "Price"] = price
    except:
        print("N/A")
        main_df.loc[main_df.index[index], "Price"] = "N/A"

    try:
        seller_name = driver.find_element_by_xpath('//a[contains(@href, "see_vendor_page")]').text
        print(seller_name)
        main_df.loc[main_df.index[index], "Seller_name"] = seller_name
    except:
        print("N/A")
        main_df.loc[main_df.index[index], "Seller_name"] = "N/A"

    # Seller rating

    try:
        seller_rating = driver.find_element_by_xpath("//span[contains(@class, 'js-seller-rating-container seller-rating-container font-size-sm gtm_K9Zfl1J3')]").text
        seller_rating = float(seller_rating)
        print("Rating from listing: ")
        main_df.loc[main_df.index[index], "Seller_rating"] = seller_rating
        sellers_ratings.update({seller_name:seller_rating})
    except:
        if seller_name != "N/A" and seller_name not in sellers_ratings:
            sellers_ratings.update({seller_name:seller_name})
            try:
                vendor_page = driver.find_element_by_xpath('//a[contains(@href, "see_vendor_page")]').click()
                time.sleep(5)
                seller_rating = driver.find_element_by_xpath("//div[contains(@class, 'vendor-general-rating-number font-bold mrg-rgt-sm')]").text
                seller_rating = float(seller_rating)
                print("Rating from vendor page: ")
                main_df.loc[main_df.index[index], "Seller_rating"] = seller_rating
                sellers_ratings.update({seller_name:seller_rating})
            except:
                print("Drugi except - 65 line")
        else:
            print("Odczytany wcześniej lub błąd w seller name")

    # Other offers
    
    if check_other_offers("//h3[contains(text(), 'Alte oferte')]") == True:
        try:
            price_2 = driver.find_element_by_xpath('//div[starts-with(@class, "table-cell-sm va-middle po-size-mdsc-sm po-size-smsc-sm po-size-xssc-full po-pad-xl po-text-small")]').text.split(" ")[0]
            price_2 = float(price_2[:-2] + "." + price_2[-2:])
            print(price_2)
            main_df.loc[main_df.index[index], "Price_2"] = price_2
        except:
            print("Price_2_fail")

        try:
            seller_2_name = driver.find_element_by_xpath('//a[contains(@href, "see_vendor_page_so")]').text
            print(seller_2_name)
            main_df.loc[main_df.index[index], "Seller_2_name"] = seller_2_name
        except:
            print("N/A")
            main_df.loc[main_df.index[index], "Seller_2_name"] = "N/A"

        try:
            seller_2_rating = driver.find_element_by_xpath('//span[contains(@class, "seller-rating-container js-seller-rating-container")]').text
            seller_2_rating = float(seller_2_rating)
            print("Rating from listing: ")
            main_df.loc[main_df.index[index], "Seller_2_rating"] = seller_2_rating
            #sellers_ratings.update({seller_name:seller_rating})
        except:
            print("N/A")
            main_df.loc[main_df.index[index], "Seller_2_rating"] = "N/A"

        # Seller 1 rating - Seller 2 rating
        # ratings_diff = main_df.loc(main_df(["Seller rating"][index])) - main_df.loc(main_df(["Seller 2 rating"][index]))


    else:
        print("No other offers")
        main_df.loc[main_df.index[index], "Price_2"] = ""
        main_df.loc[main_df.index[index], "Seller_2_name"] = ""
        main_df.loc[main_df.index[index], "Seller_2_rating"] = ""


driver.quit()
print(sellers_ratings)   
print(main_df)

# Zapisywanie pliku 
timestamp = datetime.now().strftime("%Y_%m_%d %H-%M-%S")
main_df.to_excel("Prices_" + str(timestamp) + ".xlsx", index=True)

# Wyliczanie cen
pd.set_option('chained', None)
calc_df = main_df
# print(calc_df)

calc_df["Ratings_diff"] = calc_df["Seller_rating"] - calc_df["Seller_2_rating"]
print(calc_df)

# calc_df = calc_df.loc[(calc_df["Price"] != "N/A") or (calc_df["Price_2"] != "Price_2_fail")]

# V1 - BuyBox - podnosi cenę względem ceny drugiego sprzedawcy, ale nie obniża poniżej pierwotnie ustawionej "Price". Pomija przypadki gdy na drugiej pozycji jest Trena FBM.
df_bb = calc_df.loc[(calc_df["Seller_name"] == "Trena FBE") & (calc_df["Seller_2_name"] != "Trena FBM") & (calc_df["R (5)"] > 2)]
df_bb["New_sale_price_gross"] = ((df_bb["Price_2"] - 0.1) * ((df_bb["Ratings_diff"] / 10) + 1.02)).round(decimals = 2)
df_bb = df_bb.loc[(df_bb["New_sale_price_gross"] > df_bb["Price"])]
df_bb["New_sale_price_net"] = (df_bb["New_sale_price_gross"] / 1.19).round(decimals = 4)
print(df_bb)

# V2 - Trena FBE na 2 pozycji - obniża cenę względem ceny z BuyBux'a. Nie ustawi wyższej niż pierwotnie była ("Price_2"). Pomija przypadki gdy w BB jest Trena FBM.
df_p2 = calc_df.loc[(calc_df["Seller_2_name"] == "Trena FBE") & (calc_df["Seller_name"] != "Trena FBM") & (calc_df["R (5)"] < 10)]
df_p2["New_sale_price_gross"] = ((df_p2["Price"] - 0.05) * (((df_p2["Ratings_diff"]) * (-1) / 10) + 1)).round(decimals = 2)
df_p2 = df_p2.loc[(df_p2["New_sale_price_gross"] < df_p2["Price_2"])]
df_p2["New_sale_price_net"] = (df_p2["New_sale_price_gross"] / 1.19).round(decimals = 4)
print(df_p2)

# V3 - Brak Trena FBE jako pierwszy lub drugi sprzedawca
# calc_df = calc_df.loc[(calc_df["Seller_name"] != "Trena FBE") & (calc_df["Seller_2_name"] != "Trena FBE")]

# V4 - brak innych sprzedawców 

df_res = pd.concat([df_bb, df_p2])
df_res.loc[df_res["New_sale_price_net"] < df_res["Minimum price"], "New_sale_price_net"] = df_res["Minimum price"]
df_res.loc[df_res["New_sale_price_net"] > df_res["Minimum price"], "New_sale_price_net"] = df_res["New_sale_price_net"]
df_res.loc[df_res["New_sale_price_net"] > df_res["Maximum price"], "New_sale_price_net"] = df_res["Maximum price"]
df_res.loc[df_res["New_sale_price_net"] < df_res["Maximum price"], "New_sale_price_net"] = df_res["New_sale_price_net"]
print(df_res)
timestamp = datetime.now().strftime("%Y_%m_%d %H-%M-%S")
df_res.to_excel("Res_prices_" + str(timestamp) + ".xlsx", index=True)

