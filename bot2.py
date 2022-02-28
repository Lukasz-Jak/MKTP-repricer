from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
import xlsxwriter
from openpyxl.styles import Font, Color, colors
import time
from datetime import datetime
import pandas as pd
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException 


PATH = "C:/Users/lukasz.jakubowski/Downloads/instalki/chromedriver.exe"
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

base_file_df = pd.read_excel(r"C:\Users\lukasz.jakubowski\Documents\eMAG\REPRICING\Repricing - eMAG FBE - test.xlsx")
#print(base_file_df)
main_df = base_file_df[["Brand", "Category", "EAN", "Product code", "Part number key (PNK)", "Status", "Actual stock", "Position", "Av. Sale price", "R (5)", "Minimum price", "Maximum price", "eMAG URL", "Ignore"]].copy()
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
        main_df.loc[main_df.index[index], "Price"] = int(0)

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
            main_df.loc[main_df.index[index], "Price_2"] = int(0)

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
            main_df.loc[main_df.index[index], "Seller_2_rating"] = seller_rating

        # Seller 1 rating - Seller 2 rating
        # ratings_diff = main_df.loc(main_df(["Seller rating"][index])) - main_df.loc(main_df(["Seller 2 rating"][index]))


    else:
        print("No other offers")
        main_df.loc[main_df.index[index], "Price_2"] = int(0)
        main_df.loc[main_df.index[index], "Seller_2_name"] = ""
        main_df.loc[main_df.index[index], "Seller_2_rating"] = seller_rating


driver.quit()
print(sellers_ratings)   
print(main_df)

# Zapisywanie pliku "Prices"
#Region
# timestamp = datetime.now().strftime("%Y_%m_%d %H-%M-%S")
# main_df.to_excel("Prices_" + str(timestamp) + ".xlsx", index=True)

timestamp = datetime.now().strftime("%Y_%m_%d %H-%M-%S")
writer = pd.ExcelWriter("Prices_" + str(timestamp) + ".xlsx", engine="xlsxwriter")
# main_df = main_df.drop("Ignore", 1)
main_df.to_excel(writer, index=False, sheet_name="Prices", startrow=1)
workbook = writer.book
worksheet = writer.sheets["Prices"]
worksheet.set_zoom(80)

header_format = workbook.add_format({
        "valign": "vcenter",
        "align": "center",
        "bold": True,
        "font_color": "#ECF2EF",
        "bg_color": "#011F14"
    })

# Title
title = ("Price update FBE - " + str(timestamp))
title_format = workbook.add_format({'bold': True})
title_format.set_font_size(18)
title_format.set_font_color("#393E3C")

worksheet.merge_range("A1:N1", title, title_format)
worksheet.set_row(1, 16)

for col_num, value in enumerate(main_df.columns.values):
    worksheet.write(1, col_num, value, header_format)

# title 2
title_2 = ("BuyBox")
title_2_format = workbook.add_format({'bold': True, "bg_color": "#FAC948"})
title_2_format.set_font_size(14)
title_2_format.set_font_color("#393E3C")
title_2_format.set_align('center')

worksheet.merge_range("O1:Q1", title_2, title_2_format)

# title 3
title_3 = ("2nd position")
title_3_format = workbook.add_format({'bold': True, "bg_color": "#30DCEF"})
title_3_format.set_font_size(14)
title_3_format.set_font_color("#393E3C")
title_3_format.set_align('center')

worksheet.merge_range("R1:T1", title_3, title_3_format)

# formatowanie kolumn pobranych z pliku
src_fmt = workbook.add_format({
        "align": "center",
        "bg_color": "#D5D8D7",
        'border': 1
    })
worksheet.set_column("A:A", 24, src_fmt)
worksheet.set_column("B:B", 30, src_fmt)
worksheet.set_column("C:C", 16, src_fmt)
worksheet.set_column("D:D", 12, src_fmt)
worksheet.set_column("E:E", 20, src_fmt)
worksheet.set_column("F:F", 12, src_fmt)
worksheet.set_column("G:M", 14, src_fmt)
worksheet.set_column("N:N", 12, src_fmt)

# formatowanie kolumn danych z BuyBox'a
bbdata_fmt = workbook.add_format({
        "align": "center",
        "bg_color": "#FBEBCF",
        'border': 1
    })
worksheet.set_column("O:Q", 14, bbdata_fmt)

# formatowanie kolumn danych z 2nd position
nddata_fmt = workbook.add_format({
        "align": "center",
        "bg_color": "#DDFFFD",
        'border': 1
    })
worksheet.set_column("R:T", 14, nddata_fmt)

# formatowanie waluty (round 2 decimal places)
currency_fmt = workbook.add_format({
        'num_format': '#,##0.00', 
        'border': 1,
        "align": "center",
        "bg_color": "#D5D8D7"
    })

worksheet.set_column("J:J", 12, currency_fmt)

writer.save()
#Endregion


# Wyliczanie cen
pd.set_option('chained', None)
calc_df = main_df.copy()
# print(calc_df)

calc_df["Ratings_diff"] = calc_df["Seller_rating"] - calc_df["Seller_2_rating"]
print(calc_df)

# calc_df = calc_df.loc[(calc_df["Price"] != "N/A") or (calc_df["Price_2"] != "Price_2_fail")]

# V1 - BuyBox - podnosi cenę względem ceny drugiego sprzedawcy, ale nie obniża poniżej pierwotnie ustawionej "Price". Pomija przypadki gdy na drugiej pozycji jest Trena FBM. Pomija przypadki gdy na 2 pozycji jest Trena FBM.
df_bb = calc_df.loc[(calc_df["Seller_name"] == "Trena FBE") & (calc_df["Seller_2_name"] != "Trena FBM") & (calc_df["R (5)"] > 2) & (calc_df["Ignore"] != 1)]
df_bb["New_sale_price_gross"] = ((df_bb["Price_2"] - 0.2) * ((df_bb["Ratings_diff"] / 10) + 1)).round(decimals = 2)
df_bb = df_bb.loc[(df_bb["New_sale_price_gross"] > df_bb["Price"])]
df_bb["New_sale_price_net"] = (df_bb["New_sale_price_gross"] / 1.19).round(decimals = 4)
print(df_bb)

# V2 - Trena FBE na 2 pozycji - obniża cenę względem ceny z BuyBux'a. Nie ustawia wyższej niż pierwotnie ustawiona cena ("Price_2"). Pomija przypadki gdy w BB jest Trena FBM.
df_p2 = calc_df.loc[(calc_df["Seller_2_name"] == "Trena FBE") & (calc_df["Seller_name"] != "Trena FBM") & (calc_df["R (5)"] < 10) & (calc_df["Ignore"] != 1)]
df_p2["New_sale_price_gross"] = ((df_p2["Price"] - 0.1) * (((df_p2["Ratings_diff"]) * (-1) / 10) + 1)).round(decimals = 2)
df_p2 = df_p2.loc[(df_p2["New_sale_price_gross"] < df_p2["Price_2"])]
df_p2["New_sale_price_net"] = (df_p2["New_sale_price_gross"] / 1.19).round(decimals = 4)
print(df_p2)

# V3 - Brak Trena FBE jako pierwszy lub drugi sprzedawca. Ustawia cenę względem ceny z BuyBox'a.
df_v3 = calc_df.loc[(calc_df["Seller_2_name"] != "Trena FBE") & (calc_df["Seller_name"] != "Trena FBE") & (calc_df["Actual stock"] > 0) & (calc_df["Ignore"] != 1)]
df_v3["New_sale_price_gross"] = ((df_v3["Price"] - 0.1) * (((df_v3["Ratings_diff"]) * (-1) / 10) + 1)).round(decimals = 2)
# df_v3 = df_v3.loc[(df_v3["New_sale_price_gross"] < df_v3["Price_2"])] - powinno sprawdzać czy cena jest mniejsza niż ta, która była ustawiona według pliku
df_v3["New_sale_price_net"] = (df_v3["New_sale_price_gross"] / 1.19).round(decimals = 4)
print(df_v3)

# V4 - Brak innych sprzedawców i rotacja (5) mniejsza niż 2 szt. -> obniżka ceny o 1 RON brutto
df_v4 = calc_df.loc[(calc_df["Seller_name"] == "Trena FBE") & (calc_df["Seller_2_name"] == "") & (calc_df["Actual stock"] > 0) & (calc_df["R (5)"] < 2) & (calc_df["Ignore"] != 1)]
df_v4["New_sale_price_gross"] = (df_v4["Price"] - 1).round(decimals = 2)
df_v4["New_sale_price_net"] = (df_v4["New_sale_price_gross"] / 1.19).round(decimals = 4)
print(df_v4)

#V5 -


df_res = pd.concat([df_bb, df_p2, df_v3, df_v4])
df_res.loc[df_res["New_sale_price_net"] < df_res["Minimum price"], "New_sale_price_net"] = df_res["Minimum price"]
df_res.loc[df_res["New_sale_price_net"] > df_res["Minimum price"], "New_sale_price_net"] = df_res["New_sale_price_net"]
df_res.loc[df_res["New_sale_price_net"] > df_res["Maximum price"], "New_sale_price_net"] = df_res["Maximum price"]
df_res.loc[df_res["New_sale_price_net"] < df_res["Maximum price"], "New_sale_price_net"] = df_res["New_sale_price_net"]

df_res["Product code"] = df_res["Product code"].astype(str)
df_res["EAN"] = df_res["EAN"].astype(str)


# Zapisywanie pliku "New prices"
#Region

timestamp = datetime.now().strftime("%Y_%m_%d %H-%M-%S")
writer = pd.ExcelWriter("New_prices_" + str(timestamp) + ".xlsx", engine="xlsxwriter")
df_res = df_res.drop("Ignore", 1)
df_res.to_excel(writer, index=False, sheet_name="New_prices", startrow=1)
workbook = writer.book
worksheet = writer.sheets["New_prices"]
worksheet.set_zoom(80)

header_format = workbook.add_format({
        "valign": "vcenter",
        "align": "center",
        "bold": True,
        "font_color": "#ECF2EF",
        "bg_color": "#011F14"
    })

# Title
title = ("Price update FBE - " + str(timestamp))
title_format = workbook.add_format({'bold': True})
title_format.set_font_size(18)
title_format.set_font_color("#393E3C")

worksheet.merge_range("A1:M1", title, title_format)
worksheet.set_row(1, 16)

for col_num, value in enumerate(df_res.columns.values):
    worksheet.write(1, col_num, value, header_format)

# title 2
title_2 = ("BuyBox")
title_2_format = workbook.add_format({'bold': True, "bg_color": "#FAC948"})
title_2_format.set_font_size(14)
title_2_format.set_font_color("#393E3C")
title_2_format.set_align('center')

worksheet.merge_range("N1:P1", title_2, title_2_format)

# title 3
title_3 = ("2nd position")
title_3_format = workbook.add_format({'bold': True, "bg_color": "#30DCEF"})
title_3_format.set_font_size(14)
title_3_format.set_font_color("#393E3C")
title_3_format.set_align('center')

worksheet.merge_range("Q1:S1", title_3, title_3_format)

# title_4
title_4 = ("Results")
title_4_format = workbook.add_format({'bold': True, "bg_color": "#38B15A"})
title_4_format.set_font_size(14)
title_4_format.set_font_color("#393E3C")
title_4_format.set_align('center')

worksheet.merge_range("T1:V1", title_4, title_4_format)

# formatowanie kolumn pobranych z pliku
src_fmt = workbook.add_format({
        "align": "center",
        "bg_color": "#D5D8D7",
        'border': 1
    })
worksheet.set_column("A:A", 18, src_fmt)
worksheet.set_column("B:B", 26, src_fmt)
worksheet.set_column("C:C", 16, src_fmt)
worksheet.set_column("D:D", 12, src_fmt)
worksheet.set_column("E:E", 20, src_fmt)
worksheet.set_column("F:J", 12, src_fmt)
worksheet.set_column("K:L", 18, src_fmt)
worksheet.set_column("M:M", 14, src_fmt)

# formatowanie kolumn danych z BuyBox'a
bbdata_fmt = workbook.add_format({
        "align": "center",
        "bg_color": "#FBEBCF",
        'border': 1
    })
worksheet.set_column("N:P", 14, bbdata_fmt)

# formatowanie kolumn danych z 2nd position
nddata_fmt = workbook.add_format({
        "align": "center",
        "bg_color": "#DDFFFD",
        'border': 1
    })
worksheet.set_column("Q:S", 14, nddata_fmt)

# formatowanie kolumn danych z results
results_fmt = workbook.add_format({
        "align": "center",
        "bg_color": "#C6F4BA",
        'border': 1
    })
worksheet.set_column("T:T", 12, results_fmt)
worksheet.set_column("U:V", 18, results_fmt)

# formatowanie waluty (round 2 decimal places)
currency_fmt = workbook.add_format({
        'num_format': '#,##0.00', 
        'border': 1,
        "align": "center",
        "bg_color": "#D5D8D7"
    })

worksheet.set_column("I:I", 12, currency_fmt)

writer.save()

#Endregion

print(df_res)

# df_res.to_excel("Res_prices_" + str(timestamp) + ".xlsx", index=True)
