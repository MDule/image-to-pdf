from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from pywinauto.application import Application
from pywinauto.mouse import move, click
from API.pdf.tools import Compress
from API.pdf.tools import Merge
import os
import time
import sys
from itertools import groupby
from operator import itemgetter

##############
# FOLDERI
##############

# folderi koji treba da se prebace prvo u pdf pa compress
folderi = [

    [r"C:\Users\User\Desktop\Folder1", "01-0"],
    [r"C:\Users\User\Desktop\Folder2", "01-1"],
    [r"C:\Users\User\Desktop\Folder3", "02-1"]
]


###############
# Prerequisites
###############


#API keys for different services
api = [
    "api_key_1",
    "api_key_2",
    "api_key_3",
    "api_key_4",
    "api_key_5",
    "api_key_6",
    "api_key_7",
    "api_key_8",
    "api_key_9"
]

api_cntr = 5
api_name = api[5]


###########
# Functions
###########

def selenium(path_dir):

    # da se napravi type_dir od folera koje smo deklarisali gore
    type_dir = path_dir.replace(" ", "{SPACE}")

    try:
        # otvori Chrome i idi na stranicu za konvertovanje slika u PDF
        chromeOptions = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": r"C:\Users\User\Downloads\Selenium-Download"}
        chromeOptions.add_experimental_option("prefs", prefs)
        browser = webdriver.Chrome(executable_path=r"chromedriver.exe",
                                   options=chromeOptions)  # chrome_options=chromeOptions
        browser.set_window_size(1200, 1050)
        browser.get('some web page for converting jpg to pdf')
    except:
        print("Neuspesno otvaranje Chrome-a i stranice!")

    try:
        # podesi marginu
        marginElem = browser.find_element_by_name("margin_size")
        marginElem.click()
        marginElem.send_keys(Keys.UP)
        marginElem.click()
    except:
        print("Neuspesno podesavanje margine PDF-a!")

    try:
        # podesi velicinu stranice
        page_sizeElem = browser.find_element_by_name("page_size")
        page_sizeElem.click()
        page_sizeElem.send_keys(Keys.DOWN)
        page_sizeElem.send_keys(Keys.DOWN)
        page_sizeElem.send_keys(Keys.DOWN)
        page_sizeElem.click()
    except:
        print("Neuspesno podesavanje velicine stranice PDF-a!")

    try:
        # podesi orijentaciju stranice
        orientationElem = browser.find_element_by_name("page_orientation")
        orientationElem.click()
        orientationElem.send_keys(Keys.DOWN)
        orientationElem.click()
    except:
        print("Neuspesno podesavanje orijentacije stranice PDF-a!")

    try:
        # klikni na dugme da bi se otvorio prozor za upload
        upload_filesElem = browser.find_element_by_id("select_file_button")
        upload_filesElem.click()
        time.sleep(2)
    except:
        print("Neuspesni klik na dugme za odabir fajlova!")

    try:
        # povezivanje pywinauto sa prozorom open
        app = Application().connect(title_re=".*Open")
        dlg = app.top_window()
    except:
        print("Neuspesno povezivanje PYWINAUTO sa 'Open' dijalogom!")

    try:
        # upisi link do foldera, nakon upisa selektuje sve fajlove i otvara
        dlg.type_keys(type_dir)
        dlg.type_keys("~")
        click(button='left', coords=(300, 300))
        dlg.type_keys('^a')
        dlg.type_keys("~")
        sl = len(os.listdir(path_dir))
        time.sleep(sl * 8.50)  # 4
    except:
        print("Neuspesno otvaranje foldera i selekcija slika!")

    try:
        # konvertovanje u pdf
        convertElem = browser.find_element_by_id("convert_to_pdf_button")
        convertElem.click()
        if sl > 9:
            time.sleep(sl * 3)  # 1.05
        else:
            time.sleep(15)  # 9
    except:
        print("Neuspesan klik na 'Convert to PDF' dugme!")

    try:
        # skidanje pdf-a
        downloadElem = browser.find_element_by_id("download_pdf_button")
        downloadElem.click()
        if sl > 9:
            time.sleep(sl * 3)  # 1.05
        else:
            time.sleep(15)  # 9
    except:
        print("Neuspesan klik na 'DownloadPDF' dugme!")

    try:
        browser.quit()
    except:
        print("Neuspesno vracanje na homepage!")


def compress_pdf(api, pdf_path, out_path):
    t = Compress(api, verify_ssl=True, proxies=False)
    t.add_file(pdf_path)
    t.debug = False
    t.compression_level = 'extreme'  # extreme ili recommended
    t.set_output_folder(out_path)
    t.execute()
    t.download()
    t.delete_current_task()


def merge_pdf(api, pdf_path, out_path):
    pass


##############
# RUN PROGRAM
##############

start = time.time()

# folder koji se proverava tj. broji da bi se znalo da li je zapravo skinut fajl ili nije
# da se ne bi pravila greska
check_folder = r"C:\Users\User\Downloads\Selenium-Download"


# kreiranje PDF-a
for i in range(len(folderi)):
    selenium(folderi[i][0])
    # provera da li je uspesno napravljen PDF
    if i + 1 == len(os.listdir(check_folder)):
        print("Zavrsen je jpg-to-pdf: ", folderi[i][0])
        print("Zavrsen je: ", i + 1, "/", len(folderi))
    else:
        print("Drugi pokusaj da se napravi PDF od: ", folderi[i][0])
        selenium(folderi[i][0])
        if i + 1 == len(os.listdir(check_folder)):
            print("Zavrsen je jpg-to-pdf: ", folderi[i][0])
            print("Zavrsen je: ", i + 1, "/", len(folderi))
        else:
            print("Treci pokusaj da se napravi PDF od: ", folderi[i][0])
            selenium(folderi[i][0])
            if i + 1 == len(os.listdir(check_folder)):
                print("Zavrsen je jpg-to-pdf: ", folderi[i][0])
                print("Zavrsen je: ", i + 1, "/", len(folderi))
            else:
                print("Iskorisceno 3 pokusaja, nije napravljen PDF!")
                print("Nije napravljen: ", folderi[i][0])


# za compress i rename
pdf_folder = r"C:\Users\User\Downloads\Selenium-Download"
out_folder = r"C:\Users\User\Downloads\Selenium-Compress"
pdfs = os.listdir(pdf_folder)
folderi_split = []
new_names = []


# compresovanje PDF-a
for i in range(len(pdfs)):
    try:
        pdf_path = pdf_folder + "\\" + pdfs[i]
        compress_pdf(api_name, pdf_path, out_folder)
        print("Zavrsen je compress: ", pdf_path)
        print("Zavrsen je broj: ", i + 1, "/", len(pdfs))
    except:
        api_cntr += 1
        api_name = api[api_cntr]
        pdf_path = pdf_folder + "\\" + pdfs[i]
        compress_pdf(api_name, pdf_path, out_folder)
        print("ovde je except " + str(i))
        print("Zavrsen je compress: ", pdf_path)
        print("Zavrsen je broj: ", i + 1, "/", len(pdfs))


# odradi rename
for i in range(len(os.listdir(out_folder))):
    os.rename(
        (out_folder + "\\" + os.listdir(out_folder)[i]),
        (out_folder + "\\" + folderi[i][1] + ".pdf")
    )


# odradi merge pdf
merge_paths = []
merge_prefix = []
merge_code = []
merge_list = []

for i in range(len(os.listdir(out_folder))):
    merge_paths.append(out_folder + "\\" + os.listdir(out_folder)[i])
    merge_prefix.append(os.listdir(out_folder)[i].split("-")[0])
    merge_code.append(merge_paths[i].split(
        "\\")[5].split(".")[0].split("-")[1])
    merge_list_temp = merge_prefix[i], merge_code[i], merge_paths[i]
    merge_list.append(merge_list_temp)

key_func = lambda x: x[0]

merge_list_final = []
for key, group in groupby(merge_list, key_func):
    merge_list_final.append(list(group))


f0 = r"C:\Users\User\Documents\Python\Projects\jpg-to-pdf\0-Sa lica mesta.pdf"
f1 = r"C:\Users\User\Documents\Python\Projects\jpg-to-pdf\1-Osnovni.pdf"
f2 = r"C:\Users\User\Documents\Python\Projects\jpg-to-pdf\2-Dopunski.pdf"
f3 = r"C:\Users\User\Documents\Python\Projects\jpg-to-pdf\3-kontrola trap.pdf"
f4 = r"C:\Users\User\Documents\Python\Projects\jpg-to-pdf\4-trap delovi.pdf"
f5 = r"C:\Users\User\Documents\Python\Projects\jpg-to-pdf\5-Stetnik.pdf"
f6 = r"C:\Users\User\Documents\Python\Projects\jpg-to-pdf\6-Nakon popravke.pdf"


# koliko ukupno task merge-a treba da uradi
for i in range(len(merge_list_final)):
    # zapocni task merge
    t = Merge('API_key_here',
              verify_ssl=True, proxies=False)
    # koliko elemenata ima svaki pojedinacni task
    for j in range(len(merge_list_final[i])):
        if int(merge_list_final[i][j][1]) == 0:
            t.add_file(f0)
            t.add_file(merge_list_final[i][j][2])
        elif int(merge_list_final[i][j][1]) == 1:
            t.add_file(f1)
            t.add_file(merge_list_final[i][j][2])
        elif int(merge_list_final[i][j][1]) == 2:
            t.add_file(f2)
            t.add_file(merge_list_final[i][j][2])
        elif int(merge_list_final[i][j][1]) == 3:
            t.add_file(f3)
            t.add_file(merge_list_final[i][j][2])
        elif int(merge_list_final[i][j][1]) == 4:
            t.add_file(f4)
            t.add_file(merge_list_final[i][j][2])
        elif int(merge_list_final[i][j][1]) == 5:
            t.add_file(f5)
            t.add_file(merge_list_final[i][j][2])
        elif int(merge_list_final[i][j][1]) == 6:
            t.add_file(f6)
            t.add_file(merge_list_final[i][j][2])
        elif int(merge_list_final[i][j][1]) == 7:
            t.add_file(f2)
            t.add_file(merge_list_final[i][j][2])
    t.debug = False
    t.set_output_folder(r"C:\Users\User\Downloads\Selenium-Final")
    ime = r"C:\Users\User\Downloads\Selenium-Final\\" + \
        "Fotodokumentacija-" + merge_list_final[i][j][0]
    t.output_filename = ime
    t.execute()
    t.download()
    t.delete_current_task()
    print("Zavrsen je merge: ", "Fotodokumentacija -",
          merge_list_final[i][j][0])
    print("Zavrsen je merge: ", i + 1, "/", len(merge_list_final))


# rename merge fajlova
old_merge_names = os.listdir(r"C:\Users\User\Downloads\Selenium-Final")

for i in range(len(os.listdir(r"C:\Users\User\Downloads\Selenium-Final"))):
    os.rename(
        (r"C:\Users\User\Downloads\Selenium-Final" + "/" + old_merge_names[i]),
        (r"C:\Users\User\Downloads\Selenium-Final" +
         "/" + old_merge_names[i][36::])
    )


# ukupno vreme rada
end = (time.time() - start) / 60
print("END TIME: ", end, " min")
