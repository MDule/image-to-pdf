from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from pywinauto.application import Application
from pywinauto.mouse import click
from pylovepdf.tools.compress import Compress
from pylovepdf.tools.merge import Merge
from dotenv import load_dotenv
import os
import time
import sys
from itertools import groupby

###############
# Prerequisites
###############

load_dotenv()

# define API key here
API_KEY = None

# if user didn't set API key above then check if the .env file exists with API key
# if it exists, use the API key from .env
if os.getenv("API"):
    API_KEY = os.getenv("API")

# if API_KEY is still None, raise exception
if not API_KEY:
    raise Exception("Error! No API key defined. Please enter API key and try again.")


##############
# FOLDERS
##############

# folders with images needed to be converted to PDF and/or grouped then compressed
# [folder_path, folder_code - cover_page_code]
# use absolute paths and string literals (prefix r'string'), for demonstation purposes, f-string is used with root dir
# as example files reside there
dir_path = os.getcwd()
folderi = [
    [f"{dir_path}\\imgs\\imgs1\\", "1-1"],
    [f"{dir_path}\\imgs\\imgs2\\part1\\", "2-1"],
    [f"{dir_path}\\imgs\\imgs2\\part2\\", "2-2"],
]


###########
# FUNCTIONS
###########


def selenium(path_dir):
    # construct folder path (from paths declared above) for pywinauto
    type_dir = path_dir.replace(" ", "{SPACE}")

    try:
        # open Chrome and go to website for converting images to pdf
        # set chrome options
        chromeOptions = webdriver.ChromeOptions()
        prefs = {"download.default_directory": f"{dir_path}\\1-step-Download"}
        chromeOptions.add_experimental_option("prefs", prefs)
        browser = webdriver.Chrome(
            # executable_path="chromedriver.exe",
            service=Service("chromedriver.exe"),
            options=chromeOptions,
        )
        browser.set_window_size(1200, 1050)
        browser.get("https://www.convert-jpg-to-pdf.net/")
    except:
        print("Error opening Chrome and the website!")
        sys.exit("Error opening Chrome and the website!")

    try:
        # set margin via Selenium
        marginElem = browser.find_element(By.NAME, "margin_size")
        marginElem.click()
        marginElem.send_keys(Keys.UP)
        marginElem.click()
    except:
        print("Failed to set margin_size for PDF!")
        sys.exit("Exited - Failed to set margin_size for PDF!")

    try:
        # set page size
        page_sizeElem = browser.find_element(By.NAME, "page_size")
        page_sizeElem.click()
        page_sizeElem.send_keys(Keys.DOWN)
        page_sizeElem.send_keys(Keys.DOWN)
        page_sizeElem.send_keys(Keys.DOWN)
        page_sizeElem.click()
    except:
        print("Failed to set page_size for PDF!")
        sys.exit("Exited - Failed to set page_size for PDF!")

    try:
        # set page orientation
        orientationElem = browser.find_element(By.NAME, "page_orientation")
        orientationElem.click()
        orientationElem.send_keys(Keys.DOWN)
        orientationElem.click()
    except:
        print("Failed to set page_orientation for PDF!")
        sys.exit("Exited - Failed to set page_orientation for PDF!")

    try:
        # click the button for file open dialog - pywinauto
        upload_filesElem = browser.find_element(By.ID, "select_file_button")
        upload_filesElem.click()
        time.sleep(2)
    except:
        print("Failed to open file selection dialog!")
        sys.exit("Exited - Failed to open file selection dialog!")

    try:
        # connecting pywinauto with file open dialog
        app = Application().connect(title_re=".*Open")
        dlg = app.top_window()
    except:
        print("Failed to connect PYWINAUTO with 'File Open' dialog!")
        sys.exit('Exited - Failed to connect PYWINAUTO with "File Open" dialog!')

    try:
        # pywinauto types path to folder, selects all files i opens
        dlg.type_keys(type_dir)
        dlg.type_keys("~")
        click(button="left", coords=(300, 300))
        dlg.type_keys("^a")
        dlg.type_keys("~")
        sl = len(os.listdir(path_dir))
        time.sleep(sl * 8.50)  # 4
    except:
        print("Failed to open folder and select images!")
        sys.exit("Exited - Failed to open folder and select images!")

    try:
        # convert to PDF
        convertElem = browser.find_element(By.ID, "convert_to_pdf_button")
        convertElem.click()
        if sl > 9:
            # wait for images to upload, depending on upload speed
            # 3 sec per image, can be lowered
            time.sleep(sl * 3)
        else:
            # if there are less than 9 images, wait 15 sec
            time.sleep(15)
    except:
        print("Failed click on 'Convert to PDF' button!")
        sys.exit("Exited - Failed click on 'Convert to PDF' button!")
    try:
        # downloading of pdf
        downloadElem = browser.find_element(By.ID, "download_pdf_button")
        downloadElem.click()
        if sl > 9:
            # wait for download, depends on download speed
            # 3s per image, can be lowered
            time.sleep(sl * 3)
        else:
            # if there are less than 9 images, wait 15 sec
            time.sleep(15)
    except:
        print("Failed click on 'Download PDF' button!")
        sys.exit("Exited - Failed click on 'Download PDF' button!")

    try:
        # close browser
        browser.quit()
    except:
        print("Failed to close the browser!")
        sys.exit("Exited - Failed to close the browser!")


# function for compressing PDF via API
def compress_pdf(api, pdf_path, out_path):
    t = Compress(api, verify_ssl=True, proxies=False)
    t.add_file(pdf_path)
    t.debug = False
    t.compression_level = "extreme"  # extreme or recommended
    t.set_output_folder(out_path)
    t.execute()
    t.download()
    t.delete_current_task()


##############
# RUN PROGRAM
##############

start = time.time()

# counts items in folder to make sure file was downloaded to avoid error
check_folder = f"{dir_path}\\1-step-Download"


# create PDF
for i in range(len(folderi)):
    selenium(folderi[i][0])
    # check if PDF is made successfully
    if i + 1 == len(os.listdir(check_folder)):
        print("Done jpg-to-pdf: ", folderi[i][0])
        print("Done (total): ", i + 1, "/", len(folderi))
    else:
        print("Second attempt to create PDF from: ", folderi[i][0])
        selenium(folderi[i][0])
        if i + 1 == len(os.listdir(check_folder)):
            print("Done jpg-to-pdf: ", folderi[i][0])
            print("Done (total): ", i + 1, "/", len(folderi))
        else:
            print("Third attempt to create PDF from: ", folderi[i][0])
            selenium(folderi[i][0])
            if i + 1 == len(os.listdir(check_folder)):
                print("Done jpg-to-pdf: ", folderi[i][0])
                print("Done (total): ", i + 1, "/", len(folderi))
            else:
                print("3/3 attempts failed, PDF not created!")
                print("Failed: ", folderi[i][0])


# folder where pdfs are
pdf_folder = f"{dir_path}\\1-step-Download"
# folder where compressed pdfs will be
out_folder = f"{dir_path}\\2-step-Compressed"
pdfs = os.listdir(pdf_folder)
folderi_split = []
new_names = []


# compressing PDFs
for i in range(len(pdfs)):
    try:
        pdf_path = pdf_folder + "\\" + pdfs[i]
        compress_pdf(API_KEY, pdf_path, out_folder)
        print("Finished compressing: ", pdf_path)
        print("Finished (total): ", i + 1, "/", len(pdfs))
    except:
        pdf_path = pdf_folder + "\\" + pdfs[i]
        compress_pdf(API_KEY, pdf_path, out_folder)
        print("except here " + str(i))
        print("Finished compressing: ", pdf_path)
        print("Finished (total): ", i + 1, "/", len(pdfs))


# rename files
for i in range(len(os.listdir(out_folder))):
    os.rename(
        (out_folder + "\\" + os.listdir(out_folder)[i]),
        (out_folder + "\\" + folderi[i][1] + ".pdf"),
    )


# merge pdfs
merge_paths = []
merge_prefix = []
merge_code = []
merge_list = []

for i in range(len(os.listdir(out_folder))):
    merge_paths.append(out_folder + "\\" + os.listdir(out_folder)[i])
    merge_prefix.append(os.listdir(out_folder)[i].split("-")[0])
    merge_code.append(merge_paths[i].split("\\")[-1].split(".")[0].split("-")[1])
    merge_list_temp = merge_prefix[i], merge_code[i], merge_paths[i]
    merge_list.append(merge_list_temp)

key_func = lambda x: x[0]

merge_list_final = []
for key, group in groupby(merge_list, key_func):
    merge_list_final.append(list(group))

# for multipart pdfs, set which pdf files would you like to merge and their paths
# this are cover pages you could use
f0 = f"{dir_path}\\cover_pages\\0-part.pdf"
f1 = f"{dir_path}\\cover_pages\\1-part.pdf"
f2 = f"{dir_path}\\cover_pages\\2-part.pdf"
f3 = f"{dir_path}\\cover_pages\\3-part.pdf"
f4 = f"{dir_path}\\cover_pages\\4-part.pdf"
f5 = f"{dir_path}\\cover_pages\\5-part.pdf"
f6 = f"{dir_path}\\cover_pages\\6-part.pdf"


# how many pdf merges need to be done
for i in range(len(merge_list_final)):
    # start merging
    t = Merge(
        API_KEY,
        verify_ssl=True,
        proxies=False,
    )
    # how many elements for each individual task
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
    t.set_output_folder(f"{dir_path}\\3-step-Final")
    ime = (
        f"{dir_path}\\3-step-Final\\" + "Fotodokumentacija-" + merge_list_final[i][j][0]
    )
    t.output_filename = ime
    t.execute()
    t.download()
    t.delete_current_task()
    print("Finished merging: ", "Fotodokumentacija -", merge_list_final[i][j][0])
    print("Finished (total): ", i + 1, "/", len(merge_list_final))


# rename merged files
old_merge_names = os.listdir(f"{dir_path}\\3-step-Final")

for i in range(len(os.listdir(f"{dir_path}\\3-step-Final"))):
    os.rename(
        (f"{dir_path}\\3-step-Final" + "/" + old_merge_names[i]),
        (f"{dir_path}\\3-step-Final" + "/" + old_merge_names[i][-23:]),
    )


# total time
end = (time.time() - start) / 60
print("END TIME: ", end, " min")
