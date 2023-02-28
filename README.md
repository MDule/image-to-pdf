# Image to PDF (via iLovePDF API)


An automated script that solves a specific problem:

>1. Create a PDF file with correct orientation and no margins using either:
>     - folder with images; 
>     - multiple folders with images that are then merged together.
>2. Compress PDF file so that it's small enough to be sent as an email attachment.

Creation of PDF file was achieved using third party website at [URL](https://www.convert-jpg-to-pdf.net/). Note: this particular website was chosen for a few reasons but the main two reasons were that the upload quota was around 500MB which was enough and the resulting PDF file could be compressed (which was not the case with other website or programs, even python based scripts using PIL and PDF converters). 

Compression of PDF files was achieved using [iLovePDF](https://www.ilovepdf.com/) API.

#
> By: Dušan Miletić 
>
> Date created: 1. november 2020.
#

# Dependencies

Project uses the following dependencies (which can be installed via pip install):

> **selenium** - for browser manipulation (I used Google Chrome)

> **pywinauto** - for manipulation of mouse and keyboard in Windows (for managing of  file dialogs)

> **pylovepdf** - library for iLovePDF API (at one time it was listed as library (beta) on their website back in 2020.)

Script is made for Windows OS. 

You also need to:
> 1. Sign up and register on [iLovePDF](https://developer.ilovepdf.com/) to get your API key.
> 2. Download Chrome web driver (needed for Selenium) [here](https://chromedriver.chromium.org/downloads). Make sure to download the version that matches installed Chrome version.

#

# Running the app

You can run the app in two ways:

> Using Python (if you have Python and dependencies installed)
> 
> 1. ``git clone https://github.com/MDule/Image-to-PDF.git``
> 2. ``pip install -r requirements.txt``
> 3. ``[Windows] python img2pdf.py OR py img2pdf.py`` 

#

# How does the app work

## In short, the scripts works like this:

> 1. Download Chrome Web Driver and place it in root folder of the project
> 2. Enter API key in the script (my API key is .env file, you need to supply your own)
> 3. Define paths to folders with images in format -> ``[path, "file_id-part_id"]``
>    1. path: absolute path to folder, string literal i.e. r"C:\Images\Folder1"
>    2. file_id: define id for the file [int]
>    3. part_id: if you have multiple folders with images but you want to merge them together, then you need to enter part_id as [int], files will be grouped by file_id
>    - i.e. 
> ```
>       folders=[
>           [r"C:\Images\Folder1", "1-1"]
>           [r"C:\Images\Folder2\part1", "2-1"],
>           [r"C:\Images\Folder2\part1", "2-2"],
>       ]
>```
> 4. Run the script and Selenium will take over, it will open the website via Chrome, upload files, set PDF orientation, margin size, paper size and download the files (files will be downloaded in the root folder, inside folder "1-step-Download")
> 5. Then the script will upload PDFs from "1-step-Download" folder using iLovePDF API and compress them to minimum file size (files will be downloaded to "2-step-Compressed" folder)
> 6. Script will then rename the files inside "2-step-Compressed" folder
> 7. Group them into a list
> 8. Upload and merge them via iLovePDF API
> 9. Download them to "3-step-Final" folder

#

### **Note 1**
This script was made to solve a specific problem, to create PDF files from images that that could be sent via e-mail. Images had size from 5-15MB per image and on average each folder had 30-50 images (~300-500MB in total without compression) and the number of folders would range from 10-20 (without counting the subfolders) every week.

The PDFs could be created using desktop app but it is a manual, boring and repetitive task that could be automated. And given the number of folders, it would take a lot of time to be completed manually, hence the automation's aim was to save time and remove repetitive work.

Resulting PDF had 'page or pages' which indicated from which folder the images came and which ones were grouped (business requirement).

### **Note 2**
The script was meant to be run when everyone was out of office, completely automated and was written poorly, when I was starting to learn Python and wanted to use what I learned to solve current problems in my workplace and practiced what I learned and as a bonus save time for productive work. It was used only by myself for specific use-case for I didn't waste time making the code look pretty, I just needed it to work. 

