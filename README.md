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

# How does app work

### App works in a following manner:

 1. In the Python script, you need to define API_KEY, location of folders with images which you want to convert to PDF and which ones would you like to combine (merge together).

2. Attempt to read the data from PDF file into python script as a PDF class (using pdfplumber) and it was needed to check if PDF file was created properly using pdf printers (you could extract data) because sometimes, people created a PDF file that was actually an image (i.e. JPG) inside PDF so data extraction was not possible.

3. If data could be extracted from PDF file, the file was read and one large string was created

4. String was then separated by rows (PDF had structured data, every single one
had the same structure, just different data)

5. Data needed was then selected from it's corresponding row and placed inside
a variable

6. Because the variables were in latin letters (and some documents or template were needed to be in cyrillic letters) they had to be transliterated into cyrillic letters using a custom created module

7. After transliteration, data was exported into a new Excel file - saob_data(.xlsx). 

### **Note 1**
It was easier to create a new excel file that was then imported into a new template excel document or DB then it was to directly export it into excel template (because openpyxl was deleting pictures from excel files at that time) and most of the templates had logos or other images embeded in them. Second reason, because then the python script would be too coupled with the files and applications used (like 5 separate people, with little to none IT knowledge, worked in DB and/or creating new Excel documents, doing office work at the same time and that was a bad idea.)

### **Note 2**
Mainloop of GUI has weird if statment for running because the app was placed on 
the server-pc first and all users would connect to it (file server) and use it. 
It was easier for bug fixes and improvments to have it in one place then on every
PC. But later, the possibility to run it on every pc was added, per request.

### **Disclaimer**
Provided sample PDF for testing purposes doesn't represent a real vehicles registration ID. The data has been changed to protect sensitive info, swaped with dummy data. Any similarity is entirely coincidental.
