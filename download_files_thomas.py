# -*- coding: utf-8 -*-
"""
Created on Sun Oct 13 15:37:08 2019

@author: hewi
"""

#### IF error : "ModuleNotFOundError: no module named PyPDF2"
# then uncomment line below (i.e. remove the #):
import pandas as pd
import PyPDF2
from pathlib import Path
import shutil, os
import os.path
import urllib
import glob
import requests

###!!NB!! column with URL's should be called: "Pdf_URL" and the year should be in column named: "Pub_Year"

### File names will be the ID from the ID column (e.g. BR2005.pdf)

########## EDIT HERE:

### specify path to file containing the URLs
list_pth = r'C:\Users\spac-23\PycharmProjects\Multithreading\GRI_2017_2020.xlsx'

###specify Output folder (in this case it moves one folder up and saves in the script output folder)
pth = r'C:\Users\spac-23\PycharmProjects\Multithreading\files'  # fuld loesning

#pth = r'C:\Users\spac-23\PycharmProjects\Multithreading\files\BR50041.pdf'  # test loesning med én fil


###Specify path for existing downloads
dwn_pth = r'C:\Users\spac-23\PycharmProjects\Multithreading\files'

### cheack for files already downloaded
dwn_files = glob.glob(os.path.join(dwn_pth, "*.pdf"))
exist = [os.path.basename(f)[:-4] for f in dwn_files]

###specify the ID column name
ID = "BRnum"

##########

### read in file
df = pd.read_excel(list_pth, sheet_name=0, index_col=ID)

### filter out rows with no URL
non_empty = df.Pdf_URL.notnull() == True
df = df[non_empty]
df2 = df.copy()

# writer = pd.ExcelWriter(pth+'check_3.xlsx', engine='xlsxwriter', options={'strings_to_urls': False})


### filter out rows that have been downloaded
df2 = df2[~df2.index.isin(exist)]
df_filtered = df[df['Pdf_URL'].notnull() | df['Report Html Address'].notnull()]  #AM & AL OR try column name


### loop through dataset, try to download file.
for j in df2.index:
    save_path = str(pth + '\\' + str(j) + '.pdf')
    try:
        primary = df_filtered.at[j, 'Pdf_URL']
        secondary = df_filtered.at[j, 'Pdf_URL']
        response = requests.get(primary)
        print()
        if 200 <= response.status_code < 300:
            with open(save_path, 'wb') as file:
                file.write(response.content) # write bytes in stead of string
        else:
            response = requests.get(secondary)
            print()
            if 200 <= response.status_code < 300:
                with open(save_path, 'wb') as file:
                    file.write(response.content)  # write bytes in stead of string

    # pdfFileObj = open(savefile, 'rb')
    # creating a pdf reader object
    # pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    # with open(savefile, 'rb') as pdfFileObj:
    # pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    # if pdfReader.numPages > 0:
    # df2.at[j, 'pdf_downloaded'] = "yes"
    # else:
    # df2.at[j, 'pdf_downloaded'] = "file_error"

    # except Exception as e:
    # df2.at[j, 'pdf_downloaded'] = str(e)
    # print(str(str(j)+" " + str(e)))
    # else:
    # df2.at[j, 'pdf_downloaded'] = "404"
    # print("not a file")

    except (requests.HTTPError, ConnectionResetError, Exception) as e:
        df2.at[j, "error"] = str(e)

# df2.to_excel(writer, sheet_name="dwn")
# writer.save()
# writer.close()
