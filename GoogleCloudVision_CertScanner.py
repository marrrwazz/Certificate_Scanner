############################################################################
#   Written By: Marcin Waz
#   Written On: 12/06/2017
#   Takes all jpeg files in folder specificed below and uses google cloud
#   vision's API for text detection to extract info needed when form001 is
#   updated.
#   No limit to how many images can be processed at once, but keep in mind
#   how much time is required for each image.
#   Path names can be changed and probabily will be changed.
#   If an error pops up, well the image processing isnt exact each time. 
############################################################################
############################################################################
############################################################################
############################################################################
############################################################################
from google.cloud import vision
from google.cloud.vision import types
from oauth2client.client import GoogleCredentials
from oauth2client.contrib import gce
from googleapiclient.discovery import build
from collections import OrderedDict
import httplib2
import io
import pandas as pd
from os import listdir
import xlsxwriter

#Authenticate Google Cloud API
credentials = gce.AppAssertionCredentials(
    scope = 'https://www.googleapis.com/auth/devstorage.red_write')
http = credentials.authorize(httplib2.Http())
service = build('compute', 'v1', credentials=credentials)
client = vision.ImageAnnotatorClient()

def extract_jpgs(filenames):
    """Extract jpegs from folder"""
    just_jpg = []
    for name in filenames:
        if "jpg" in name:
            just_jpg.append(name)
    return just_jpg

#   This is the first of two locations where you need to change the path
#   to whatever images it is that you are using.
#   Must be JPEG's because of Google.
filenames = listdir(r'C:\Users\H-63\Pictures\ricelake')
filenames = extract_jpgs(filenames)


form_info ={
        "Current SN": [],
        "Traceable Report Number": [],
        "Weight": [],
        "Unit": [],
        "Category": [],
        "Description": [],
        "CLS": [],
        "Certificate Date": [],
        "Cert. Life": [],
        "Certificate Due": [],
        "Cal Service Provider": [],
        "Accredited By": [],
        "UNC": [],
        "UNC-unit": [],
        "TOL": [],
        "TOL-unit": [],
        }

def form_001_info(list_of_detected_words):
    """Constructs scanned info into dict"""
    document_text = create_document(list_of_detected_words)
    document_text.lower()
    form_info['Accredited By'].append("NVLAP")
    if "Description of Weights" in document_text:
        splitdoc = document_text.partition("Description of Weights:")
        description = splitdoc[2].partition("Nominal")
        description = description[0]
        serial_number = description.partition("S/N")
        weight_description = serial_number[0]
        cValue = serial_number[0]
        cValue = cValue.partition("Class")
        #print(cValue)
        cValue = cValue[2]
        serial_number = serial_number[2].partition(" ")
        serial_number = serial_number[2].partition(" ")
        if serial_number[0] != "":
            form_info["Current SN"].append(serial_number[0])
        else:
            form_info["Current SN"].append(" ")
        if cValue != "":  
            form_info["CLS"].append(cValue)
        else:
            form_info["CLS"].append(" ")
    if "Traceable Certificate Number:" in document_text:
        report_number = document_text.partition("Traceable Certificate Number")
        report_number = report_number[2]
        report_number = report_number.partition(" ")
        report_number = report_number[2]
        report_number = report_number.partition(" ")
        if report_number[0]!="":
            if "Contractor" not in report_number[0]:
                form_info["Traceable Report Number"].append(report_number[0])
            elif "Contractor" in report_number[0]:
                report_number = document_text.partition("Contractor:")
                report_number = report_number[2]
                report_number.lstrip()
                report_number = report_number.partition(" ")
                report_number = report_number[2]
                report_number = report_number.partition(" ")
                report_number = report_number[0]
                #print(report_number)
                if report_number != "":
                    form_info["Traceable Report Number"].append(report_number)
                else:
                    form_info["Traceable Report Number"].append(" ")
        else:
            form_info["Traceable Report Number"].append(" ")

    if "Date Calibrated:" in document_text:
        date_cal = document_text.partition("Date Calibrated: ")
        date_cal = date_cal[2]
        date_cal = date_cal.partition("Recall Date: ")
        date_due = date_cal[2]
        date_due = date_due.partition("Temperature")
        date_due = date_due[0]
        date_cal = date_cal[0]
        try:
            num_date_due = int(date_due[7:])
        except:
            num_date_due = 0
        try:
            num_date_cal = int(date_cal[7:])
        except:
            num_date_cal = 0
        form_info["Certificate Date"].append(date_cal)
        form_info["Certificate Due"].append(date_due)
        form_info["Cert. Life"].append(num_date_due - num_date_cal)
        
    if "RICE LAKE" or "Rice Lake" in document.text:
        form_info["Cal Service Provider"].append("Rice Lake")
        
    if "weight" or "Weight" in document.text:
        form_info["Category"].append("Weight/Hanger")
        #weight_description = serial_number[0]
        weight_description = weight_description.partition(",")
        weight_description = weight_description[0]
        weight_description = weight_description.partition(" kg ")
        if weight_description[2] == "":
            try:
                weight_description = weight_description[0].partition(" g ")
            except TypeError:
                pass;
        if weight_description[2] == "":
            try:
                weight_description = weight_description[0].partition(" lb ")
            except TypeError:
                pass;
        if weight_description[2] == "":
            try:
                weight_description = weight_description[0].partition(" oz ")
            except TypeError:
                pass;
        if weight_description[2] != "":
            form_info["Unit"].append(weight_description[1])
            form_info["Description"].append(weight_description[2])
            form_info["Weight"].append(weight_description[0])
        else:
            form_info["Unit"].append(" ")
            form_info["Description"].append(" ")
            form_info["Weight"].append("")            
    if " 2 " in document_text:
        TOL = document_text.partition(" 2 ")
        UNC = TOL[0]
        TOL = TOL[2]
        TOL = TOL.lstrip()
        TOL = TOL.partition(" ")
        TOL = TOL[0]
        UNC = UNC.split()
        UNC = UNC[len(UNC)-1]
        if UNC != "":
            form_info['UNC'].append(UNC)
        else:
            form_info['UNC'].append(" ")
        if TOL!= "":
            form_info['TOL'].append(TOL)
        else:
            form_info['TOL'].append(" ")
            
    if "(mg)" in document_text:
        form_info['TOL-unit'].append("mg")
        form_info['UNC-unit'].append("mg")
    else:
        form_info['TOL-unit'].append(" ")
        form_info['UNC-unit'].append(" ")        
            
    return form_info
    
def create_document(list_of_words):
    """Creates 'wall of text' from identified text"""
    document_text = ""
    for word in list_of_words:
        document_text = document_text + " " + word         
    return document_text

def detect_document(path):
    """Detects text in document"""
    client = vision.ImageAnnotatorClient()

    with io.open(path, 'rb') as image_file:
        content = image_file.read()

    image = types.Image(content=content)

    response = client.document_text_detection(image=image)
    document = response.full_text_annotation

    for page in document.pages:
        for block in page.blocks:
            block_words = []
            for paragraph in block.paragraphs:
                block_words.extend(paragraph.words)

            block_symbols = []
            for word in block_words:
                block_symbols.extend(word.symbols)

            block_text = ''
            for symbol in block_symbols:
                block_text = block_text + symbol.text
                
    #Sift through output to find all of the words and phrases
    list_of_detected_words = []
    temp_phrase = []
    for word in document.text:
        if word != '\n':
            temp_phrase.append(word)
        elif word == '\n':
            temp_word = ""
            for letter in temp_phrase:
                temp_word = temp_word + letter
            list_of_detected_words.append(temp_word)
            temp_phrase = []
            
    #Write the list of all detected words and phrases to csv file       
    textfile = r'list_of_detected_words.csv'    
    with io.open(textfile,'w') as text:
        for word in list_of_detected_words:
            try:
                text.write(word +  " ")
            except:
                word = " "         
        text.close()

    #Write to list of necessary information to excel file
    form_info = form_001_info(list_of_detected_words)
    try:
        writer = pd.ExcelWriter('list_of_detected_words.xlsx', engine='xlsxwriter')
        list_of_detected_words = pd.DataFrame.from_records(form_info)
        list_of_detected_words.to_excel(writer, "Sheet1",startrow=1, columns = ["Current SN",
        "Traceable Report Number", "Weight", "Unit", "Category","Description",
        "CLS","Certificate Date", "Cert. Life", "Certificate Due","Accredited By",
        "Cal Service Provider","UNC", "UNC-unit","TOL","TOL-unit"])
    except ValueError:
        print(form_info)

#
for file in filenames:
    path = r'C:\Users\H-63\Pictures\ricelake\\'
    file = path + file
    detect_document(file)
    
print("You're Done!")
