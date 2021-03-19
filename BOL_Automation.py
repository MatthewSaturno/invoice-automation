import PyPDF2
import csv
import re
import os
import datetime

bol_list = []

#Create list of files in the directory
dirName = "INSERT_LIST_OF_BOLs_TO_LOOP_THROUGH_HERE"
listOfFiles = list()
for (dirpath, dirnames, filenames) in os.walk(dirName):
    listOfFiles += [os.path.join(dirpath, file) for file in filenames]

data = []
columns = 5
    
#Loop through the files    
for file in listOfFiles:

    #Open PDF
    f = open(file,'rb')
    pdf_reader = PyPDF2.PdfFileReader(f)
    page_one_text = []

    #Loop through the pages and combine text
    for page in list(range(0,pdf_reader.getNumPages())):
        page_one = pdf_reader.getPage(page)
        page_one_text.append(page_one.extractText())
    
    page_one_text = " ".join(page_one_text)
    bol_list.append(page_one_text)
    
#Extract from the bol csv
data_file = input("\nPlease paste BOL file name here \n")
data_data = open(data_file,encoding='utf-8-sig')
csv_data = csv.reader(data_data)
data_lines = list(csv_data)


#Go through each bol and extract data
bol_data = []
for item in bol_list:
    
    #Extract BOL and Shipping Doc
    for num in data_lines:
        vlookup = []
        if num[0] in item and num[1] in item:
            bol_data.append(num[0])
            bol_data.append(num[1])
        else:
            vlookup.append(1)

        if len(vlookup) == len(data_lines):
            bol_data.append("Not found")
            bol_data.append("Not found")
    
    #Extract ship date
    ship_date = re.search(r'\d+/+\d+/+\d+',item)
    ship_date_word = re.search('Ship Date',item)
    a = ship_date.end()
    b = ship_date_word.end()

    if item[a:a+1] == '\n':
        ship_date = re.search(r'\d+/+\d+/+\d+',item).group()+'20'
    elif item[b:b+1] == ' ':
        ship_date = re.search(r'\d+/+\d+/+\d+',item[b:]).group()+'20'
    else:
        ship_date = re.search(r'\d+/+\d+/+\d+',item).group()+'20'

    #date format
    ds = ship_date.split("/")
    ship_date = datetime.date(int(ds[2]),int(ds[0]),int(ds[1]))
    bol_data.append(ship_date)

    
    #Extact weight 
    my_list = []
    for x in re.finditer(r'\d+?(\,\d+)?\d+(\.\d{4})',item):
        my_list.append(x.group())
    my_list = my_list[-3:]

    try:
        if round(float(my_list[2].replace(',',''))+float(my_list[1].replace(',',''))) == round(float(my_list[0].replace(',',''))):
            bol_data.append(my_list[2])
        elif round(float(my_list[2].replace(',',''))-float(my_list[1].replace(',',''))) == round(float(my_list[0].replace(',',''))):
            bol_data.append(my_list[0])
        else:
            bol_data.append(my_list[2])
    except:
        bol_data.append(my_list[-1])
        
    #Extract metric
    r = re.compile(r'\bKG\b | \bLB\b', flags=re.I | re.X)
    match = r.findall(item)
    try:
        bol_data.append(match[-1])
    except:
        bol_data.append('Uknown Metric')
        
#Create Data Set
data_set =[]
y = 0
rows = int(len(bol_data)/columns)

#Split list and create data set
for num in range(0,rows):
    data_set.append(bol_data[y:y+columns])
    y+= columns
  
f2 = open('INSERT_DESTINATION_FILE_PATH_HERE','w',newline='')
csv_writer = csv.writer(f2, delimiter=',')
csv_writer.writerow(['Sales Order #','BOL #','Ship Date','Weight','Unit'])

for row in data_set:
    csv_writer.writerow([row[0],row[1],row[2],row[3],row[4]])

#Close PDF and csv
f2.close()
    

