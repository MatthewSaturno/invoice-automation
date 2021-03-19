import PyPDF2
import csv
import re
import os

invoice_list = []

#Create list of files in the directory
dirName = "INSERT_FOLDER_WITH_INVOICES_TO_LOOP_HERE"
listOfFiles = list()
for (dirpath, dirnames, filenames) in os.walk(dirName):
    listOfFiles += [os.path.join(dirpath, file) for file in filenames]

data = []
columns = 7
    
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
    invoice_list.append(page_one_text)
    
    #Extract Invoice Total
    f_list=[]
    f1 = re.finditer("SALES TAX TOTAL",page_one_text)
    for item in f1:
        f_list.append(item.end())

    f2_list=[]
    f2 = re.finditer("Subject to",page_one_text)
    for item in f2:
        f2_list.append(item.start())    

    try:
        match_list = page_one_text[f_list[-1]+1:f2_list[-1]-1].split(" ")

        subtotal = match_list[0]
        sales_tax = match_list[1]
        total = match_list[2]
    except:
        subtotal = "Not found"
        sales_tax = "Not found"
        total = "Not found"


    #Extract Customer
    search_list = []
    try:
        x = re.search("SHIP TO:",page_one_text)
        y = re.search("PO",page_one_text[x.end():])
        z = re.search(r'\d+',page_one_text[x.end():])
        w = re.search("c/o",page_one_text[x.end():])

        search_list.append(y)
        search_list.append(z)
        search_list.append(w)

        m=[]
        for item in search_list:
            if item is not None:
                m.append(x.end()+item.start())

        m.sort(reverse=False)        
        customer = page_one_text[(x.end()+1):(m[0]-1)]
    except:
        customer = "Not found"


    #Extract Invoice Number
    try:
        x = re.search('I N V O I C E', page_one_text)
        y = re.search(r'\d+' , page_one_text[x.end():])
    except:
        x = re.search('INVOICE / FACTURE', page_one_text)
        y = re.search(r'\d+' , page_one_text[x.end():])

    try:        
        invoice_number = page_one_text[y.start()+x.end():y.end()+x.end()]
    except:
        invoice_number = "Not found"
      

    #Extract Invoice Date
    import datetime

    x = re.search(r'\d+/+\d+/+\d+',page_one_text) 

    try:
        if page_one_text[x.end()+1] == 'D':
            y = page_one_text[(x.start()-2)]
        else:
            y = page_one_text[x.end()+10]

        invoice_date = x.group()+'02'+y

        #date format
        ds = invoice_date.split("/")
        invoice_date = datetime.date(int(ds[2]),int(ds[0]),int(ds[1]))
    except:
        invoice_date = "Not found"
    
    #Extract Currency
    x = re.search('CURRENCY BRANCH PLANT', page_one_text)
    try:
        x1 = page_one_text[x.end():x.end()+10]
        if 'USD' in x1:
            currency = 'USD'
        elif 'CAD' in x1:
            currency = 'CAD'
        else:
            currency = 'foreign currency' 
    except AttributeError:
        currency = 'Not found'

      
    data.append(invoice_number)
    data.append(invoice_date)
    data.append(customer)
    data.append(subtotal)
    data.append(sales_tax)
    data.append(total)
    data.append(currency)
    
    
#Create Data Set
data_set =[]
y = 0
rows = int(len(data)/columns)

#Split list and create data set
for num in range(0,rows):
    data_set.append(data[y:y+columns])
    y+= columns

    
f2 = open("INSERT_DESTINATION_FILE_PATH_HERE",'w',newline='')
csv_writer = csv.writer(f2, delimiter=',')
csv_writer.writerow(['Invoice #','Invoice Date','Customer','Subtotal','Sales Tax','Total','Currency'])

for row in data_set:
    csv_writer.writerow([row[0],row[1],row[2],row[3],row[4],row[5],row[6]])

#Close PDF and csv
f2.close()