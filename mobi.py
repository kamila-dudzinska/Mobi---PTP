# -*- coding: utf-8 -*-
"""
Created on Wed Mar 11 23:40:59 2026
@author: Kamila Dudzińska

''' Program przechodzi przez plik excel, sprawdza status PO - jeśli jest 
'ordered', ale delivery date jest w przeszłości, to pogram wysyła maila do 
użytkownika,aby wyjasnił kwestię z kupcem. 

W procuremencie, aby faktura była zaksięgowana musi być spełniony 3 way match, 
czyli zamówienie = przyjęcie = faktura. Gdy brakuje przyjęcia (GR) faktura 
nie może zostać zapłacona, Program dodatkowooblicza statystyki i wysyła maila 
do administratora orazdodaje załącznik w formacie pdf. Poprzez automatyzację 
wysyłki maili do kupców - program rozwiązuje realny problem biznesowy 
często spotykany w dużych centrach usług wspólnych.

Program pracuje na pliku procurement_dataset1.csv, który został przez mnie 
stworzony na podstawie skryptu (napisanego przeze mnie) w pythoie i służacego 
do generowania mockowych danych procurementowych, z zachowaniem realnej 
inżynierii danychgenerowanych w systemie SAP Ariba. 
"""

# IMPORT MODULES I
import os
import time
from datetime import datetime, timedelta
import pandas as pd
import win32com.client
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from mobi_function import divide_z 


# LOAD DATA II
file_path = "./mobi_ptp/procurement_mock_2500.csv"
df = pd.read_csv(file_path)

# DATA CLEANING - PREPROCESSING
df['Requester Name'] = df['Requester Name'].astype(str)
df['PO Number'] = df['PO Number'].astype(str)
df['Delivery Date'] = df['Delivery Date'].astype(str).str.strip()
df['Delivery Date'] = pd.to_datetime(df['Delivery Date'], 
                                     errors='coerce',     
                                     format="%Y-%m-%d")
df['Amount'] = df['Amount'].astype(int)

# DATA PREPARATION FOR STATISTICS III
ordered_count = 0
received_count = 0
emails_sent = 0
suma = ordered_count + received_count

# wartoci procentowe 
a = ordered_count
b = suma                    #ordered_count + received_count

# daty
today = datetime.today()
new_date = today - timedelta(days=3)
print(type(new_date))

# OUTLOOK PART IV
outlook =  win32com.client.Dispatch("Outlook.Application")

# EXCEL & OUTLOOK PART V
#iteracja po wierszach excela
for index, row in df.iterrows():
    
    status = row['Order Status']
    mail = row['Requester Mail']
    name = row['Requester Name']
    po_number = row['PO Number']
    delivery_date = row['Delivery Date']
    amount =  row['Amount']
    
    if status == "ordered" and delivery_date < new_date:
        ordered_count += 1
        subject = "Please check GR Missing"
        body = f"""
        Hello {name},
        
        The PO {po_number}  is in status ordered, however the delivery date 
        {delivery_date} is in the past. 
        Could you be so kind and check it? The order amount is {amount}.
        
        Thank you in advance!
        
        Kind regards,
        Admin
        """
        
        # tworzenie maila
        mail = outlook.CreateItem(0)
        mail.To = mail
        mail.Subject = subject
        mail.Body = body
        emails_sent += 1
        
        #wygenrowanie maila* 
        mail.Display()
        
        print(f'Email sent/created* to {mail}')
        
    elif status == "received":
        
        received_count += 1
        
    else:
        
        ordered_count +=1
        
    
        
    suma +=1 

# obliczenia do statystyk --> (divide_z) 
a = ordered_count
b = received_count  
percentage_ordered = divide_z(a, suma) *100
percentage_received = divide_z(b, suma) *100
        
print(ordered_count, received_count, suma)
              
# ADMIN PART 
#tworzenie nowego pliku pdf
pdf_file = r"C:\Users\lila_\Desktop\GitHub\mobi_ptp\report.pdf"

c=canvas.Canvas(pdf_file, pagesize=A4)

# tytuł
y=720
c.setFont("Helvetica-Bold", 14)
c.setFillColorRGB(0,0,50)
c.drawString(100, y, "Report of PO status")
c.setLineWidth(1)
c.line(100, y-2, 100+250, y-3)

#reszta raportu
text = c.beginText(100, y-50)
c.setFont("Helvetica", 12)
c.setFillColorRGB(0,0,0)
text.setLeading(30)


text.textLine(f"PO with the status 'ordered' {ordered_count}.")
text.textLine(f"PO with the status 'recieved' {received_count}")
text.textLine(f"Python has sent {emails_sent} email")
text.textLine(f"W raporcie mamy {suma} PO")
text.textLine(f"w tym {ordered_count} czyli {percentage_ordered:.2f} % PO ze statusem ordered")
text.textLine(f" oraz {received_count} czyli {percentage_received:.2f} % PO ze statusem received.")

c.drawText(text)
c.save()
time.sleep(2)


#PATH TO FILE
pdf_file = r'C:\Users\lila_\Desktop\GitHub\mobi_ptp\report.pdf'

print(os.path.exists(pdf_file)) #spr czy plik istnieje

# email to admin
admin_email = "admin@example.com"
report_mail = outlook.CreateItem(0)
report_mail.To = admin_email
report_mail_Subject = " Daily report of PO status"
report_mail.Body = """ Hello Sap Admin,

here the report of the PO status. Our Python program -Mobi has done a great work!
Mobi has sent to requestors and asked to check, if GR can be done.
                        
Mobi wishes you a great day! 
"""

#spr czy plik już istnieje
print("Pdf exists: ", pdf_file)   

#załącznik - plik pdf
report_mail.Attachments.Add(pdf_file)

#wysyłka maila
report_mail.Send()

print("Zakończyłem zadanie. Do zobaczenia!")








