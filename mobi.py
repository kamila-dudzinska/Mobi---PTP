# -*- coding: utf-8 -*-
"""
Created on Wed Mar 11 23:40:59 2026
@author: lila_
"""

#import modułów

import pandas as pd
import os
import win32com.client
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


# wczytywanie pliku excel
file_path = "C:\\Users\\lila_\\Desktop\\po_status.xlsx"
df = pd.read_excel(file_path)

# czyszczenie danych - musimy pamietać, że excel i spyder widzą te same dane inaczej
# dlatego ważne jest użycie .astype() i konwersja na odpowieni format
df['Name'] = df['Name'].astype(str)
df['PO Number'] = df['PO Number'].astype(str)
df['Delivery date'] = df['Delivery date'].astype(str).str.strip()
df['Delivery date'] = pd.to_datetime(df['Delivery date'], 
                                     errors='coerce', 
                                     format="%Y-%m-%d")
df['Amount'] = df['Amount'].astype(int)


#zmienne do statystyk
ordered_count = 0
received_count = 0
emails_sent = 0
suma = ordered_count + received_count


# wartoci procentowe i funkcja zabezpieczająca dzielenie przez zero
a = ordered_count
b = suma                    #ordered_count + received_count

#ta funkcja mogłaby się znależć w osobnym module, gdyby kod był bardziej rozbudowany
def divide_z(a, b, default=0):
    """
    Robimy funkcję dzielenia z zabezpieczeniem dzielenia przez zero.
    
    """
    try:
        # Sprawdzenie typu danych
        if not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
            raise TypeError("Oba argumenty muszą być liczbami.")
        
        return a / b
    except ZeroDivisionError:
        return default
    except TypeError as e:
        print(f"Błąd: {e}")
        return default



# dzisiejszy dzień
today = datetime.today()

new_date = today - timedelta(days=3)

print(type(new_date))


# połączenie z outlookiem
outlook =  win32com.client.Dispatch("Outlook.Application")

#iteracja po wierszach excela
for index, row in df.iterrows():
    
    status = row['Status']
    mail = row['Mail']
    name = row['Name']
    po_number = row['PO Number']
    delivery_date = row['Delivery date']
    amount =  row['Amount']
    
    if status == "Ordered" and delivery_date < new_date:
        ordered_count += 1
        subject = "Please check GR Missing"
        body = f"""
        Hello {name},
        
        The PO {po_number}  is in status ordered, however the delivery date {delivery_date} is in the past. 
        Could you be so kind and check it? The order amount is {amount}.
        
        Thank you in advance!
        
        Kind regards,
        Kamila
        """
        
        # tworzenie maila
        mail = outlook.CreateItem(0)
        mail.To = mail
        mail.Subject = subject
        mail.Body = body
        emails_sent += 1
        
        #wysyłanie maila
        mail.Send()
        
        print(f'Email sent to {mail}')
        
    elif status == "Received":
        
        received_count += 1
        
    else:
        
        ordered_count +=1
        
    
        
    suma +=1 

# obliczenia do statystyk   
# korzystamy ze stworzonej przez nas funkcji 
a = ordered_count
b = received_count  
percentage_ordered = divide_z(a, suma) *100
percentage_received = divide_z(b, suma) *100
        
print(ordered_count, received_count, suma)
              
# czesc adminowa

#tworzenie nowego pliku pdf
pdf_file = os.path.abspath("C:\\Users\\lila_\\Desktop\\Statistics.pdf")

c=canvas.Canvas(pdf_file, pagesize=A4)

# tytuł
y=720
c.setFont("Helvetica-Bold", 14)
c.setFillColorRGB(0,0,50)
c.drawString(100, y, f"Report of PO status")
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

print("Pdf exists: ", pdf_file)   #spr czy plik istnieje

#podajemy cieżkę do pliku
pdf_file = os.path.abspath("C:\\Users\\lila_\\Desktop\\Statistics.pdf")

print(os.path.exists(pdf_file))

# dane administatora
admin_email = "kamila.dudzinska@onet.pl"
report_mail = outlook.CreateItem(0)
report_mail.To = admin_email
report_mail_Subject = " Daily report of PO status"
report_mail.Body = """ Hello Admin,

here the report of the PO status. Our Python program -Mobi has done a great work!
Mobi has sent to requestors and asked to check, if GR can be done.
                        
Mobi wishes you a great day! 
                        
                        
  """

#załącznik - plik pdf
report_mail.Attachments.Add(pdf_file)

#wysyłka maila
report_mail.Send()

print("Zakończyłem zadanie. Do zobaczenia!")
