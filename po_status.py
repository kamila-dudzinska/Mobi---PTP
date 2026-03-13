import pandas as pd
import os

print(os.getcwd())

# %%
#zmiana working directory
new_dir = "C:\\Users\\lila_\\Desktop\\Programowanie_All\\Krakowiak_Pandas\\big"
os.chdir(new_dir)
print(new_dir)

# %%
"""
@author: krakowiakpawel9@gmail.com
@site: e-smartdata.org
"""

import pandas as pd

# biblioteka seaborn ma przykładowe zbiory danych
from seaborn import load_dataset


#pobieranie danych finansowych ze stoog
def fetch_financial_data(company='AMZN'):
    """
    This function fetch stock market quotations.
    """
    import pandas_datareader.data as web
    return web.DataReader(name=company, data_source='stooq')


#generowanie obiektów typ DataFrame
#zapisywanie DF jako csv.
def main():
    
    #ludzie
    df = pd.DataFrame(data={'age': [12, 13, 21, 18],
                            'name': ['Paul', 'John', 'Mike', 'Donald'],
                            'has_married': [False, False, True, False],
                            'has_house': [0, 0, 1, 1],
                            'height': [185.0, 176.5, 192.0, 182.5]})

    df.to_csv('people.csv', index=False)


    # dane numeryczne
    df = pd.DataFrame(data={'var1': [3, 2, 4, 1],
                            'var2': [1.2, 4.2, 2.4, 0.2],
                            'var3': [1 / 3, 1 / 7, 1 / 4, 1 / 21]})

    df.to_csv('numeric_types.csv', index=False)

    # dane Bool
    df = pd.DataFrame(data={'var1': [0, 1, 0, 1],
                            'var2': [True, False, False, True],
                            'var3': ['T', 'F', 'F', 'F'],
                            'var4': ['True', 'True', 'False', 'False']})

    df.to_csv('boolean_types.csv', index=False)

    #dane tekstowa
    df = pd.DataFrame(data={'var1': ['001', '002', '003', '004'],
                            'var2': ['python', 'sql', 'java', 'scala'],
                            'var3': ['Python 3.8.0', 'SQL', 'Java SE 8 (LTS)', 'Scala 2.13.1']})

    df.to_csv('string_types.csv', index=False)

    #csv
    df = pd.DataFrame(data={'var1': pd.date_range('2019-01-01', periods=4, freq='H'),
                            'var2': pd.date_range('2019-01-01', periods=4, freq='D'),
                            'var3': pd.date_range('2019-01-01', periods=4, freq='BQS'),
                            'var4': pd.date_range('2019-01-01', periods=4, freq='Min'),
                            'var5': ['01-01-2019', '02-01-2019', '03-01-2019', '04-01-2019'],
                            'var6': ['00:01:00', '00:02:30', '00:03:00', '00:04:45']})

    df.to_csv('date_time_types.csv', index=False)


    # zbiór z biblitoeki Seaborn
    df = load_dataset('tips')
    df.to_csv('tips.csv', index=False)

    #notowania google
    df = fetch_financial_data('GOOGL')
    df.to_csv('google.csv')

# __name__= '__main__' - uruchaminaie funkcji main
if __name__ == '__main__':
    main()
    print('Dane zostały wygenerowane.')
    
# %%   
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
    supplier_email = row['Mail']
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
        mail.To = supplier_email
        mail.Subject = subject
        mail.Body = body
        
        emails_sent += 1
        
        #wysyłanie maila
        mail.Display()
        
        print(f'Email sent to {supplier_email}')
        
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

        
# %%
        

print(os.getcwd())
print(os.path.exists(pdf_file))


















    
    
    
    