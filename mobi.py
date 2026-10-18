# -*- coding: utf-8 -*-
"""
Created on Wed Mar 11 23:40:59 2026
Extended on Sun Oct 18 18:15:20 2026

@author: Kamila Dudzińska

The program processes an Excel file and checks the status of each purchase 
order (PO). If the PO status is marked as "Ordered" but the delivery date 
is in the past, the program automatically sends an email notification to the 
responsible user, requesting clarification with the buyer or requester 
regarding the order status.

In procurement, for an invoice to be processed and paid, the three-way match 
principle must be fulfilled, meaning that the purchase order, goods receipt, 
and invoice must all match. If the goods receipt (GR) is missing, the invoice 
cannot be paid because there is no confirmation that the goods have been 
received or in what quantity.

Additionally, the program calculates key statistics and sends a summary email 
to the administrator, including a PDF report as an attachment. By automating 
email notifications to buyers and requesters, the solution addresses a real 
business issue commonly encountered in large Shared Service Centers (SSCs).

The program operates on the file procurement_dataset1.csv, which was generated 
using a Python script developed by me. The script creates realistic mock 
procurement datasets while preserving the data structure and business logic 
typically found in SAP Ariba procurement environments.

In large FMCG organizations, a so-called Auto-GR threshold is often used 
(for example, EUR 500), allowing low-value invoices to be paid without manual 
goods receipt verification. This helps streamline the invoice processing and 
payment workflow while reducing administrative effort.


"""

# IMPORT MODULES I
import os
from datetime import datetime, timedelta
import pandas as pd
import reportlab
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from mobi_function import divide_z

try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False   # no Outlook (ex. Linux) -> dry run


# SETTINGS
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_NAME = "procurement_mock_2500.csv"
PDF_FILE = os.path.join(BASE_DIR, "report.pdf")
ADMIN_EMAIL = "admin@example.com"
DRY_RUN = True          # True = max open windows - MAX_PREVIEW
MAX_PREVIEW = 3         # max. Outlook windows to open in test mode
GR_BUFFER_DAYS = 3      # remind only when delivery date is older than 3 days


# LOAD DATA II
# CSV next to the script, or in the data/ folder (GitHub version)
file_path = os.path.join(BASE_DIR, CSV_NAME)
if not os.path.exists(file_path):
    file_path = os.path.join(BASE_DIR, "data", CSV_NAME)
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
emails_sent = 0

# daty
today = datetime.today()
new_date = today - timedelta(days=GR_BUFFER_DAYS)

# GR MISSING - ONE EMAIL PER BUYER V
gr_missing = df[(df['Order Status'] == "ordered")
                & (df['Delivery Date'] < new_date)]
buyers = gr_missing.groupby(['Requester Mail', 'Requester Name'])
po_missing_gr = len(gr_missing)

# STATISTICS IV
total_po = len(df)
status_counts = df['Order Status'].value_counts()    # all statuses, largest first

# GR missing - value per currency (next step: FX API, convert to EUR)
missing_value = (gr_missing.groupby('Currency')['Amount'].sum()
                 .sort_values(ascending=False))

# GR missing - aging (days after delivery date)
days_overdue = (today - gr_missing['Delivery Date']).dt.days
aging = pd.cut(days_overdue,
               bins=[0, 30, 90, 180, float('inf')],
               labels=["up to 30 days", "31-90 days", "91-180 days", "over 180 days"])
aging_counts = aging.value_counts().sort_index()

# requestors with the most PO without GR
requestors_with_max_missings = (gr_missing.groupby('Requester Name').size()
                                .sort_values(ascending=False)
                                .head(5))

# OUTLOOK PART VI
outlook = None
if OUTLOOK_AVAILABLE:
    outlook = win32com.client.Dispatch("Outlook.Application")

for (email, name), pos in buyers:
    pos = pos.sort_values('Delivery Date')

    # one table row per PO
    rows_html = ""
    for _, po in pos.iterrows():
        rows_html += (
            f"<tr><td>{po['PO Number']}</td>"
            f"<td>{po['Supplier Name']}</td>"
            f"<td>{po['Delivery Date']:%Y-%m-%d}</td>"
            f"<td style='text-align:right'>{po['Amount']:,} {po['Currency']}</td></tr>"
        )

    subject = f"GR missing - {len(pos)} PO(s) waiting for Goods Receipt"
    body = f"""
    <p>Hello {name},</p>
    <p>the following PO(s) are still in status <b>ordered</b>, although the
    delivery date has passed. Could you please post the Goods Receipt (GR)
    in Ariba or let us know if the delivery is delayed?</p>
    <table border="1" cellpadding="5" style="border-collapse:collapse;font-family:Calibri">
    <tr style="background:#BFEFFF"><th>PO Number</th><th>Supplier</th>
    <th>Delivery Date</th><th>Amount</th></tr>
    {rows_html}
    </table>
    <p>Thank you in advance!</p>
    <p>Kind regards,<br>Admin</p>
    """

    # tworzenie maila
    if outlook and (not DRY_RUN or emails_sent < MAX_PREVIEW):
        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = subject
        mail.HTMLBody = body
        mail.Display()   # Display = preview only; change to Send() to send
        emails_sent += 1
        print(f"Email created for {email} ({len(pos)} PO)")
    else:
        print(f"[dry run] {email}: {len(pos)} PO")

print(f"Buyers to remind: {buyers.ngroups}, PO without GR: {po_missing_gr}")

# ADMIN PART - PDF REPORT VII
c = canvas.Canvas(PDF_FILE, pagesize=A4)
width, height = A4
NAVY = colors.HexColor("#1F3A5F")
BLUE = colors.HexColor("#BFEFFF")
GREY = colors.HexColor("#6B7280")

# header bar
c.setFillColor(NAVY)
c.rect(0, height - 90, width, 90, stroke=0, fill=1)
c.setFillColor(colors.white)
c.setFont("Helvetica-Bold", 22)
c.drawString(50, height - 50, "Mobi - GR Missing Report")
c.setFont("Helvetica", 11)
c.drawString(50, height - 72, f"Generated: {today:%Y-%m-%d %H:%M}")

# KPI tiles
tiles = [("PO in report", total_po),
         ("Buyers reminded", buyers.ngroups),
         ("PO without GR", po_missing_gr),
         ("Emails created", emails_sent)]
tile_w, tile_h, gap = 115, 70, 12
y_tiles = height - 190
for i, (label, value) in enumerate(tiles):
    x = 50 + i * (tile_w + gap)
    c.setFillColor(BLUE)
    c.roundRect(x, y_tiles, tile_w, tile_h, 8, stroke=0, fill=1)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(x + tile_w / 2, y_tiles + 32, f"{value:,}")
    c.setFont("Helvetica", 9)
    c.setFillColor(GREY)
    c.drawCentredString(x + tile_w / 2, y_tiles + 14, label)

def draw_bar(y, label, count, share, bar_max=260):
    """One row: label, grey background bar, navy bar for the share, number."""
    c.setFont("Helvetica", 10)
    c.setFillColor(colors.black)
    c.drawString(50, y + 3, label)
    c.setFillColor(colors.HexColor("#E5E7EB"))
    c.rect(170, y, bar_max, 13, stroke=0, fill=1)
    c.setFillColor(NAVY)
    c.rect(170, y, bar_max * share / 100, 13, stroke=0, fill=1)
    c.setFillColor(colors.black)
    c.drawString(170 + bar_max + 10, y + 3, f"{count:,}  ({share:.1f}%)")


def section_title(y, title):
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(50, y, title)
    c.setStrokeColor(BLUE)
    c.setLineWidth(2)
    c.line(50, y - 6, width - 50, y - 6)


# 1. PO status - all statuses
y = y_tiles - 40
section_title(y, "PO status (all PO)")
y -= 12
for status, count in status_counts.items():
    y -= 22
    draw_bar(y, status.capitalize(), count, divide_z(int(count), total_po) * 100)

# 2. GR missing - aging
y -= 40
section_title(y, "GR missing - days after delivery date")
y -= 12
for bucket, count in aging_counts.items():
    y -= 22
    draw_bar(y, bucket, count, divide_z(int(count), po_missing_gr) * 100)

# 3. GR missing - value per currency
y -= 40
section_title(y, "GR missing - blocked value per currency")
y -= 12
for currency, amount in missing_value.items():
    y -= 20
    c.setFont("Helvetica", 10)
    c.setFillColor(colors.black)
    c.drawString(50, y, currency)
    c.drawRightString(260, y, f"{amount:,.0f}")

# 4. requestors with max missings
y -= 40
section_title(y, "Requestors with max missings (PO without GR)")
y -= 12
for rank, (name, count) in enumerate(requestors_with_max_missings.items(), start=1):
    y -= 22
    draw_bar(y, f"{rank}. {name}", count, divide_z(int(count), po_missing_gr) * 100)

# footer
c.setFont("Helvetica-Oblique", 9)
c.setFillColor(GREY)
c.drawString(50, 40, f"Rule: status 'ordered' and delivery date older than {GR_BUFFER_DAYS} days.")
c.drawString(50, 28, "Source: SAP Ariba mock data | Mobi by Kamila Dudzinska")
c.save()
print("PDF saved:", PDF_FILE)

# email to admin
report_body = f"""Hello SAP Admin,

here is today's GR missing report.
Mobi reminded {buyers.ngroups} buyers about {po_missing_gr} PO(s) without Goods Receipt.
Details are in the attached PDF.

Mobi wishes you a great day!
"""

if outlook:
    report_mail = outlook.CreateItem(0)
    report_mail.To = ADMIN_EMAIL
    report_mail.Subject = "Daily report of PO status"
    report_mail.Body = report_body
    report_mail.Attachments.Add(PDF_FILE)
    if DRY_RUN:
        report_mail.Display()   # test mode: preview only
    else:
        report_mail.Send()
else:
    print("Outlook not available - admin email not created.")

print("Done! See ya!")
