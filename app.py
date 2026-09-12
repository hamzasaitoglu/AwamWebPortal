import streamlit as st
import pandas as pd
import io
import json
import datetime
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ReportLab Engine
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ---------------------------------------------------------
# AWAM INVOICE ENGINE (UPDATED & ACCURATE)
# ---------------------------------------------------------
def clean_turkish_text(text):
    """تحويل الحروف التركية إلى مقابلها الإنجليزي القياسي لمنع ظهور المربعات السوداء"""
    if not isinstance(text, str):
        return ""
    replacements = {
        'İ': 'I', 'I': 'I', 'ı': 'i',
        'Ş': 'S', 'ş': 's',
        'Ğ': 'G', 'ğ': 'g',
        'Ü': 'U', 'ü': 'u',
        'Ö': 'O', 'ö': 'o',
        'Ç': 'C', 'ç': 'c'
    }
    for search, replace in replacements.items():
        text = text.replace(search, replace)
    return text

def build_pdf_invoice(invoice_num, invoice_date, customer_info, items_data, tax_percent=0.0, logo_path="AG-LOGO.png"):
    pdf_buffer = io.BytesIO()
    
    # 36 pt margins (0.5 inch border around page)
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()

    header_title_style = ParagraphStyle(
        'HT', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=12, leading=14,
        alignment=1, textColor=colors.HexColor("#0B1B3D")
    )
    
    header_sub_style = ParagraphStyle(
        'HS', parent=styles['Normal'],
        fontName='Helvetica', fontSize=8.5, leading=11,
        alignment=1, textColor=colors.HexColor("#1A2530")
    )

    cell_style = ParagraphStyle('CS', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=10, alignment=0)
    cell_center = ParagraphStyle('CC', parent=cell_style, alignment=1)
    cell_right = ParagraphStyle('CR', parent=cell_style, alignment=2)
    th_style = ParagraphStyle('TH', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=10, alignment=1, textColor=colors.white)

    # 1. Logo
    if os.path.exists(logo_path):
        logo = RLImage(logo_path, width=110, height=65)
        logo.hAlign = 'CENTER'
        story.append(logo)
        story.append(Spacer(1, 8))

    # 2. Company Header (Cleaned Turkish Text)
    comp_title = clean_turkish_text("AWAM GLOBAL LOJISTIK TICARET LIMITED SIRKETI")
    comp_address = clean_turkish_text("ADDRESS: Mahmudiye Mahallesi Ertugrulgazi Caddesi No:55 ic kapi: 3 Inegol / BURSA / TURKIYE")
    
    story.append(Paragraph(f"<b>{comp_title}</b>", header_title_style))
    story.append(Spacer(1, 3))
    
    header_info = f"""
    {comp_address}<br/>
    <b>VN:</b> 0911212625 &nbsp;&nbsp; <b>VD:</b> INEGOL<br/>
    <b>EMAIL:</b> tr.finans@awamlogistics.com &nbsp;&nbsp; <b>TEL:</b> +90 224 502 8395
    """
    story.append(Paragraph(header_info, header_sub_style))
    story.append(Spacer(1, 15))

    # 3. Customer Box & Invoice Box (Separated Design)
    cust_clean = clean_turkish_text(customer_info).replace('\n', '<br/>')
    
    cust_table_data = [
        [Paragraph("Customer Details", th_style)],
        [Paragraph(cust_clean, cell_style)]
    ]
    cust_table = Table(cust_table_data, colWidths=[260], rowHeights=[20, 50])
    cust_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor("#0B1B3D")),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#0B1B3D")),
        ('BACKGROUND', (0, 1), (0, 1), colors.white),
    ]))

    inv_table_data = [
        [Paragraph("Invoice Number", th_style)],
        [Paragraph(clean_turkish_text(invoice_num), cell_center)],
        [Paragraph("Date", th_style)],
        [Paragraph(clean_turkish_text(invoice_date), cell_center)]
    ]
    inv_table = Table(inv_table_data, colWidths=[260], rowHeights=[20, 20, 20, 20])
    inv_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor("#0B1B3D")),
        ('BACKGROUND', (0, 2), (0, 2), colors.HexColor("#0B1B3D")),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#0B1B3D")),
        ('BACKGROUND', (0, 1), (0, 1), colors.white),
        ('BACKGROUND', (0, 3), (0, 3), colors.white),
    ]))

    meta_wrapper = Table([[cust_table, '', inv_table]], colWidths=[260, 20, 260])
    meta_wrapper.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(meta_wrapper)
    story.append(Spacer(1, 15))

    # 4. Items Table
    items_table_data = [[
        Paragraph("NO", th_style),
        Paragraph("SHIPPER", th_style),
        Paragraph("DESCRIPTION", th_style),
        Paragraph("UNITS", th_style),
        Paragraph("UNIT PRICE", th_style),
        Paragraph("TOTAL", th_style)
    ]]

    subtotal = 0.0

    for idx, item in enumerate(items_data, start=1):
        shipper_val = clean_turkish_text(str(item.get("shipper", "")))
        desc_val = clean_turkish_text(str(item.get("description", "")))
        
        # Filter out NaN/empty rows
        if not shipper_val.strip() and not desc_val.strip() and str(item.get("shipper")) == "nan":
            continue

        try: units = float(item.get("units", 0))
        except: units = 0.0

        try: unit_price = float(item.get("unit_price", 0))
        except: unit_price = 0.0

        line_total = units * unit_price
        subtotal += line_total

        items_table_data.append([
            Paragraph(str(idx), cell_center),
            Paragraph(shipper_val, cell_style),
            Paragraph(desc_val, cell_style),
            Paragraph(str(int(units) if units.is_integer() else units), cell_center),
            Paragraph(f"${unit_price:,.2f}", cell_right),
            Paragraph(f"${line_total:,.2f}", cell_right)
        ])

    # Fill empty rows up to 7 items cleanly without 'nan'
    while len(items_table_data) < 8:
        items_table_data.append([
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style)
        ])

    tax_amount = subtotal * (tax_percent / 100.0)
    grand_total = subtotal + tax_amount

    # Financial Totals Rows (SUBTOTAL, TAX, GRAND TOTAL)
    items_table_data.append(['', '', '', '', Paragraph("<b>SUBTOTAL</b>", cell_right), Paragraph(f"<b>${subtotal:,.2f}</b>", cell_right)])
    items_table_data.append(['', '', '', '', Paragraph("<b>TAX</b>", cell_right), Paragraph(f"<b>${tax_amount:,.2f}</b>", cell_right)])
    items_table_data.append(['', '', '', '', Paragraph("<font color='white'><b>GRAND TOTAL:</b></font>", cell_right), Paragraph(f"<font color='white'><b>${grand_total:,.2f}</b></font>", cell_right)])

    items_table = Table(items_table_data, colWidths=[30, 110, 170, 50, 90, 90])
    items_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2B709E")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -4), 0.5, colors.HexColor("#2B709E")),
        ('SPAN', (0, -3), (3, -3)),
        ('SPAN', (0, -2), (3, -2)),
        ('SPAN', (0, -1), (3, -1)),
        ('GRID', (4, -3), (5, -1), 0.5, colors.HexColor("#0B1B3D")),
        ('BACKGROUND', (4, -1), (5, -1), colors.HexColor("#0B1B3D")),
    ]))
    story.append(items_table)
    story.append(Spacer(1, 20))

    # 5. Page Border & Footer
    def add_page_border(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#0B1B3D"))
        canvas.setLineWidth(1)
        # Outer Border Rectangle
        canvas.rect(18, 18, 576, 756)
        canvas.restoreState()

    story.append(Paragraph("www.awamlogistics.com", ParagraphStyle('FT', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, alignment=1)))

    doc.build(story, onFirstPage=add_page_border, onLaterPages=add_page_border)
    pdf_buffer.seek(0)
    return pdf_buffer
