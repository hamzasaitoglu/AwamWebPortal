import os
import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

def generate_awam_invoice(filename, invoice_num, invoice_date, customer_info, items_data, logo_path="AG-LOGO.png"):
    doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()

    # Custom Text Styles
    header_company_style = ParagraphStyle(
        'HeaderCompany',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=14,
        alignment=1,
        textColor=colors.HexColor("#0B1B3D")
    )
    
    header_details_style = ParagraphStyle(
        'HeaderDetails',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=12,
        alignment=1,
        textColor=colors.HexColor("#1A2530")
    )

    cell_style = ParagraphStyle(
        'CellText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=11,
        alignment=0
    )
    
    cell_center = ParagraphStyle(
        'CellCenter',
        parent=cell_style,
        alignment=1
    )

    cell_right = ParagraphStyle(
        'CellRight',
        parent=cell_style,
        alignment=2
    )

    th_style = ParagraphStyle(
        'THText',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=11,
        alignment=1,
        textColor=colors.white
    )

    # 1. Logo Insertion
    if os.path.exists(logo_path):
        logo = Image(logo_path, width=120, height=72)
        logo.hAlign = 'CENTER'
        story.append(logo)
        story.append(Spacer(1, 10))

    # 2. Company Header Info
    company_header_text = """
    <b>AWAM GLOBAL LOJİSTİK TİCARET LİMİTED ŞİRKETİ</b><br/>
    <font size=8><b>ADDRESS:</b> Mahmudiye Mahallesi Ertugrulgazi Caddesi No:55 ic kapi:<br/>
    3 Inegol / BURSA / TURKIYE<br/>
    <b>VN:</b> 0911212625 &nbsp;&nbsp; <b>VD:</b> INEGOL<br/>
    <b>EMAIL:</b> tr.finans@awamlogistics.com<br/>
    <b>TEL:</b> +90 224 502 8395</font>
    """
    story.append(Paragraph(company_header_text, header_company_style))
    story.append(Spacer(1, 15))

    # 3. Customer Details & Invoice Info Box
    customer_html = f"<b>Customer Details</b><br/>{customer_info.replace('\n', '<br/>')}"
    
    cust_p = Paragraph(f"<font color='white'><b>Customer Details</b></font>", th_style)
    cust_val = Paragraph(customer_info.replace('\n', '<br/>'), cell_style)
    
    inv_num_th = Paragraph("Invoice Number", th_style)
    inv_num_val = Paragraph(invoice_num, cell_center)
    inv_date_th = Paragraph("Date", th_style)
    inv_date_val = Paragraph(invoice_date, cell_center)

    meta_table_data = [
        [cust_p, inv_num_th],
        [cust_val, inv_num_val],
        ['', inv_date_th],
        ['', inv_date_val]
    ]

    meta_table = Table(meta_table_data, colWidths=[330, 210])
    meta_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (0, 0)),
        ('SPAN', (0, 1), (0, 3)),
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor("#0B1B3D")),
        ('BACKGROUND', (1, 0), (1, 0), colors.HexColor("#0B1B3D")),
        ('BACKGROUND', (1, 2), (1, 2), colors.HexColor("#0B1B3D")),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#0B1B3D")),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#0B1B3D")),
        ('BACKGROUND', (0, 1), (0, 3), colors.white),
        ('BACKGROUND', (1, 1), (1, 1), colors.white),
        ('BACKGROUND', (1, 3), (1, 3), colors.white),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 15))

    # 4. Line Items Table Construction
    items_table_data = [
        [
            Paragraph("NO", th_style),
            Paragraph("SHIPPER", th_style),
            Paragraph("DESCRIPTION", th_style),
            Paragraph("UNITS", th_style),
            Paragraph("UNIT PRICE", th_style),
            Paragraph("TOTAL", th_style)
        ]
    ]

    subtotal = 0.0
    for idx, item in enumerate(items_data, start=1):
        units = float(item.get("units", 0))
        unit_price = float(item.get("unit_price", 0))
        line_total = units * unit_price
        subtotal += line_total

        items_table_data.append([
            Paragraph(str(idx), cell_center),
            Paragraph(item.get("shipper", ""), cell_style),
            Paragraph(item.get("description", ""), cell_style),
            Paragraph(str(int(units) if units.is_integer() else units), cell_center),
            Paragraph(f"${unit_price:,.2f}", cell_right),
            Paragraph(f"${line_total:,.2f}", cell_right)
        ])

    # Minimum rows padding for design structure
    while len(items_table_data) < 9:
        items_table_data.append([
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style)
        ])

    # Subtotal and Grand Total Rows
    grand_total = subtotal  # Add tax logic here if required
    
    items_table_data.append([
        '', '', '', '',
        Paragraph("<b>SUBTOTAL</b>", cell_right),
        Paragraph(f"<b>${subtotal:,.2f}</b>", cell_right)
    ])
    items_table_data.append([
        '', '', '', '',
        Paragraph("<font color='white'><b>GRAND TOTAL:</b></font>", cell_right),
        Paragraph(f"<font color='white'><b>${grand_total:,.2f}</b></font>", cell_right)
    ])

    items_table = Table(items_table_data, colWidths=[30, 110, 175, 55, 85, 85])
    
    table_style_cmd = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2B709E")),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -3), 0.5, colors.HexColor("#2B709E")),
        ('BOX', (0, 0), (-1, -3), 1, colors.HexColor("#0B1B3D")),
        ('SPAN', (0, -2), (3, -2)),
        ('SPAN', (0, -1), (3, -1)),
        ('BACKGROUND', (4, -1), (5, -1), colors.HexColor("#0B1B3D")),
        ('GRID', (4, -2), (5, -1), 0.5, colors.HexColor("#0B1B3D")),
    ]
    
    items_table.setStyle(TableStyle(table_style_cmd))
    story.append(items_table)
    story.append(Spacer(1, 20))

    # Footer
    footer_style = ParagraphStyle(
        'FooterText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        alignment=1,
        textColor=colors.HexColor("#1A2530")
    )
    story.append(Paragraph("www.awamlogistics.com", footer_style))

    doc.build(story)