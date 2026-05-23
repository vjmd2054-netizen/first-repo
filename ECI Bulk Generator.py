import threading
import tkinter
from tkinter import ttk, messagebox
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import pandas as pd
import csv
import os
from datetime import datetime


class DirectPDFGenerator:
    """Generate PDF invoices directly from Excel data without DOCX files"""

    def __init__(self):
        self.styles = getSampleStyleSheet()

        # NEW: Create paragraph styles for table cells with proper wrapping
        self.cell_style_center = ParagraphStyle(
            'CellStyleCenter',
            parent=self.styles['Normal'],
            fontSize=7,
            leading=9,  # Line spacing
            alignment=TA_CENTER,  # Center alignment
            wordWrap='CJK'  # Enable text wrapping
        )

        self.cell_style_left = ParagraphStyle(
            'CellStyleLeft',
            parent=self.styles['Normal'],
            fontSize=7,
            leading=9,
            alignment=TA_LEFT,  # Left alignment
            wordWrap='CJK'  # Enable text wrapping
        )

        self.cell_style_right = ParagraphStyle(
            'CellStyleRight',
            parent=self.styles['Normal'],
            fontSize=7,
            leading=9,
            alignment=TA_RIGHT,  # Right alignment
            wordWrap='CJK'  # Enable text wrapping
        )

        self.header_style = ParagraphStyle(
            'HeaderStyle',
            parent=self.styles['Normal'],
            fontSize=6,
            leading=9,
            alignment=TA_CENTER, # Center alignment
            fontName='Helvetica-Bold',
            wordWrap='CJK' # tyjmyught,bnm]
        )

    def wrap_text_in_cell(self, text, style, max_width=200):
        """NEW: Helper function to wrap text into multiple lines"""
        if not text or text == '':
            return Paragraph("", style)

        # Convert to string and handle line breaks
        text = str(text)

        # Replace newlines with <br/> for ReportLab
        text = text.replace('\n', '<br/>')

        # Create paragraph with automatic wrapping
        return Paragraph(text, style)

    def create_items_table_for_pdf(self, data):
        """Create the items table with proper text wrapping and auto row height"""

        # Define styles for table cells
        header_style = self.header_style
        cell_style_center = self.cell_style_center
        cell_style_left = self.cell_style_left
        cell_style_right = self.cell_style_right

        # Table headers (as Paragraphs for consistency)
        headers = [
            self.wrap_text_in_cell('# of pkgs', header_style),
            self.wrap_text_in_cell('# of Units', header_style),
            self.wrap_text_in_cell('Gross Weight\n(KG)', header_style),
            self.wrap_text_in_cell('Description of Goods', header_style),
            self.wrap_text_in_cell('Harmonized Tariff\nNumber', header_style),
            self.wrap_text_in_cell('Country/ Terr.\nof MFR', header_style),
            self.wrap_text_in_cell('Unit Value', header_style),
            self.wrap_text_in_cell('Total Value', header_style)
        ]

        # Start table data with headers
        table_data = [headers]

        # Add items (handle single item from Excel)
        # NEW: Use Paragraphs instead of plain text for automatic wrapping
        description_text = data.get('DESC', '')
        tariff_text = data.get('TARIFF', '')
        country_text = data.get('SCOUNTRY', '')

        table_data.append([
            self.wrap_text_in_cell(data.get('NPAK', '1'), cell_style_center),
            self.wrap_text_in_cell(data.get('NPAK', '1'), cell_style_center),
            self.wrap_text_in_cell(data.get('WEIGHT', '0'), cell_style_center),
            self.wrap_text_in_cell(description_text, cell_style_left),  # Description can be long, will wrap
            self.wrap_text_in_cell(tariff_text, cell_style_center),
            self.wrap_text_in_cell(country_text, cell_style_center),
            self.wrap_text_in_cell(f"{float(data.get('UNIT_VALUE', '0')):.2f}", cell_style_center),
            self.wrap_text_in_cell(f"{float(data.get('VALUE', '0')):.2f}", cell_style_center)
        ])

        # Add empty rows to fill up to 8 total rows (including header)
        # NEW: Use empty Paragraphs for blank cells
        empty_paragraph = Paragraph("", cell_style_center)
        empty_row = [empty_paragraph] * 8

        while len(table_data) < 9:  # 1 header + 8 items
            table_data.append(empty_row.copy())

        # Add totals row at the end
        table_data.append([
            self.wrap_text_in_cell(data.get('NPAK', '1'), cell_style_center),
            self.wrap_text_in_cell(data.get('NPAK', '1'), cell_style_center),
            self.wrap_text_in_cell(data.get('WEIGHT', '0'), cell_style_center),
            self.wrap_text_in_cell('', cell_style_center),
            self.wrap_text_in_cell('', cell_style_center),
            self.wrap_text_in_cell('', cell_style_center),
            self.wrap_text_in_cell('', cell_style_center),
            self.wrap_text_in_cell(f"{float(data.get('VALUE', '0')):.2f}", cell_style_center)
        ])

        return table_data

    def generate_pdf(self, data, output_path):
        """Generate PDF invoice directly from data"""
        try:
            # Create PDF document with wider margins
            doc = SimpleDocTemplate(
                output_path,
                pagesize=letter,
                leftMargin=0.25 * inch,
                rightMargin=0.25 * inch,
                topMargin=0.25 * inch,
                bottomMargin=0.25 * inch
            )

            # Create story (content)
            story = []

            # Create custom styles
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=self.styles['Normal'],
                fontSize=8,
                leading=10,
                wordWrap='CJK'  # Enable text wrapping
            )

            # ========== FIRST TABLE: Exporter/Shipper vs Shipment Details ==========

            # ========== Header ==========
            top_header_style = ParagraphStyle(
                'Header',
                parent=self.styles['Normal'],
                fontSize=14,
                alignment=1,  # Center
                spaceBefore=12,
                fontName='Helvetica-Bold'
            )

            top_header_style_2 = ParagraphStyle(
                'Header',
                parent=self.styles['Normal'],
                fontSize=7,
                alignment=0,  # Center
                spaceBefore=12,
                fontName='Helvetica-Bold'
            )

            upper_text = []

            story.append(Paragraph("COMMERCIAL INVOICE", top_header_style))
            story.append(Spacer(0,15))

            upper_text_right = "This invoice must be completed in English."
            upper_text_left = "Page 1 of 1"

            upper_text.append([
                Paragraph(upper_text_right, top_header_style_2),
                Paragraph(upper_text_left, top_header_style_2)
            ])

            upper_table_text = Table(
                upper_text,
                colWidths=[7 * inch, .7 * inch]
            )

            upper_table_text.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
            ]))

            story.append(upper_table_text)
            table1_data = []

            # NEW: Wrap text in exporter section
            exporter_text = f"""
            <b>EXPORTER:</b><br/>
            <b>Tax ID#:</b> {data.get('EXPORTER_TAX_ID', '')}<br/>
            <b>Contact Name:</b> {data.get('SNAME', '')}<br/>
            <b>Telephone No.:</b> {data.get('EXPORTER_PHONE', '')}<br/>
            <b>E-Mail:</b> {data.get('EXPORTER_EMAIL', '')}<br/>
            <b>Company Name/Address:</b><br/>
            {data.get('SCOMP', '')}<br/><br/>
            {data.get('SADD1', '')}<br/>
            {data.get('SADD2', '')}<br/>
            {data.get('SCITY', '')} {data.get('SZIP', '')}<br/>
            <b>Country/Territory:</b> {data.get('SCOUNTRY', '')}
            """

            # Shipment column (right)
            shipment_text = f"""
            <b>Ship Date:</b> {data.get('SHIPDATE', '')}<br/><br/>
            <b>Air Waybill No. / Tracking No.:</b> {data.get('AWB', '')}<br/><br/>
            <b>Invoice No.: Purchase Order No.:</b><br/>
            {data.get('INVOICE_NO', data.get('AWB', ''))}<br/><br/>
            <b>Payment Terms: Bill of Lading:</b><br/>
            {data.get('PAYMENT_TERMS', 'Net 30')}<br/><br/>
            <b>Purpose of Shipment:</b><br/>
            {data.get('PURPOSE', 'SOLD')}
            """

            # NEW: Use Paragraph for auto-wrapping
            table1_data.append([
                Paragraph(exporter_text, normal_style),
                Paragraph(shipment_text, normal_style)
            ])

            # Parties to Transaction row
            parties_text = f"<b>Parties to Transaction:</b><br/>Related &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; X &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Non-Related"
            if data.get('TRANSACTION_TYPE', 'Non-Related') == "Related":
                parties_text = f"<b>Parties to Transaction:</b><br/><b>X</b> Related &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Non-Related"
            else:
                parties_text = f"<b>Parties to Transaction:</b><br/>Related &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>X</b> Non-Related"

            table1_data.append([
                Paragraph(parties_text, normal_style),
                ''
            ])

            table1 = Table(table1_data, colWidths=[4 * inch, 3.5 * inch])

            # MODIFIED: Added VALIGN='TOP' and improved padding
            table1.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Top alignment for text
                ('PADDING', (0, 0), (-1, -1), 6),  # Increased padding
                ('SPAN', (0, 1), (-1, 1)),
                # NEW: Allow row height to auto-expand
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 20), (-1, -1), 4),
            ]))

            story.append(table1)
            story.append(Spacer(1, 12))

            # ========== SECOND TABLE: Consignee vs Importer ==========
            table2_data = []

            # Consignee column (left)
            consignee_text = f"""
            <b>CONSIGNEE:</b><br/>
            <b>Tax ID#:</b> {data.get('CONSIGNEE_TAX_ID', '')}<br/>
            <b>Contact Name:</b> {data.get('CNAME', '')}<br/>
            <b>Telephone No.:</b> {data.get('CNUMBER', '')}<br/>
            <b>E-Mail:</b> {data.get('CONSIGNEE_EMAIL', '')}<br/>
            <b>Company Name/Address:</b><br/>
            {data.get('CCOMP', '')}<br/><br/>
            {data.get('CADD1', '')}<br/>
            {data.get('CADD2', '')}<br/>
            {data.get('CCITY', '')} {data.get('CZIP', '')}<br/>
            <b>Country/Territory:</b> {data.get('CCOUNTRY', 'PHILIPPINES')}
            """

            # Importer column (right)
            if data.get('SAME_AS_CONSIGNEE', True):
                importer_text = "<b>SOLD TO / IMPORTER (if different from Consignee):</b><br/><b>X Same as CONSIGNEE:</b>"
            else:
                importer_text = f"""
                <b>SOLD TO / IMPORTER (if different from Consignee):</b><br/>
                <b>Tax ID#:</b> {data.get('IMPORTER_TAX_ID', '')}<br/>
                <b>Company Name/Address:</b> {data.get('IMPORTER_COMPANY', '')}<br/>
                {data.get('IMPORTER_ADDRESS', '')}<br/>
                <b>Country/Territory:</b> {data.get('IMPORTER_COUNTRY', '')}
                """

            table2_data.append([
                Paragraph(consignee_text, normal_style),
                Paragraph(importer_text, normal_style)
            ])

            table2 = Table(table2_data, colWidths=[4 * inch, 3.5 * inch])
            table2.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Top alignment
                ('PADDING', (0, 0), (-1, -1), 6),  # Increased padding
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ]))

            story.append(table2)
            story.append(Spacer(1, 12))

            # ========== BROKER INFORMATION ==========
            if data.get('BROKER_NAME', ''):
                broker_text = f"""
                <b>If there is a designated broker for this shipment, please provide contact information.</b><br/>
                <b>Name of Broker</b> {data.get('BROKER_NAME', '')} &nbsp;&nbsp;&nbsp; 
                <b>Tel. No.</b> {data.get('BROKER_PHONE', '')} &nbsp;&nbsp;&nbsp; 
                <b>Contact Name</b> {data.get('BROKER_CONTACT', '')}
                """
                story.append(Paragraph(broker_text, normal_style))
                story.append(Spacer(1, 6))

            # ========== DUTIES AND TAXES ==========
            duties_text = f"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>Duties and Taxes Payable by</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "

            if data.get('DUTIES_PAYABLE_BY', 'Exporter') == "Exporter":
                duties_text += "<b>X Exporter</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Consignee"
            elif data.get('DUTIES_PAYABLE_BY', 'Exporter') == "Consignee":
                duties_text += "Exporter &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>X Consignee</b>"
            else:
                duties_text += f"Exporter &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Consignee &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>Other If Other, please specify</b> {data.get('DUTIES_OTHER', '')}"

            story.append(Paragraph(duties_text, normal_style))
            story.append(Spacer(1, 12))

            # ========== ITEMS TABLE ==========
            # NEW: Get table data with wrapped text
            items_table_data = self.create_items_table_for_pdf(data)

            # Column widths matching the template
            col_widths = [
                0.5 * inch,  # No. of Packages
                0.7 * inch,  # No. of Units
                0.9 * inch,  # Gross Weight
                2.0 * inch,  # Description (will auto-wrap)
                1.0 * inch,  # Tariff Number
                1.0 * inch,  # Country of MFR
                0.7 * inch,  # Unit Value
                0.7 * inch  # Total Value
            ]

            # MODIFIED: Added rowHeights=None to auto-calculate row heights
            items_table = Table(items_table_data, colWidths=col_widths, repeatRows=1, rowHeights=None)

            # Apply table styling
            items_table.setStyle(TableStyle([
                # Header row
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 7),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
                ('TOPPADDING', (0, 0), (-1, 0), 4),

                # Data rows
                ('FONTSIZE', (0, 1), (-1, -2), 7),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # MODIFIED: Changed to TOP for multi-line
                ('ALIGN', (0, 1), (1, -1), 'CENTER'),  # Packages and Units centered
                ('ALIGN', (2, 1), (2, -1), 'CENTER'),  # Weight centered
                ('ALIGN', (6, 1), (-1, -1), 'RIGHT'),  # Values right aligned
                ('ALIGN', (3, 1), (3, -1), 'LEFT'),  # Description left aligned

                # Borders for all cells
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),

                # NEW: Allow text to wrap and row height to expand
                ('WORDWRAP', (0, 0), (-1, -1), True),  # Enable word wrapping

                # Totals row styling
                ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, -1), (-1, -1), 8),

                # NEW: Add padding for better readability
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 1), (-1, -2), 6),
                ('BOTTOMPADDING', (0, 1), (-1, -2), 6),
            ]))

            story.append(items_table)
            story.append(Spacer(1, 12))

            # ========== SPECIAL INSTRUCTIONS AND TOTALS TABLE ==========
            table3_data = []

            # Calculate totals
            subtotal = float(data.get('VALUE', '0'))
            insurance = float(data.get('INSURANCE', '0'))
            freight = float(data.get('FREIGHT', '0'))
            packing = float(data.get('PACKING', '0'))
            handling = float(data.get('HANDLING', '0'))
            other = float(data.get('OTHER_CHARGES', '0'))
            invoice_total = subtotal + insurance + freight + packing + handling + other

            # Left column: Special Instructions and Declaration
            # NEW: Wrap special instructions text
            special_instructions = data.get('SPECIAL_INSTRUCTIONS', 'N/A')
            declaration = data.get('DECLARATION',
                                   'I declare that all the information contained in this invoice to be true and correct.')
            representative = data.get('REPRESENTATIVE', data.get('SNAME', ''))

            left_text = f"""
            <b>Special Instructions:</b><br/>
            {special_instructions}<br/><br/>
            <b>Declaration Statement(s):</b><br/>
            {declaration}<br/><br/>
            <b>Originator or Name of Company Representative if the invoice is being completed</b><br/>
            <b>on behalf of a company or individual:</b><br/>
            <u>{representative}</u>
            """

            # Right column: Totals
            right_data = [
                ['', 'Subtotal:', f"{subtotal:.2f}"],
                ['', 'Insurance:', f"{insurance:.2f}"],
                ['', 'Freight:', f"{freight:.2f}"],
                ['', 'Packing:', f"{packing:.2f}"],
                ['', 'Handling:', f"{handling:.2f}"],
                ['', 'Other:', f"{other:.2f}"],
                ['', 'Invoice Total:', f"{invoice_total:.2f}"],
                ['', 'Currency Code:', data.get('VCURR', 'USD')]
            ]

            # Create totals table
            totals_table = Table(right_data, colWidths=[1.3 * inch, 1.1 * inch, .5 * inch])
            totals_table.setStyle(TableStyle([
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ALIGN', (2, 0), (2, -1), 'LEFT'),
                ('FONTNAME', (1, -2), (2, -2), 'Helvetica-Bold'),
                ('FONTNAME', (1, -1), (2, -1), 'Helvetica-Bold'),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2)
            ]))

            # NEW: Wrap left text for proper formatting
            table3_data.append([
                Paragraph(left_text, normal_style),
                totals_table
            ])

            table3 = Table(table3_data, colWidths=[4.5 * inch, 3 * inch])
            table3.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('PADDING', (0, 0), (-1, -1), 4),
                ('WORDWRAP', (0, 0), (-1, -1), True),  # Enable wrapping
            ]))

            story.append(table3)
            story.append(Spacer(1, 12))

            # Build PDF
            doc.build(story)

            return True, f"PDF generated successfully: {output_path}"

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            return False, f"Failed to create PDF:\n{str(e)}\n\nDetails:\n{error_details}"


def get_country_from_code(code, country_csv_path="assets\\country_code.csv"):
    """Convert country code to country name"""
    try:
        if os.path.exists(country_csv_path):
            with open(country_csv_path, "r") as data_country:
                data = csv.DictReader(data_country)
                for result in data:
                    if result["code"] == code:
                        return result["country"].upper()
        return code
    except:
        return code


def process_excel_and_generate_pdfs(excel_path="format.xlsx", output_folder="output",
        country_csv_path="assets\\country_code.csv",
        progress_callback=None):

    """Process Excel file and generate PDFs directly for each row"""

    if not os.path.exists(excel_path):
        print(f"Error: Excel file '{excel_path}' not found!")
        return [], []

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Created output folder: {output_folder}")

    df = pd.read_excel(excel_path)
    df.columns = df.columns.str.replace("&nbsp;", "", regex=False).str.strip()

    pdf_gen = DirectPDFGenerator()
    successful = []
    failed = []
    total = len(df)

    for index, row_info in df.iterrows():
        if progress_callback:
            progress_callback(index + 1, total)
        try:
            AWB = row_info.get("AWB", "")
            CADD1 = row_info.get("ConsigneeAddr1", "")
            CADD2 = row_info.get("ConsigneeAddr2", "")
            CCITY = row_info.get("ConsigneeCity", "")
            CCOMP = row_info.get("ConsigneeCompany", "")
            CNAME = row_info.get("Consignee", "")
            CNUMBER = row_info.get("ConsigneePhone", "")
            CZIP = row_info.get("ConsigneePostal", "")
            DESC = row_info.get("ShipDesc", "")
            NPAK = row_info.get("PieceQty", "1")
            SADD1 = row_info.get("ShiprAddr1", "")
            SADD2 = row_info.get("ShiprAddr2", "")
            SCITY = row_info.get("ShiprCity", "")
            SCOMP = row_info.get("ShiprCompany", "")
            SCOUNTRY_CODE = row_info.get("OrigLocCntryCd", "")
            SHIPDATE = row_info.get("ShipDate", datetime.now())
            SNAME = row_info.get("Shipper", "")
            SZIP = row_info.get("ShiprPostalCd", "")
            VALUE = row_info.get("CustomVal", "0")
            VCURR = row_info.get("CurrencyCd", "USD")
            WEIGHT = row_info.get("KiloWgt", "0")

            if SCOUNTRY_CODE:
                SCOUNTRY = get_country_from_code(SCOUNTRY_CODE, country_csv_path)
            else:
                SCOUNTRY = ""

            if pd.notna(SHIPDATE):
                if isinstance(SHIPDATE, str):
                    SHIPDATE_STR = SHIPDATE
                else:
                    SHIPDATE_STR = SHIPDATE.strftime("%y-%m-%d")
            else:
                SHIPDATE_STR = datetime.now().strftime("%y-%m-%d")

            invoice_data = {
                'SCOMP': str(SCOMP) if pd.notna(SCOMP) else "",
                'SNAME': str(SNAME) if pd.notna(SNAME) else "",
                'SADD1': str(SADD1) if pd.notna(SADD1) else "",
                'SADD2': str(SADD2) if pd.notna(SADD2) else "",
                'SCITY': str(SCITY) if pd.notna(SCITY) else "",
                'SZIP': str(SZIP) if pd.notna(SZIP) else "",
                'SCOUNTRY': str(SCOUNTRY) if pd.notna(SCOUNTRY) else "",
                'EXPORTER_TAX_ID': row_info.get("ExporterTaxID", ""),
                'EXPORTER_PHONE': row_info.get("ExporterPhone", ""),
                'EXPORTER_EMAIL': row_info.get("ExporterEmail", ""),
                'AWB': str(AWB) if pd.notna(AWB) else "",
                'SHIPDATE': SHIPDATE_STR,
                'INVOICE_NO': f"INV-{AWB}" if pd.notna(AWB) else f"INV-{datetime.now().strftime('%Y%m%d')}-{index}",
                'PAYMENT_TERMS': row_info.get("PaymentTerms", "Net 30"),
                'PURPOSE': row_info.get("Purpose", "SOLD"),
                'TRANSACTION_TYPE': row_info.get("TransactionType", "Non-Related"),
                'CCOMP': str(CCOMP) if pd.notna(CCOMP) else "",
                'CNAME': str(CNAME) if pd.notna(CNAME) else "",
                'CADD1': str(CADD1) if pd.notna(CADD1) else "",
                'CADD2': str(CADD2) if pd.notna(CADD2) else "",
                'CCITY': str(CCITY) if pd.notna(CCITY) else "",
                'CZIP': str(CZIP) if pd.notna(CZIP) else "",
                'CNUMBER': str(CNUMBER) if pd.notna(CNUMBER) else "",
                'CCOUNTRY': row_info.get("ConsigneeCountry", "PHILIPPINES"),
                'CONSIGNEE_TAX_ID': row_info.get("ConsigneeTaxID", ""),
                'CONSIGNEE_EMAIL': row_info.get("ConsigneeEmail", ""),
                'SAME_AS_CONSIGNEE': row_info.get("SameAsConsignee", True),
                'IMPORTER_TAX_ID': row_info.get("ImporterTaxID", ""),
                'IMPORTER_COMPANY': row_info.get("ImporterCompany", ""),
                'IMPORTER_ADDRESS': row_info.get("ImporterAddress", ""),
                'IMPORTER_COUNTRY': row_info.get("ImporterCountry", ""),
                'DESC': str(DESC) if pd.notna(DESC) else "",
                'NPAK': str(NPAK) if pd.notna(NPAK) else "1",
                'WEIGHT': str(WEIGHT) if pd.notna(WEIGHT) else "0",
                'VALUE': str(VALUE) if pd.notna(VALUE) else "0",
                'VCURR': str(VCURR) if pd.notna(VCURR) else "USD",
                'UNIT_VALUE': str(float(VALUE) / float(NPAK) if NPAK and VALUE else "0"),
                'TARIFF': row_info.get("TariffNumber", ""),
                'INSURANCE': str(row_info.get("Insurance", "0")),
                'FREIGHT': str(row_info.get("Freight", "0")),
                'PACKING': str(row_info.get("Packing", "0")),
                'HANDLING': str(row_info.get("Handling", "0")),
                'OTHER_CHARGES': str(row_info.get("OtherCharges", "0")),
                'DUTIES_PAYABLE_BY': row_info.get("DutiesPayableBy", "Consignee"),
                'DUTIES_OTHER': row_info.get("DutiesOther", ""),
                'BROKER_NAME': row_info.get("BrokerName", ""),
                'BROKER_PHONE': row_info.get("BrokerPhone", ""),
                'BROKER_CONTACT': row_info.get("BrokerContact", ""),
                'SPECIAL_INSTRUCTIONS': row_info.get("SpecialInstructions", ""),
                'DECLARATION': row_info.get("Declaration",
                                            "I declare that all the information contained in this invoice to be true and correct."),
                'REPRESENTATIVE': row_info.get("Representative", SCOMP if pd.notna(SCOMP) else SNAME),
            }

            awb_str = str(AWB) if pd.notna(AWB) and str(
                AWB).strip() else f"INV_{datetime.now().strftime('%Y%m%d')}_{index + 1}"
            awb_str = "".join(c for c in awb_str if c.isalnum() or c in ('-', '_'))
            output_pdf_path = os.path.join(output_folder, f"{awb_str}.pdf")

            success, message = pdf_gen.generate_pdf(invoice_data, output_pdf_path)

            if success:
                successful.append(output_pdf_path)
            else:
                failed.append({'row': index + 1, 'awb': awb_str, 'error': message})
        except Exception as e:
            awb_str = str(row_info.get("AWB", f"row_{index + 1}"))
            failed.append({'row': index + 1, 'awb': awb_str, 'error': str(e)})

    return successful, failed

class BulkECIApp:
    def __init__(self):
        self.completed_text = None
        self.root = tkinter.Tk()
        self.root.iconbitmap("icon.ico")
        self.root.title("ECI Bulk Generator")
        self.root.geometry("620x150")
        self.build_ui()
        self.root.mainloop()

    def build_ui(self):

        style = ttk.Style()
        style.configure(
            "Run.TButton",
            background="#2E7D32",
            foreground="#194F5E",
            font=("Segoe UI", 10, "bold")
        )

        style.map(
            "Run.TButton",
            background=[
                ("active", "#388E3C"),
                ("disabled", "#A5D6A7")
            ]
        )

        text_style = ttk.Style()
        text_style.configure(
            "Text.TLabel",
            font=("Segoe UI", 10, "bold"),
            foreground="#194F5E"
        )

        self.progress = ttk.Progressbar(
            self.root, orient="horizontal", length=580, mode="determinate"
        )
        self.progress.pack(padx=10, pady=10)

        self.percent_text = ttk.Label(self.root, text="Percent: 0%", style='Text.TLabel')
        self.percent_text.pack(padx=10, pady=5)

        self.completed_text = ttk.Label(self.root, text="Completed 0 / 0", style='Text.TLabel')
        self.completed_text.pack(padx=10, pady=5)

        frame = ttk.Frame(self.root)
        frame.pack()


        self.run_btn = ttk.Button(frame, text="Run", command=self.start, style="Run.TButton")
        self.run_btn.pack(side="left", padx=5)

        ttk.Button(
            frame, text="Open output folder", style="Run.TButton",
            command=lambda: os.startfile("output"),
        ).pack(side="left", padx=5)

        ttk.Button(
            frame, text="Open format.xlsx",
            command=lambda: os.startfile("format.xlsx"), style="Run.TButton",
        ).pack(side="left", padx=5)

    def start(self):
        self.run_btn.config(state="disabled")
        self.progress["value"] = 0
        self.percent_text.config(text="Percent: 0%")
        self.completed_text.config(text="Completed 0 / 0")

        threading.Thread(target=self.run_job, daemon=True).start()

    def run_job(self):
        def update_progress(current, total):
            percent = int((current / total) * 100)
            self.root.after(
                0,
                lambda: (
                    self.progress.config(value=percent),
                    self.percent_text.config(text=f"Percent: {percent}%"),
                    self.completed_text.config(
                        text=f"Completed {current} / {total}"
                    )
                )
            )

        success, failed = process_excel_and_generate_pdfs(
            progress_callback=update_progress
        )

        self.root.after(0, lambda: self.on_done(success, failed))

    def on_done(self, success, failed):

        self.run_btn.config(state="normal")
        messagebox.showinfo(
            "Completed",
            f"Success: {len(success)}\nFailed: {len(failed)}",
        )

if __name__ == "__main__":
    BulkECIApp() # just for sample changes lang po
