import os
import pdfplumber
import pandas as pd
import re
import openpyxl

# Define sections with their subcolumns
sections = [
    ('4A, 4B, 6B, 6C - B2B, SEZ, DE Invoices', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess']),
    ('5A, 5B - B2C (Large) Invoices', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Cess']),
    ('9B - Credit / Debit Notes (Registered)', ['No. of Records', 'Total Note Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess']),
    ('9B - Credit / Debit Notes (Unregistered)', ['No. of Records', 'Total Note Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Cess']),
    ('6A - Exports Invoices', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax']),
    ('7 - B2C (Others)', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess']),
    ('8 - Nil rated, exempted and non GST outward supplies', ['No. of Records', 'Total Nil Amount', 'Total Exempted Amount', 'Total Non-GST Amount']),
    ('11A(1), 11A(2) - Tax Liability (Advances Received)', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess']),
    ('11B(1), 11B(2) - Adjustment of Advances', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess']),
    ('12 - HSN-wise summary of outward supplies', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess']),
    ('13 - Documents Issued', ['No. of Records', 'Documents Issued', 'Documents Cancelled', 'Net Issued Documents']),
    ('9A - Amended B2B Invoices', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess']),
    ('9A - Amended B2C (Large) Invoices', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Cess']),
    ('9C - Amended Credit/Debit Notes (Registered)', ['No. of Records', 'Total Note Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess']),
    ('9C - Amended Credit/Debit Notes (Unregistered)', ['No. of Records', 'Total Note Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Cess']),
    ('9A - Amended Exports Invoices', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax']),
    ('10 - Amended B2C(Others)', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess']),
    ('11A - Amended Tax Liability (Advance Received)', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess']),
    ('11B - Amendment of Adjustment of Advances', ['No. of Records', 'Total Invoice Value', 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess']),
]

def extract_data_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text() + '\n'

    # Extract metadata
    year_match = re.search(r'Year\s*(\d{4}-\d{2})', text)
    year = year_match.group(1) if year_match else 'Unknown'

    period_match = re.search(r'Period\s*([A-Za-z]+\(M\))', text)
    if not period_match:
        period_match = re.search(r'Period\s*([A-Za-z]+\(M\s*\))', text)
    period = period_match.group(1) if period_match else 'Unknown'

    gstin_match = re.search(r'GSTIN\s*([0-9A-Z]+)', text)
    gstin = gstin_match.group(1) if gstin_match else 'Unknown'

    legal_name_match = re.search(r'Legal name of the registered person\s*([A-Z\s]+)', text)
    legal_name = legal_name_match.group(1).strip() if legal_name_match else 'Unknown'

    trade_name_match = re.search(r'Trade name, if any\s*([A-Z\s]+)', text)
    trade_name = trade_name_match.group(1).strip() if trade_name_match else 'Unknown'

    arn_match = re.search(r'ARN\s*([A-Z0-9]+)', text)
    arn = arn_match.group(1) if arn_match else 'Unknown'

    submission_date_match = re.search(r'ARN date\s*(\d{2}/\d{2}/\d{4})', text)
    if not submission_date_match:
        submission_date_match = re.search(r'Date:\s*(\d{2}/\d{2}/\d{4})', text)
    submission_date = submission_date_match.group(1) if submission_date_match else 'Unknown'

    lines = text.split('\n')
    section_data = {}
    for section_name, subcols in sections:
        section_data[section_name] = ['0'] * len(subcols)
        for i, line in enumerate(lines):
            if section_name in line:
                # Find the next line with numbers
                for j in range(i+1, len(lines)):
                    next_line = lines[j].strip()
                    if re.match(r'^\d', next_line):
                        numbers = re.findall(r'\d+\.?\d*', next_line)
                        if len(numbers) >= len(subcols):
                            # Round to nearest integer
                            section_data[section_name] = [int(round(float(num))) for num in numbers[:len(subcols)]]
                        break
                break

    # Flatten the data
    data = {
        'Year': year,
        'Period': period,
        'GSTIN': gstin,
        'Legal Name': legal_name,
        'Trade Name': trade_name,
        'ARN': arn,
        'Submission Date': submission_date,
    }
    for section_name, values in section_data.items():
        for sec in sections:
            if sec[0] == section_name:
                subcols = sec[1]
                break
        for subcol, val in zip(subcols, values):
            data[f'{section_name} - {subcol}'] = val

    return data

def main():
    pdf_dir = 'files'
    all_data = []
    for filename in os.listdir(pdf_dir):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(pdf_dir, filename)
            data = extract_data_from_pdf(pdf_path)
            all_data.append(data)

    df = pd.DataFrame(all_data)

    # Create multi-level headers
    header1 = ['Year', 'Period', 'GSTIN', 'Legal Name', 'Trade Name', 'ARN', 'Submission Date']
    header2 = ['', '', '', '', '', '', '']
    for section_name, subcols in sections:
        header1.extend([section_name] + [''] * (len(subcols) - 1))
        header2.extend(subcols)

    output_file = 'consolidated_gstr1.xlsx'
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write data starting from row 3 (0-indexed row 2)
        df.to_excel(writer, startrow=2, index=False, header=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        # Write header1 to row 1 (1-indexed)
        for i, val in enumerate(header1):
            worksheet.cell(row=1, column=i+1, value=val)
        # Write header2 to row 2 (1-indexed)
        for i, val in enumerate(header2):
            worksheet.cell(row=2, column=i+1, value=val)
        # Merge cells for each section
        col = 8  # Start after metadata columns
        for section_name, subcols in sections:
            end_col = col + len(subcols) - 1
            worksheet.merge_cells(start_row=1, start_column=col, end_row=1, end_column=end_col)
            worksheet.cell(row=1, column=col).alignment = openpyxl.styles.Alignment(horizontal='center', wrap_text=True)
            col = end_col + 1
        # Enable wrap text and center alignment for all header cells
        for row in [1, 2]:
            for col in range(1, len(header1) + 1):
                cell = worksheet.cell(row=row, column=col)
                cell.alignment = openpyxl.styles.Alignment(wrap_text=True, horizontal='center')
    print("Consolidated data saved to consolidated_gstr1.xlsx")

if __name__ == '__main__':
    main()
