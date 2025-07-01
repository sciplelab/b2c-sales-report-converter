from flask import Flask, render_template, request, url_for, make_response, send_file
from io import StringIO
import os
import pandas as pd
import time
import openpyxl
import argparse
import pycountry
import logging

app = Flask(__name__)
app.config['IMPORTS_FOLDER'] = os.path.join(os.getcwd(), 'data', 'imports')
app.config['EXPORTS_FOLDER'] = os.path.join(os.getcwd(), 'data', 'exports')

# Set pandas option to opt into future behavior for downcasting
pd.set_option('future.no_silent_downcasting', True)

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%H:%M:%S')


def clean_and_map_csv(filepath):
    logging.info(f"Cleaning and mapping CSV file: {filepath}")
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()

    # Replace line separator and paragraph separator with newlines
    content = content.replace("\u2028", "\n").replace("\u2029", "\n")

    # Read cleaned content directly into a data frame
    cleaned_df = pd.read_csv(StringIO(content), low_memory=False)

    return cleaned_df


@app.route('/')
def index():
    logging.info("Rendering index page")
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    logging.info("Handling file upload")
    if 'file' not in request.files:
        return make_response("<script>alert('No file attachment in request.'); window.location.href='/';</script>")

    file = request.files['file']

    if not file.filename or not isinstance(file.filename, str):
        return make_response("<script>alert('No file uploaded. Please upload a .csv file.'); window.location.href='/';</script>")

    if file.filename.endswith('.csv'):
        filepath = os.path.join(app.config['IMPORTS_FOLDER'], str(file.filename))
        file.save(filepath)

        start_time = time.time()

        # Clean and map the uploaded CSV file
        uploaded_df = clean_and_map_csv(filepath)
        columns_df = pd.read_csv(os.path.join('templates', 'columns.csv'))


        # Column Mappings (source:target)
        column_mapping = [
            {'source': 'Billing Name', 'target': "Buyer's Name"},
            {'source': 'Email', 'target': "Buyer's E-mail"},
            {'source': 'Billing Address1', 'target': "Buyer's Address Line 1"},
            {'source': 'Billing City', 'target': "Buyer's City"},
            {'source': 'Billing Zip', 'target': "Buyer's Postal Zone"},
            {'source': 'Billing Province', 'target': "Buyer's State"},
            {'source': 'Billing Country', 'target': "Buyer's Country"},
            {'source': 'Billing Phone', 'target': "Buyer's Contact Number"},

            {'source': 'Financial Status', 'target': 'Document Type'},
            {'source': 'Name', 'target': 'Document Number'},
            {'source': 'Created at', 'target': 'Document Date'},
            {'source': 'Created at', 'target': 'Document Time'},
            {'source': 'Currency', 'target': 'Document Currency Code'},

            {'source': 'Lineitem name', 'target': 'Description of Product or Service'},
            {'source': 'Lineitem price', 'target': 'Unit Price'},
            {'source': 'Lineitem quantity', 'target': 'Quantity'},
            {'source': 'Subtotal', 'target': 'Subtotal excluding taxes discounts & charges'},
            {'source': 'Discount Amount', 'target': 'InvoiceLine Discount Amount'},
            {'source': 'Total', 'target': 'Total Excluding Tax on Line Level'},
            {'source': 'Total', 'target': 'Amount Exempted from Tax/Taxable Amount'},

            {'source': 'Total', 'target': 'Invoice Total Amount Excluding Tax'},
            {'source': 'Total', 'target': 'Invoice Total Amount Including Tax'},
            {'source': 'Total', 'target': 'Invoice Total Payable Amount'}
        ]


        # Default Column Values (target)
        default_values = {
            "Supplier's Name": "BloomThis Flora Sdn. Bhd.",
            "Supplier's TIN": "C24046757040",
            "Supplier's Registration Type": "BRN",
            "Supplier's Registration Number": "201501029070",
            "Supplier's E-mail": "accounts@bloomthis.co",
            "Supplier's MSIC code": "47734",
            "Supplier's Address Line 1": "9, Lorong 51A/227C, Seksyen 51A",
            "Supplier's City Name": "Petaling Jaya",
            "Supplier's Postal Zone": "46100",
            "Supplier's State": "Selangor",
            "Supplier's Country": "Malaysia",
            "Supplier's Contact Number": "+60162992263",

            "Buyer's TIN": "EI00000000010",

            "Classification": "008",
            "Unit of Measurement": "EA",
            "Tax Type": "E",
            "Tax Rate": "0%",
            "Tax Amount": "0",
            "Tax Exemption Reason": "Fresh flowers is exempted as per Customs Sales Tax Order 2022",
            "Sales tax exemption certificate number special exemption as per Gazette": "Fresh flowers is exempted as per Customs Sales Tax Order 2022",

            "Invoice Total Tax Amount": "0"
        }


        # Dynamic mapping logic
        mapped_df = pd.DataFrame()
        for column in columns_df.columns:
            matching_sources = [mapping['source'] for mapping in column_mapping if mapping['target'] == column]
            if matching_sources:
                for source_column in matching_sources:
                    if source_column in uploaded_df.columns:
                        # Mapping for Financial Status:Document Type
                        if column == 'Document Type' and source_column == 'Financial Status':
                            mapped_df[column] = uploaded_df[source_column].map(
                                {'paid': 'Invoice', 'Custom (POS)': 'Invoice', 'refunded': 'Refund Note', 'partially_refunded': 'Refund Note', 'expired': 'Expired', 'cancelled': 'Cancelled'}).fillna('')

                        # Split Created at into Document Date & Document Time
                        elif column == 'Document Date' or column == 'Document Time':
                            if source_column == 'Created at':
                                mapped_df['Document Date'] = uploaded_df[source_column].str.split(' ').str[0]
                                mapped_df['Document Time'] = uploaded_df[source_column].str.split(' ').str[1]
                        
                        # Calculate Subtotal.. from Lineitem price * Lineitem quantity
                        elif column == 'Subtotal excluding taxes discounts & charges' and 'Lineitem price' in uploaded_df.columns and 'Lineitem quantity' in uploaded_df.columns:
                            mapped_df[column] = uploaded_df['Lineitem price'] * uploaded_df['Lineitem quantity']
                        
                        # Calculate Total.. and Amount Exempted.. from Lineitem price - Discount Amount
                        elif column in ['Total Excluding Tax on Line Level', 'Amount Exempted from Tax/Taxable Amount'] and 'Lineitem price' in uploaded_df.columns and 'Discount Amount' in uploaded_df.columns:
                            mapped_df[column] = uploaded_df['Lineitem price'] - uploaded_df['Discount Amount'].fillna(0)
                        
                        # Handle ISO conversion from alpha-2 to alpha-3 for Buyer's Country
                        elif column == "Buyer's Country" and source_column == 'Billing Country':
                            def convert_to_alpha_3(alpha2):
                                try:
                                    country = pycountry.countries.get(alpha_2=alpha2)
                                    return country.alpha_3 if country else None
                                except LookupError:
                                    return None

                            mapped_df[column] = uploaded_df[source_column].apply(convert_to_alpha_3)
                            break
                        
                        # Handle leading ' for Buyer's Postal Zone
                        elif column == "Buyer's Postal Zone" and source_column == 'Billing Zip':
                            cleaned_series = uploaded_df[source_column].astype(str).str.lstrip("'").replace(['', 'nan'], pd.NA)
                            mapped_df[column] = cleaned_series

                        # Add mapping table for Buyer's State in Malaysia
                        elif column == "Buyer's State" and source_column == 'Billing Province':
                            state_mapping = {
                                'JHR': '01', 'KDH': '02', 'KTN': '03', 'MLK': '04', 'NSN': '05', 
                                'PHG': '06', 'PNG': '07', 'PRK': '08', 'PLS': '09', 'SGR': '10', 
                                'TRG': '11', 'SBH': '12', 'SWK': '13', 'KUL': '14', 'LBN': '15', 
                                'PJY': '16' 
                            }
                            cleaned_series = uploaded_df[source_column].map(state_mapping).replace('', pd.NA)
                            mapped_df[column] = cleaned_series.dropna()
                        
                        else:
                            mapped_df[column] = uploaded_df[source_column]
                        break
            else:
                mapped_df[column] = default_values.get(column, None)


        # Set Original Document Reference Number value as NA based on Document Date
        if 'Document Type' in mapped_df.columns and 'Document Date' in mapped_df.columns:
            mapped_df['Original Document Reference Number'] = mapped_df['Original Document Reference Number'].fillna('')
            mapped_df.loc[
                (mapped_df['Document Type'] == 'Refund Note') & (pd.to_datetime(mapped_df['Document Date']) < pd.to_datetime('2024-08-01')),
                'Original Document Reference Number'
            ] = 'NA'


        # Fill missing columns with default values
        for default_column, default_value in default_values.items():
            if default_column not in mapped_df.columns:
                mapped_df[default_column] = default_value
            else:
                mapped_df[default_column] = mapped_df[default_column].fillna(default_value)


        # Forward-fill columns based on Document Number
        mapped_df['Document Type'] = mapped_df['Document Type'].replace('', pd.NA)
        columns_to_fill = ["Buyer's Name", "Buyer's Address Line 1", "Buyer's City", "Buyer's Postal Zone", "Buyer's State", "Buyer's Country", "Buyer's Contact Number", 'Document Type', 'Document Date', 'Document Time', 'Document Currency Code', 'Invoice Total Amount Excluding Tax', 'Invoice Total Amount Including Tax', 'Invoice Total Payable Amount']
        mapped_df[columns_to_fill] = mapped_df.groupby('Document Number')[columns_to_fill].transform(lambda group: group.ffill()).infer_objects(copy=False)


        # Fill missing values for Buyer's State and Buyer's Contact Number
        mapped_df["Buyer's State"] = mapped_df["Buyer's State"].fillna('17')
        mapped_df["Buyer's Contact Number"] = mapped_df["Buyer's Contact Number"].fillna('0000000000')


        # Ensure 'Document Type' and 'Document Number' are mapped correctly
        if 'Document Type' in mapped_df.columns and 'Document Number' in mapped_df.columns:
            mapped_df.loc[mapped_df['Document Type'] == 'Refund Note', 'Original Document Reference Number'] = mapped_df.groupby('Document Number')['Document Number'].transform('first')
            mapped_df.loc[mapped_df['Document Type'] == 'Refund Note', 'Document Number'] = mapped_df['Document Number'] + '-R'


        # Create new invoice lines for Refund Note and change Document Type to Invoice
        if 'Document Type' in mapped_df.columns:
            refund_rows = mapped_df[mapped_df['Document Type'] == 'Refund Note']
            new_invoice_lines = refund_rows.copy()
            new_invoice_lines['Document Type'] = 'Invoice'
            new_invoice_lines['Document Number'] = new_invoice_lines['Document Number'].str.replace('-R', '', regex=False)
            new_invoice_lines['Original Document Reference Number'] = ''
            mapped_df = pd.concat([mapped_df, new_invoice_lines], ignore_index=True)


        # Remove all rows where Document Type is 'Expired' or 'Cancelled'
        mapped_df = mapped_df[~mapped_df['Document Type'].isin(['Expired', 'Cancelled'])]


        # Autofill directly into B2C Sales -Template-new.xlsx
        template_path = os.path.join('templates', 'B2C Sales -Template-new.xlsx')
        workbook = openpyxl.load_workbook(template_path)
        sheet = workbook.worksheets[0]

        for row_idx, row in enumerate(mapped_df.itertuples(index=False), start=4):
            for col_idx, value in enumerate(row, start=1):
                sheet.cell(row=row_idx, column=col_idx, value=value)

        output_filename = f"{os.path.splitext(os.path.basename(filepath))[0]}_cleaned_mapped.xlsx"
        output_path = os.path.join(app.config['EXPORTS_FOLDER'], output_filename)
        workbook.save(output_path)

        logging.info(f"File uploaded and processed: {filepath}")

        runtime = round(time.time() - start_time, 3)

        return render_template('index.html', download_url=url_for('download_file', filename=output_filename), runtime=runtime)
    return make_response("<script>alert('Invalid file format. Please upload a .csv file.'); window.location.href='/';</script>")


@app.route('/download/<filename>')
def download_file(filename):
    logging.info(f"Downloading file: {filename}")
    filepath = os.path.join(app.config['EXPORTS_FOLDER'], filename)

    if not os.path.exists(filepath):
        return make_response("<script>alert('File not found.'); window.location.href='/';</script>")

    return send_file(filepath, as_attachment=True)


def clear_folders():
    logging.info("Clearing folders")
    exports_folder = app.config['EXPORTS_FOLDER']
    imports_folder = app.config['IMPORTS_FOLDER']

    for folder in [exports_folder, imports_folder]:
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                logging.error(f"Error deleting file {file_path}: {e}")
    logging.info(f"Deleted {len(os.listdir(exports_folder))} exports and {len(os.listdir(imports_folder))} imports files")


if __name__ == "__main__":
    logging.info("Starting Flask application")
    parser = argparse.ArgumentParser(description="Run the Flask app or clear folders.")
    parser.add_argument("--clear", action="store_true", help="Delete files in the exports and imports folders.")
    args = parser.parse_args()

    if args.clear:
        clear_folders()
    else:
        if not os.path.exists(app.config['EXPORTS_FOLDER']):
            os.makedirs(app.config['EXPORTS_FOLDER'])
        if not os.path.exists(app.config['IMPORTS_FOLDER']):
            os.makedirs(app.config['IMPORTS_FOLDER'])
        app.run(host='0.0.0.0', port=5000, debug=True)
