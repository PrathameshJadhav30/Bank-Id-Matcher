from flask import Flask, request, send_file, render_template
import pandas as pd
import re
from io import BytesIO

app = Flask(__name__)

def extract_transaction_ids(df, column):
    # Regex pattern to match both alphanumeric IDs and numeric IDs
    pattern = re.compile(r'\b[A-Z]{2,5}\d{4,}\b|\b\d{6,}\b')
    transaction_ids = set()
    for text in df[column].dropna():
        matches = pattern.findall(str(text))
        transaction_ids.update(matches)
    return transaction_ids

def filter_rows_by_ids(df, ids, column):
    return df[df[column].apply(lambda x: any(id in str(x) for id in ids))]

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/match', methods=['POST'])
def match_data():
    try:
        file1 = request.files['file1']
        file2 = request.files['file2']
        column = request.form['column']

        # Check if files are provided
        if not file1 or not file2:
            return render_template('index.html', message="Please upload both files.")

        # Load Excel files
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # Check if the column exists in both DataFrames
        if column not in df1.columns or column not in df2.columns:
            return render_template('index.html', message="The specified column does not exist in both files.")

        # Extract transaction IDs from both files
        ids1 = extract_transaction_ids(df1, column)
        ids2 = extract_transaction_ids(df2, column)

        # Find matches and unmatched data
        matched_ids = ids1.intersection(ids2)
        unmatched_ids = ids1 - matched_ids

        # Filter rows in Excel 1 based on matched and unmatched IDs
        matched_rows = filter_rows_by_ids(df1, matched_ids, column)
        unmatched_rows = filter_rows_by_ids(df1, unmatched_ids, column)

        # Save to bytes
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            matched_rows.to_excel(writer, sheet_name='Matched Data', index=False)
            unmatched_rows.to_excel(writer, sheet_name='Unmatched Data', index=False)
        
        output.seek(0)
        return send_file(output, as_attachment=True, download_name='matching_results.xlsx')

    except Exception as e:
        return render_template('index.html', message=f"An error occurred: {e}")

if __name__ == "__main__":
    app.run(debug=True)
