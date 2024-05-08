from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import io 
import numpy as np  # Add numpy for handling NaN values

app = Flask(__name__)

# Load reference data
reference_df = pd.read_excel('unique_surnames_varna.xls')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    # Check if file is uploaded
    if 'file' not in request.files:
        return redirect(request.url)

    file = request.files['file']
    # Check if filename is empty
    if file.filename == '':
        return redirect(request.url)

    # Read uploaded file into DataFrame
    input_df = pd.read_excel(file)

    # Match entries and extract associated data
    result_df = pd.merge(input_df, reference_df, on='lastname', how='left')

    # Replace NaN values with empty string ('')
    result_df.replace(np.nan, '', inplace=True)

    # Save matched and unmatched data to Excel files
    matched_output = io.BytesIO()
    unmatched_output = io.BytesIO()

    with pd.ExcelWriter(matched_output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, sheet_name='Matched Data')

    unmatched_last_names = input_df[~input_df['lastname'].isin(result_df['lastname'])]
    with pd.ExcelWriter(unmatched_output, engine='xlsxwriter') as writer:
        unmatched_last_names.to_excel(writer, index=False, sheet_name='Unmatched Last Names')

    matched_output.seek(0)
    unmatched_output.seek(0)

    # Render result template with data and pass result_df
    return render_template('result.html', tables=[result_df.to_html(classes='data')],
                           titles=result_df.columns.values.tolist(),
                           matched_data=matched_output.getvalue(),
                           unmatched_data=unmatched_output.getvalue(),
                           result_df=result_df)  # Pass result_df to the template
