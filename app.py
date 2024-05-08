from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import io 
import numpy as np  

app = Flask(__name__)

# Loading reference data
reference_df = pd.read_excel('unique_surnames_varna.xls')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
  
    if 'file' not in request.files:
        return redirect(request.url)

    file = request.files['file']
   
    if file.filename == '':
        return redirect(request.url)

   
    input_df = pd.read_excel(file)

 
    result_df = pd.merge(input_df, reference_df, on='lastname', how='left')

   
    result_df.replace(np.nan, '', inplace=True)

    
    matched_output = io.BytesIO()
    unmatched_output = io.BytesIO()

    with pd.ExcelWriter(matched_output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, sheet_name='Matched Data')

    unmatched_last_names = input_df[~input_df['lastname'].isin(result_df['lastname'])]
    with pd.ExcelWriter(unmatched_output, engine='xlsxwriter') as writer:
        unmatched_last_names.to_excel(writer, index=False, sheet_name='Unmatched Last Names')

    matched_output.seek(0)
    unmatched_output.seek(0)

    
    return render_template('result.html', tables=[result_df.to_html(classes='data')],
                           titles=result_df.columns.values.tolist(),
                           matched_data=matched_output.getvalue(),
                           unmatched_data=unmatched_output.getvalue(),
                           result_df=result_df)  # Passing result_df to the template
