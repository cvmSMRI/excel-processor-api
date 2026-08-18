from flask import Flask, request, send_file
import pandas as pd
import io

app = Flask(__name__)

# Target columns you want to extract from each sheet
KSP_COLUMNS = [
    "SKU No.", "Item Description", "Vendor Name", 
    "Unit Cost", "Category", "Status"
]

TWP_COLUMNS = [
    "SKU No.", "Selling Price", "Stock Qty", 
    "Branch Code", "Reorder Level"
]

@app.route('/process-excel', methods=['POST'])
def process_excel():
    try:
        # Check if file was included in the request
        if 'file' not in request.files:
            return {"error": "No file uploaded"}, 400
        
        uploaded_file = request.files['file']
        
        # Load the uploaded Excel workbook
        excel_reader = pd.ExcelFile(uploaded_file)
        
        # Output buffer to write the clean Excel file in memory
        output_stream = io.BytesIO()
        
        with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
            
            # --- Process KSP Sheet ---
            if 'KSP' in excel_reader.sheet_names:
                df_ksp = pd.read_excel(excel_reader, sheet_name='KSP')
                # Filter only existing target columns
                ksp_cols_exist = [col for col in KSP_COLUMNS if col in df_ksp.columns]
                df_ksp_filtered = df_ksp[ksp_cols_exist]
                df_ksp_filtered.to_excel(writer, sheet_name='KSP Data', index=False)
            
            # --- Process TWP Sheet ---
            if 'TWP' in excel_reader.sheet_names:
                df_twp = pd.read_excel(excel_reader, sheet_name='TWP')
                twp_cols_exist = [col for col in TWP_COLUMNS if col in df_twp.columns]
                df_twp_filtered = df_twp[twp_cols_exist]
                df_twp_filtered.to_excel(writer, sheet_name='TWP Data', index=False)

        output_stream.seek(0)
        
        # Return the processed file directly as an .xlsx download
        return send_file(
            output_stream,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Processed_Summary.xlsx'
        )

    except Exception as e:
        return {"status": "error", "message": str(e)}, 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)