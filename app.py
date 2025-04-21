
from flask import Flask, render_template, request, send_file
import tempfile
import os
from process import process_excel_file

app = Flask(__name__)

REFERENCE_FILE = "modele.xlsx"

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files or request.files['file'].filename == '':
            return "Aucun fichier sélectionné", 400
        
        uploaded_file = request.files['file']
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xls') as temp_input:
            input_path = temp_input.name
            uploaded_file.save(input_path)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_output:
            output_path = temp_output.name
        
        try:
            process_excel_file(input_path, REFERENCE_FILE, output_path)
            return send_file(output_path, as_attachment=True, download_name="fichier_traite.xlsx")
        except Exception as e:
            return f"Erreur lors du traitement du fichier : {e}", 500
        finally:
            os.remove(input_path)
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
