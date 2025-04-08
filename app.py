from flask import Flask, request, render_template, send_file
import os
import uuid
from procesar_datos import procesar_archivo

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file1' not in request.files or 'file2' not in request.files:
        return "Error: Debes subir ambos archivos."

    file1 = request.files['file1']
    file2 = request.files['file2']

    if file1.filename == '' or file2.filename == '':
        return "Error: No se seleccionaron ambos archivos."

    # 🔹 Generar ID único para esta sesión
    session_id = str(uuid.uuid4())[:8]

    file1_path = os.path.join(UPLOAD_FOLDER, f"file1_{session_id}.xls")
    file2_path = os.path.join(UPLOAD_FOLDER, f"file2_{session_id}.xlsx")
    output_filename = f"archivo_modificado_{session_id}.xls"
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)

    file1.save(file1_path)
    file2.save(file2_path)

    resultado = procesar_archivo(file1_path, file2_path, output_path)

    if resultado:
        return send_file(output_path, as_attachment=True, download_name=output_filename,
                         mimetype='application/vnd.ms-excel')
    else:
        return "Error al procesar el archivo."

if __name__ == '__main__':
    app.run(debug=True, host="127.0.0.1", port=5000)
