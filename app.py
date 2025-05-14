from flask import Flask, render_template, request, send_file
import os
from procesar_excel_bdc import procesar_excel
import tempfile

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        margen = float(request.form.get('margenBDC', 5.0))
        if file:
            with tempfile.TemporaryDirectory() as tmpdir:
                input_path = os.path.join(tmpdir, file.filename)
                file.save(input_path)
                output_path = procesar_excel(input_path, margen, return_path=True)
                return send_file(output_path, as_attachment=True)
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True) 