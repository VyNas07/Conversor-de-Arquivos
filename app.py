import os
import subprocess
from flask import Flask, request, redirect, url_for, render_template
from werkzeug.utils import secure_filename

# Configurações de upload
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['ALLOWED_EXTENSIONS'] = {'pdf'}

# Função para verificar se a extensão do arquivo é permitida
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# Função para converter PDF para Word usando LibreOffice
def convert_pdf_to_word(pdf_path, output_folder):
    # Verifica se o arquivo PDF existe
    if not os.path.isfile(pdf_path):
        print(f"File not found: {pdf_path}")
        return None
    
    # Define o caminho do arquivo de saída
    docx_path = os.path.join(output_folder, os.path.splitext(os.path.basename(pdf_path))[0] + '.docx')
    
    # Executa o comando LibreOffice para converter o PDF para Word
    result = subprocess.run(['soffice', '--headless', '--convert-to', 'docx', '--outdir', output_folder, pdf_path], capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"LibreOffice conversion error: {result.stderr}")
        return None
    
    if not os.path.isfile(docx_path):
        print(f"Converted file not found: {docx_path}")
        return None
    
    return os.path.basename(docx_path)

# Rota para a página inicial
@app.route('/')
def index():
    return render_template('index.html')

# Rota para lidar com o upload do arquivo
@app.route('/upload', methods=['POST'])
def upload_file():
    # Verifica se o arquivo está presente no request
    if 'file' not in request.files:
        return redirect(request.url)
    
    file = request.files['file']
    
    # Verifica se o usuário selecionou um arquivo
    if file.filename == '':
        return redirect(request.url)
    
    # Verifica se o arquivo tem uma extensão permitida
    if file and allowed_file(file.filename):
        # Protege o nome do arquivo contra ataques de path traversal
        filename = secure_filename(file.filename)
        # Define o caminho completo para salvar o arquivo
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        # Salva o arquivo no diretório de uploads
        file.save(file_path)
        
        # Converter PDF para Word
        docx_filename = convert_pdf_to_word(file_path, app.config['UPLOAD_FOLDER'])
        
        if docx_filename:
            # Renderiza a página de resultado com o resultado da conversão
            return render_template('result.html', result=docx_filename)
        else:
            return "Conversion failed", 500
    
    # Redireciona de volta para a página de upload se algo deu errado
    return redirect(request.url)

if __name__ == '__main__':
    app.run(debug=True)
