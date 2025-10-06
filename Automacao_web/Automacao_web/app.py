from flask import Flask, render_template, request, send_from_directory
import os
from werkzeug.utils import secure_filename
from processador import executar_processamento

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['OUTPUT_FOLDER'] = 'output/'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/processar', methods=['POST'])
def processar():
    arquivo_dre = request.files['dre']
    path_dre = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(arquivo_dre.filename))
    arquivo_dre.save(path_dre)

    path_irrestrito, path_restrito = None, None
    
    arquivo_irrestrito = request.files.get('irrestrito')
    if arquivo_irrestrito and arquivo_irrestrito.filename != '':
        path_irrestrito = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(arquivo_irrestrito.filename))
        arquivo_irrestrito.save(path_irrestrito)

    arquivo_restrito = request.files.get('restrito')
    if arquivo_restrito and arquivo_restrito.filename != '':
        path_restrito = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(arquivo_restrito.filename))
        arquivo_restrito.save(path_restrito)

    nome_arquivo_saida, erro = executar_processamento(path_dre, path_irrestrito, path_restrito)

    return render_template('sucesso.html', nome_arquivo=nome_arquivo_saida, erro=erro)

@app.route('/download/<nome_arquivo>')
def download(nome_arquivo):
    return send_from_directory(app.config['OUTPUT_FOLDER'], nome_arquivo, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
