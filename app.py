from flask import Flask, render_template, request, redirect, url_for, send_file, session, flash
from werkzeug.utils import secure_filename
import os
from analise import executar_analise  # sua função de análise

app = Flask(__name__)
app.secret_key = 'sua_chave_supersecreta'

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# ✅ Garante que a pasta 'uploads' existe (importante para o Render)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Login simples
USUARIOS = {
    'admin': 'senha123'
}

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        usuario = request.form['usuario']
        senha = request.form['senha']
        if usuario in USUARIOS and USUARIOS[usuario] == senha:
            session['usuario'] = usuario
            return redirect(url_for('painel'))
        else:
            flash('Usuário ou senha inválidos.')
    return render_template('login.html')

@app.route('/painel', methods=['GET', 'POST'])
def painel():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        arquivos = []
        for i in range(1, 4):
            arquivo = request.files.get(f'arquivo{i}')
            if not arquivo or not arquivo.filename.endswith('.xlsx'):
                flash(f'Arquivo {i} inválido.')
                return redirect(request.url)
            filename = secure_filename(arquivo.filename)
            caminho = os.path.join(app.config['UPLOAD_FOLDER'], f'arquivo{i}.xlsx')
            try:
                arquivo.save(caminho)
            except Exception as e:
                flash(f'Erro ao salvar o arquivo {i}: {e}')
                return redirect(request.url)
            arquivos.append(caminho)

        try:
            caminho_saida = os.path.join('static', 'resultado_analise.xlsx')
            executar_analise(arquivos[0], arquivos[1], arquivos[2], caminho_saida)
            flash(f'✅ Análise concluída. <a href="/static/resultado_analise.xlsx" target="_blank">Clique aqui para baixar</a>', 'success')
        except Exception as e:
            flash(f'Erro na análise: {e}', 'error')

        return redirect(request.url)

    return render_template('painel.html')

@app.route('/logout')
def logout():
    session.pop('usuario', None)
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True)