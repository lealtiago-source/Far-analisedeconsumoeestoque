from flask import Flask, render_template, request, redirect, url_for, send_file, session, flash
from werkzeug.utils import secure_filename
import os
from datetime import datetime
from analise import executar_analise  # análise principal
from analise2 import executar_analise_remume  # análise REMUME

app = Flask(__name__)
app.secret_key = 'sua_chave_supersecreta'

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# ✅ Garante que a pasta 'uploads' existe (importante para o Render)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Login simples
USUARIOS = {
    'admin': 'senha123',
    'tiago': 'ticopinico',
    'teste': '12435'
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
            executar_analise(arquivos[0], arquivos[1], arquivos[2])
            flash(f'✅ Análise concluída. <a href="/static/resultado_analise.xlsx" target="_blank">Clique aqui para baixar</a>', 'success')
        except Exception as e:
            flash(f'Erro na análise: {e}', 'error')

        return redirect(request.url)

    return render_template('painel.html')

@app.route('/analise2', methods=['POST'])
def analise2():
    arquivo_estoque = request.files.get('estoque_remume')
    arquivo_remume = request.files.get('remume')

    if not arquivo_estoque or not arquivo_remume:
        flash('Arquivos de estoque e REMUME são obrigatórios.', 'error')
        return redirect(url_for('painel'))

    estoque_path = os.path.join(app.config['UPLOAD_FOLDER'], 'estoque_remume.xlsx')
    remume_path = os.path.join(app.config['UPLOAD_FOLDER'], 'remume.xlsx')

    arquivo_estoque.save(estoque_path)
    arquivo_remume.save(remume_path)

    try:
        executar_analise_remume(estoque_path, remume_path)
        hoje = datetime.today().strftime('%d-%m-%Y')
        output = f'Estoque_REMUME_atualizado_{hoje}.xlsx'
        flash(f'✅ Análise REMUME concluída. <a href="/static/{output}" target="_blank">Baixar resultado</a>', 'success')
    except Exception as e:
        flash(f'Erro na análise REMUME: {e}', 'error')

    return redirect(url_for('painel'))


# ✅ NOVA ROTA: Análise de Orçamento
@app.route('/analise_orcamento', methods=['POST'])
def analise_orcamento():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    try:
        arquivo_dispensacao = request.files.get('orcamento_dispensacao')
        arquivo_distribuicao = request.files.get('orcamento_distribuicao')

        if not arquivo_dispensacao or not arquivo_distribuicao:
            flash('Arquivos de dispensação e distribuição são obrigatórios para a análise de orçamento.', 'error')
            return redirect(url_for('painel'))

        # Caminhos dos uploads
        dispensacao_path = os.path.join(app.config['UPLOAD_FOLDER'], 'orcamento_dispensacao.xlsx')
        distribuicao_path = os.path.join(app.config['UPLOAD_FOLDER'], 'orcamento_distribuicao.xlsx')

        arquivo_dispensacao.save(dispensacao_path)
        arquivo_distribuicao.save(distribuicao_path)

        # Importa a função (para evitar erro se o módulo não existir ainda)
        from analise_orcamento import executar_analise_orcamento
        resultado_path = executar_analise_orcamento(dispensacao_path, distribuicao_path)

        # Move resultado para /static/
        destino = os.path.join('static', 'resultado_orcamento.xlsx')
        os.replace(resultado_path, destino)

        flash(f'✅ Análise de orçamento concluída. <a href="/static/resultado_orcamento.xlsx" target="_blank">Baixar resultado</a>', 'success')
    except Exception as e:
        flash(f'❌ Erro na análise de orçamento: {e}', 'error')

    return redirect(url_for('painel'))


@app.route('/logout')
def logout():
    session.pop('usuario', None)
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True)
