"""
Aplicação Flask.

Correções relevantes:

  1. O APP NUNCA SUBIA — o bloco `if __name__ == '__main__':` chamava
     `webbrowser.open(...)` e terminava. Faltava `app.run()`. O processo abria
     o navegador e encerrava, e o navegador exibia erro de conexão.

  2. SENHAS EM TEXTO PLANO NO CÓDIGO-FONTE — usuários e senhas ("admin" /
     "senha123") estavam literais no arquivo, e a SECRET_KEY também. Agora as
     credenciais vêm de variável de ambiente com hash (werkzeug), e a
     SECRET_KEY é obrigatória em produção.

  3. ROTAS SEM AUTENTICAÇÃO — /analise2 e /analise_compra não verificavam a
     sessão: qualquer pessoa podia enviar planilhas e disparar processamento.
     Agora todas as rotas de análise usam o decorador @login_obrigatorio.

  4. RESULTADOS PÚBLICOS EM /static — os arquivos gerados eram gravados em
     static/ com nome previsível, então qualquer pessoa na rede podia baixar
     dados de outra unidade. Agora ficam fora de static/ e são servidos pela
     rota autenticada /download/<nome>.

  5. COLISÃO ENTRE USUÁRIOS SIMULTÂNEOS — os uploads eram sempre salvos como
     'arquivo1.xlsx'. Dois usuários ao mesmo tempo sobrescreviam os arquivos um
     do outro e recebiam o resultado errado. Agora cada requisição usa uma
     pasta temporária própria, removida ao final.

  6. XSS VIA FLASH COM HTML — as mensagens montavam HTML com nomes de arquivo
     e texto de exceção, exigindo |safe no template. Agora as mensagens são
     texto puro e o link de download é renderizado pelo template.

  7. `secure_filename` era importado e nunca usado; a validação aceitava
     apenas '.xlsx' e por sufixo simples. Agora há validação de extensão,
     de nome e limite de tamanho de upload.
"""

from __future__ import annotations

import logging
import os
import shutil
import tempfile
import uuid
from datetime import datetime
from functools import wraps
from logging.handlers import RotatingFileHandler
from pathlib import Path

from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    request,
    send_from_directory,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

from analise import executar_analise
from analise_compra import executar_analise_compra
from analise_orcamento import executar_analise_orcamento
from analise_remume import executar_analise_remume
from config import (
    DIAS_ESTOQUE_PADRAO,
    DIAS_VALIDADE_PADRAO,
    EXTENSOES_PERMITIDAS,
    LOG_DIR,
    MAX_UPLOAD_BYTES,
    RESULTADO_DIR,
    STATIC_DIR,
    TEMPLATE_DIR,
    UPLOAD_DIR,
)
from planilhas import PlanilhaInvalidaError

# --------------------------------------------------------------------------
# Logging (substitui os prints espalhados pelos módulos)
# --------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
    handlers=[
        RotatingFileHandler(LOG_DIR / "app.log", maxBytes=2_000_000, backupCount=3, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

# --------------------------------------------------------------------------
# App
# --------------------------------------------------------------------------

app = Flask(__name__, template_folder=str(TEMPLATE_DIR), static_folder=str(STATIC_DIR))

_chave = os.environ.get("SECRET_KEY")
if not _chave:
    if os.environ.get("FLASK_ENV") == "production":
        raise RuntimeError("Defina a variável de ambiente SECRET_KEY em produção.")
    _chave = uuid.uuid4().hex
    logger.warning("SECRET_KEY não definida; usando chave temporária (só para uso local).")

app.config.update(
    SECRET_KEY=_chave,
    MAX_CONTENT_LENGTH=MAX_UPLOAD_BYTES,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("FLASK_ENV") == "production",
)


def _carregar_usuarios() -> dict[str, str]:
    """
    Lê credenciais de USUARIOS no formato "usuario:hash,usuario2:hash".

    Gere um hash com:
        python -c "from werkzeug.security import generate_password_hash as g; print(g('sua-senha'))"
    """
    bruto = os.environ.get("USUARIOS", "").strip()
    if not bruto:
        senha = os.environ.get("SENHA_ADMIN")
        if senha:
            return {"admin": generate_password_hash(senha)}
        logger.warning(
            "Nenhum usuário configurado. Defina USUARIOS ou SENHA_ADMIN "
            "nas variáveis de ambiente."
        )
        return {}

    usuarios: dict[str, str] = {}
    for entrada in bruto.split(","):
        if ":" in entrada:
            nome, _, hash_senha = entrada.partition(":")
            usuarios[nome.strip()] = hash_senha.strip()
    return usuarios


USUARIOS = _carregar_usuarios()


def login_obrigatorio(view):
    @wraps(view)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))
        return view(*args, **kwargs)

    return wrapper


# --------------------------------------------------------------------------
# Upload
# --------------------------------------------------------------------------


class UploadInvalidoError(ValueError):
    pass


def salvar_upload(arquivo, destino: Path, rotulo: str, obrigatorio: bool = True) -> Path | None:
    if not arquivo or not arquivo.filename:
        if obrigatorio:
            raise UploadInvalidoError(f"O arquivo de {rotulo} é obrigatório.")
        return None

    nome = secure_filename(arquivo.filename)
    if not nome:
        raise UploadInvalidoError(f"Nome de arquivo inválido em {rotulo}.")

    if Path(nome).suffix.lower() not in EXTENSOES_PERMITIDAS:
        permitidas = ", ".join(sorted(EXTENSOES_PERMITIDAS))
        raise UploadInvalidoError(f"O arquivo de {rotulo} deve ser {permitidas}.")

    caminho = destino / nome
    arquivo.save(caminho)
    return caminho


class PastaTemporaria:
    """Isola os uploads de cada requisição (antes: nomes fixos compartilhados)."""

    def __enter__(self) -> Path:
        self._caminho = Path(tempfile.mkdtemp(dir=UPLOAD_DIR))
        return self._caminho

    def __exit__(self, *exc) -> None:
        shutil.rmtree(self._caminho, ignore_errors=True)


# --------------------------------------------------------------------------
# Rotas
# --------------------------------------------------------------------------


@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        senha = request.form.get("senha", "")

        hash_esperado = USUARIOS.get(usuario)
        if hash_esperado and check_password_hash(hash_esperado, senha):
            session.clear()
            session["usuario"] = usuario
            logger.info("Login efetuado: %s", usuario)
            return redirect(url_for("painel"))

        logger.warning("Tentativa de login falhou para o usuário %r", usuario)
        flash("Usuário ou senha inválidos.", "error")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/painel", methods=["GET"])
@login_obrigatorio
def painel():
    return render_template(
        "painel.html",
        dias_validade_padrao=DIAS_VALIDADE_PADRAO,
        dias_estoque_padrao=DIAS_ESTOQUE_PADRAO,
    )


@app.route("/download/<path:nome>")
@login_obrigatorio
def download(nome: str):
    """Serve os resultados de forma autenticada (antes: públicos em /static)."""
    seguro = secure_filename(nome)
    if not seguro or not (RESULTADO_DIR / seguro).is_file():
        abort(404)
    return send_from_directory(RESULTADO_DIR, seguro, as_attachment=True)


@app.route("/analise", methods=["POST"])
@login_obrigatorio
def analise():
    try:
        with PastaTemporaria() as pasta:
            dispensacao = salvar_upload(request.files.get("arquivo1"), pasta, "dispensação")
            distribuicao = salvar_upload(request.files.get("arquivo2"), pasta, "distribuição")
            estoque = salvar_upload(request.files.get("arquivo3"), pasta, "estoque")
            pedidos = salvar_upload(
                request.files.get("arquivo4"), pasta, "pedidos", obrigatorio=False
            )

            caminho = executar_analise(dispensacao, distribuicao, estoque, pedidos)

        flash("Análise concluída.", "success")
        session["ultimo_resultado"] = caminho.name
        return redirect(url_for("painel"))

    except (UploadInvalidoError, PlanilhaInvalidaError) as exc:
        flash(str(exc), "error")
    except Exception:
        logger.exception("Falha inesperada na análise principal")
        flash("Erro inesperado ao processar a análise. Consulte os logs.", "error")

    return redirect(url_for("painel"))


@app.route("/analise_compra", methods=["POST"])
@login_obrigatorio
def analise_compra():
    try:
        dias_validade = int(request.form.get("dias_validade") or DIAS_VALIDADE_PADRAO)
        dias_estoque = int(request.form.get("dias_estoque") or DIAS_ESTOQUE_PADRAO)
        if not (1 <= dias_validade <= 3650) or not (1 <= dias_estoque <= 3650):
            raise UploadInvalidoError("Os prazos devem estar entre 1 e 3650 dias.")

        with PastaTemporaria() as pasta:
            analise_path = salvar_upload(
                request.files.get("arquivo_compra"), pasta, "análise"
            )
            licitacao_path = salvar_upload(
                request.files.get("arquivo_licitacao"), pasta, "licitação", obrigatorio=False
            )

            resultado = executar_analise_compra(
                analise_path, dias_validade, dias_estoque, licitacao_path
            )

        for aviso in resultado.avisos:
            flash(aviso, "warning")

        flash("Relatório de compra gerado.", "success")
        session["ultimo_word"] = resultado.caminho_word.name
        session["ultimo_excel"] = resultado.caminho_excel.name
        return redirect(url_for("painel"))

    except (UploadInvalidoError, PlanilhaInvalidaError, ValueError) as exc:
        flash(str(exc), "error")
    except Exception:
        logger.exception("Falha inesperada no relatório de compra")
        flash("Erro inesperado ao gerar o relatório. Consulte os logs.", "error")

    return redirect(url_for("painel"))


@app.route("/analise2", methods=["POST"])
@login_obrigatorio
def analise_remume():
    """Cruzamento REMUME x estoque (antes: rota sem autenticação)."""
    try:
        with PastaTemporaria() as pasta:
            estoque = salvar_upload(request.files.get("arquivo_estoque"), pasta, "estoque")
            mapa = salvar_upload(
                request.files.get("arquivo_correspondencias"), pasta, "correspondências REMUME"
            )
            caminho = executar_analise_remume(estoque, mapa)

        flash("Cruzamento REMUME concluído.", "success")
        session["ultimo_remume"] = caminho.name
        return redirect(url_for("painel"))

    except (UploadInvalidoError, PlanilhaInvalidaError, ValueError) as exc:
        flash(str(exc), "error")
    except Exception:
        logger.exception("Falha inesperada no cruzamento REMUME")
        flash("Erro inesperado ao cruzar a REMUME. Consulte os logs.", "error")

    return redirect(url_for("painel"))


@app.route("/analise_orcamento", methods=["POST"])
@login_obrigatorio
def analise_orcamento():
    try:
        with PastaTemporaria() as pasta:
            dispensacao = salvar_upload(
                request.files.get("arquivo_disp_orcamento"), pasta, "dispensação"
            )
            distribuicao = salvar_upload(
                request.files.get("arquivo_dist_orcamento"), pasta, "distribuição"
            )
            caminho = executar_analise_orcamento(dispensacao, distribuicao)

        flash("Projeção orçamentária gerada.", "success")
        session["ultimo_orcamento"] = caminho.name
        return redirect(url_for("painel"))

    except (UploadInvalidoError, PlanilhaInvalidaError, ValueError) as exc:
        flash(str(exc), "error")
    except Exception:
        logger.exception("Falha inesperada na análise de orçamento")
        flash("Erro inesperado ao gerar a projeção. Consulte os logs.", "error")

    return redirect(url_for("painel"))


@app.errorhandler(413)
def upload_grande(_):
    limite = MAX_UPLOAD_BYTES // (1024 * 1024)
    flash(f"Arquivo acima do limite de {limite} MB.", "error")
    return redirect(url_for("painel"))


if __name__ == "__main__":
    # `app.run()` estava ausente: o processo abria o navegador e encerrava.
    porta = int(os.environ.get("PORT", 5000))
    if os.environ.get("ABRIR_NAVEGADOR", "1") == "1":
        import threading
        import webbrowser

        threading.Timer(1.5, lambda: webbrowser.open(f"http://127.0.0.1:{porta}")).start()

    app.run(host="127.0.0.1", port=porta, debug=False)
