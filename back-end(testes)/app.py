import os
import io
import json
import uuid
import secrets
import bleach
from functools import wraps
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file
from werkzeug.utils import secure_filename
from CRUD import (
    buscar_RDC, criar_tabela, adicionar_RDC, listar_RDC, deletar_RDC, editar_RDC,
    adicionar_foto, listar_fotos, listar_fotos_conteudo, buscar_foto, remover_foto, autenticar_usuario,
    SETOR_SAC, SETOR_QUALIDADE, listar_RDC_qualidade, adicionar_acao_corretiva, listar_acoes_corretivas,
    atualizar_acao_individual, editar_acao_corretiva, remover_acao_corretiva,
    listar_causas_distintas, contar_rdc_por_causa, contar_rdc_sem_causa, contar_rdc_por_ano,
    listar_reclamacoes_por_mes, listar_causas_por_mes, contar_causas_unidade,
    contar_causas_unidades, listar_causas_por_ano_unidades, contar_rdc_por_ano_unidades,
    buscar_itens_nota_fiscal, listar_catalogo_produtos, adicionar_produtos_defeituosos,
    listar_produtos_defeituosos, substituir_produtos_defeituosos
)
from gerar_documento import gerar_ficha_reclamacao
from gerar_excel import gerar_excel_unidade, gerar_excel_combinado

app = Flask(__name__)

EXTENSOES_PERMITIDAS = {"png", "jpg", "jpeg", "gif", "webp"}

TAGS_RELATO_PERMITIDAS = ["b", "strong", "i", "em", "mark", "br", "p", "div"]

def sanitizar_relato(html):
    return bleach.clean(html or "", tags=TAGS_RELATO_PERMITIDAS, attributes={}, strip=True)

UNIDADE_MATRIZ = "Matriz"
UNIDADE_FILIAL = "Filial - Mogi"

HISTORICO_RDC_POR_ANO = {
    "2023": 80, "2024": 112, "2025": 158
}

NOMES_MESES_ABREV = {
    "01": "Jan", "02": "Fev", "03": "Mar", "04": "Abr", "05": "Mai", "06": "Jun",
    "07": "Jul", "08": "Ago", "09": "Set", "10": "Out", "11": "Nov", "12": "Dez"
}

def formatar_mes_label(mes_iso):
    if not mes_iso or "-" not in mes_iso:
        return mes_iso or ""
    ano, mes = mes_iso.split("-")
    return f"{NOMES_MESES_ABREV.get(mes, mes)}/{ano}"

def montar_dados_mensais(unidade):
    linhas = listar_reclamacoes_por_mes(unidade)
    return [
        {
            "mes": mes,
            "mes_label": formatar_mes_label(mes),
            "total": total,
            "indice": round(total / 2000, 4)
        }
        for mes, total in linhas
    ]

def montar_causas_por_mes(unidade):
    linhas = listar_causas_por_mes(unidade)

    meses = []
    causas = []
    matriz = {}

    for mes, causa, total in linhas:
        if mes not in meses:
            meses.append(mes)
        if causa not in causas:
            causas.append(causa)
        matriz[(mes, causa)] = total

    causas.sort()

    tabela = [
        {
            "mes": mes,
            "mes_label": formatar_mes_label(mes),
            "valores": [matriz.get((mes, causa), 0) for causa in causas]
        }
        for mes in meses
    ]

    return meses, causas, tabela

def montar_pareto_a_partir_de_linhas(linhas):
    total_geral = sum(total for _, total in linhas)

    resultado = []
    freq_acumulada = 0

    for causa, total in linhas:
        freq_acumulada += total
        percentual = (total / total_geral * 100) if total_geral else 0
        percentual_acumulado = (freq_acumulada / total_geral * 100) if total_geral else 0
        resultado.append({
            "causa": causa,
            "frequencia": total,
            "frequencia_acumulada": freq_acumulada,
            "percentual": round(percentual, 2),
            "percentual_acumulado": round(percentual_acumulado, 2)
        })

    return resultado

def montar_pareto(unidade):
    return montar_pareto_a_partir_de_linhas(contar_causas_unidade(unidade))

def montar_pareto_unidades(unidades):
    return montar_pareto_a_partir_de_linhas(contar_causas_unidades(unidades))

def montar_rdc_por_ano_completo():
    dados = dict(HISTORICO_RDC_POR_ANO)
    for ano, total in contar_rdc_por_ano():
        if ano and ano not in dados:
            dados[ano] = total
    return sorted(dados.items())

def montar_top_causas_por_ano(unidades, limite=4):
    causas_top = [causa for causa, _ in contar_causas_unidades(unidades)[:limite]]

    linhas = listar_causas_por_ano_unidades(unidades)
    anos = []
    matriz = {}

    for ano, causa, total in linhas:
        if causa not in causas_top:
            continue
        if ano not in anos:
            anos.append(ano)
        matriz[(ano, causa)] = total

    anos.sort()

    tabela = [
        {
            "ano": ano,
            "valores": [matriz.get((ano, causa), 0) for causa in causas_top]
        }
        for ano in anos
    ]

    return causas_top, tabela

def montar_rdc_por_ano_completo_unidades(unidades):
    dados = dict(HISTORICO_RDC_POR_ANO)
    for ano, total in contar_rdc_por_ano_unidades(unidades):
        if ano and ano not in dados:
            dados[ano] = total
    return sorted(dados.items())

def montar_pareto_por_ano_unidades(unidades):
    linhas = listar_causas_por_ano_unidades(unidades)
    por_ano = {}

    for ano, causa, total in linhas:
        por_ano.setdefault(ano, []).append((causa, total))

    resultado = {}
    for ano, pares in por_ano.items():
        pares_ordenados = sorted(pares, key=lambda item: (-item[1], item[0]))
        resultado[ano] = montar_pareto_a_partir_de_linhas(pares_ordenados)

    return resultado

def montar_mensal_combinado():
    filial = {item["mes"]: item["total"] for item in montar_dados_mensais(UNIDADE_FILIAL)}
    matriz = {item["mes"]: item["total"] for item in montar_dados_mensais(UNIDADE_MATRIZ)}
    meses = sorted(set(filial.keys()) | set(matriz.keys()))

    return [
        {
            "mes": mes,
            "mes_label": formatar_mes_label(mes),
            "filial": filial.get(mes, 0),
            "matriz": matriz.get(mes, 0),
            "total": filial.get(mes, 0) + matriz.get(mes, 0)
        }
        for mes in meses
    ]

def carregar_js_inline(caminho_relativo):
    caminho = os.path.join(app.root_path, "static", caminho_relativo)
    with open(caminho, "r", encoding="utf-8") as arquivo:
        conteudo = arquivo.read()
    return conteudo.replace("</script", "<\\/script")

MASCARAS_JS = carregar_js_inline("js/mascaras.txt")
LUCIDE_JS = carregar_js_inline("vendor/lucide/lucide.min.txt")
CHARTJS_JS = carregar_js_inline("vendor/chartjs/chart.umd.txt")
CHARTJS_DATALABELS_JS = carregar_js_inline("vendor/chartjs/chartjs-plugin-datalabels.txt")
PRODUTOS_DEFEITUOSOS_JS = carregar_js_inline("js/produtos_defeituosos.txt")

def montar_catalogo_produtos_json():
    return json.dumps([
        {"codigo": codigo, "descricao": descricao}
        for codigo, descricao in listar_catalogo_produtos()
    ], ensure_ascii=False)

@app.context_processor
def injetar_scripts_inline():
    return dict(
        MASCARAS_JS=MASCARAS_JS, LUCIDE_JS=LUCIDE_JS,
        CHARTJS_JS=CHARTJS_JS, CHARTJS_DATALABELS_JS=CHARTJS_DATALABELS_JS,
        PRODUTOS_DEFEITUOSOS_JS=PRODUTOS_DEFEITUOSOS_JS,
        CATALOGO_PRODUTOS_JSON=CATALOGO_PRODUTOS_JSON
    )

CHAVE_SECRETA_PATH = os.environ.get("SECRET_KEY_PATH") or os.path.join(app.root_path, ".secret_key")
if os.path.exists(CHAVE_SECRETA_PATH):
    with open(CHAVE_SECRETA_PATH, "r") as arquivo_chave:
        app.secret_key = arquivo_chave.read().strip()
else:
    nova_chave = secrets.token_hex(32)
    with open(CHAVE_SECRETA_PATH, "w") as arquivo_chave:
        arquivo_chave.write(nova_chave)
    app.secret_key = nova_chave

criar_tabela()

CATALOGO_PRODUTOS_JSON = montar_catalogo_produtos_json()

def login_required(f):
    @wraps(f)
    def decorada(*args, **kwargs):
        if "usuario_id" not in session:
            return redirect(url_for("login", proximo=request.path))
        return f(*args, **kwargs)
    return decorada

def setor_required(setor):
    def decorador(f):
        @wraps(f)
        def decorada(*args, **kwargs):
            if session.get("usuario_setor") != setor:
                return "Acesso restrito ao setor responsável por este painel.", 403
            return f(*args, **kwargs)
        return decorada
    return decorador

def extensao_permitida(nome_arquivo):
    return (
        "." in nome_arquivo
        and nome_arquivo.rsplit(".", 1)[1].lower() in EXTENSOES_PERMITIDAS
    )

def salvar_anexos(rdc_id, arquivos):
    for arquivo in arquivos:
        if not arquivo or not arquivo.filename:
            continue

        if not extensao_permitida(arquivo.filename):
            continue

        nome_arquivo = secure_filename(arquivo.filename) or f"{uuid.uuid4().hex}"
        mime_type = arquivo.mimetype or "application/octet-stream"
        conteudo = arquivo.read()
        adicionar_foto(rdc_id, nome_arquivo, mime_type, conteudo)

def anexos_validos(arquivos):
    return [a for a in arquivos if a and a.filename and extensao_permitida(a.filename)]

def existe_arquivo_invalido(arquivos):
    return any(a and a.filename and not extensao_permitida(a.filename) for a in arquivos)

def coletar_emails(formulario_form):
    emails = []
    for email in formulario_form.getlist("email"):
        email = (email or "").strip()
        if email and email not in emails:
            emails.append(email)
    return emails

def coletar_produtos_defeituosos(formulario_form):
    bruto = formulario_form.get("produtos_defeituosos_json") or "[]"

    try:
        itens = json.loads(bruto)
    except (TypeError, ValueError):
        return []

    if not isinstance(itens, list):
        return []

    resultado = []
    for item in itens:
        if not isinstance(item, dict):
            continue

        codigo = str(item.get("codigo") or "").strip()
        descricao = str(item.get("descricao") or "").strip()
        if not codigo and not descricao:
            continue

        try:
            quantidade = float(item.get("quantidade") or 0)
        except (TypeError, ValueError):
            quantidade = 0

        origem = item.get("origem") if item.get("origem") in ("banco", "manual") else "manual"
        resultado.append({"origem": origem, "codigo": codigo, "descricao": descricao, "quantidade": quantidade})

    return resultado

def montar_produtos_defeituosos_json(rdc_id):
    return json.dumps([
        {"origem": origem, "codigo": codigo, "descricao": descricao, "quantidade": quantidade}
        for _, origem, codigo, descricao, quantidade in listar_produtos_defeituosos(rdc_id)
    ], ensure_ascii=False)

def resumir_produtos_defeituosos(itens):
    partes = []
    for item in itens:
        quantidade = item["quantidade"]
        quantidade_texto = str(int(quantidade)) if float(quantidade).is_integer() else str(quantidade)
        rotulo = " — ".join(parte for parte in (item["codigo"], item["descricao"]) if parte)
        partes.append(f"{quantidade_texto}x {rotulo}".strip())
    return "; ".join(partes)

def processar_edicao_rdc(id):
    """Processa os campos comuns de edição de uma RDC a partir de request.form/files.

    Retorna (sucesso, erro, fotos_atuais). Quando sucesso é True, erro é None e
    fotos_atuais é None (não é mais necessária pelo chamador)."""

    numero = request.form.get("numero")
    cliente = request.form.get("cliente")
    emails = coletar_emails(request.form)
    email = "; ".join(emails)
    cpf_cnpj = request.form.get("cpf_cnpj")
    telefone = (request.form.get("telefone") or "").strip()
    endereco = request.form.get("endereco")
    cep = request.form.get("cep")
    nota_fiscal = request.form.get("nota_fiscal")
    data = request.form.get("data")
    vendedor = request.form.get("vendedor")
    relato = sanitizar_relato(request.form.get("relato"))
    produtos_defeituosos = coletar_produtos_defeituosos(request.form)
    quantidade_defeitos = resumir_produtos_defeituosos(produtos_defeituosos)
    participantes = request.form.get("participantes")
    preferencia_cliente = request.form.get("preferencia_cliente")
    causa_principal = (request.form.get("causa_principal") or "").strip()
    unidade = (request.form.get("unidade") or "").strip()

    if not emails and not telefone:
        return False, "Informe pelo menos um e-mail ou um telefone.", listar_fotos(id)

    anexos = request.files.getlist("anexos")

    if existe_arquivo_invalido(anexos):
        return False, "Só é permitido anexar arquivos de imagem (PNG, JPG, JPEG, GIF ou WEBP).", listar_fotos(id)

    fotos_atuais = listar_fotos(id)

    ids_remover = set()
    for foto_id_str in request.form.getlist("remover_fotos"):
        try:
            foto_id = int(foto_id_str)
        except (TypeError, ValueError):
            continue
        if any(fid == foto_id for fid, _ in fotos_atuais):
            ids_remover.add(foto_id)

    validos = anexos_validos(anexos)
    fotos_restantes = len(fotos_atuais) - len(ids_remover) + len(validos)
    fotos = 1 if fotos_restantes > 0 else 0

    sucesso, erro = editar_RDC(
        id,
        numero,
        cliente,
        email,
        cpf_cnpj,
        telefone,
        endereco,
        cep,
        nota_fiscal,
        data,
        vendedor,
        fotos,
        relato,
        quantidade_defeitos,
        participantes,
        preferencia_cliente,
        causa_principal,
        unidade
    )

    if not sucesso:
        return False, erro, fotos_atuais

    for foto_id in ids_remover:
        remover_foto(foto_id)

    salvar_anexos(id, validos)
    substituir_produtos_defeituosos(id, produtos_defeituosos)

    return True, None, None

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        usuario = request.form.get("usuario")
        senha = request.form.get("senha")

        linha = autenticar_usuario(usuario, senha)

        if linha is None:
            return render_template(
                "login.html",
                erro="Usuário ou senha inválidos.",
                usuario_preenchido=usuario
            )

        session.clear()
        session["usuario_id"] = linha[0]
        session["usuario_nome"] = linha[1]
        session["usuario_setor"] = linha[3]

        destino_padrao = url_for("painel_qualidade") if linha[3] == SETOR_QUALIDADE else url_for("lista")
        proximo = request.args.get("proximo") or destino_padrao
        return redirect(proximo)

    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/formulario", methods=["GET", "POST"])
@login_required
def formulario():

    if request.method == "POST":

        eh_ajax = request.headers.get("X-Requested-With") == "XMLHttpRequest"

        numero = request.form.get("numero")
        cliente = request.form.get("cliente")
        emails = coletar_emails(request.form)
        email = "; ".join(emails)
        cpf_cnpj = request.form.get("cpf_cnpj")
        telefone = (request.form.get("telefone") or "").strip()
        endereco = request.form.get("endereco")
        cep = request.form.get("cep")
        nota_fiscal = request.form.get("nota_fiscal")
        data = request.form.get("data")
        vendedor = request.form.get("vendedor")
        relato = sanitizar_relato(request.form.get("relato"))
        produtos_defeituosos = coletar_produtos_defeituosos(request.form)
        quantidade_defeitos = resumir_produtos_defeituosos(produtos_defeituosos)
        participantes = request.form.get("participantes")
        preferencia_cliente = request.form.get("preferencia_cliente")
        causa_principal = (request.form.get("causa_principal") or "").strip()
        unidade = (request.form.get("unidade") or "").strip()

        if not emails and not telefone:
            erro = "Informe pelo menos um e-mail ou um telefone."
            if eh_ajax:
                return jsonify(sucesso=False, erro=erro), 400
            return render_template("formulario.html", erro=erro, dados=request.form, relato=relato, causas=listar_causas_distintas())

        anexos = request.files.getlist("anexos")

        if existe_arquivo_invalido(anexos):
            erro = "Só é permitido anexar arquivos de imagem (PNG, JPG, JPEG, GIF ou WEBP)."
            if eh_ajax:
                return jsonify(sucesso=False, erro=erro), 400
            return render_template("formulario.html", erro=erro, dados=request.form, relato=relato, causas=listar_causas_distintas())

        validos = anexos_validos(anexos)
        fotos = 1 if validos else 0
        criado_em = datetime.now().isoformat()

        sucesso, erro, novo_id = adicionar_RDC(
            numero,
            cliente,
            email,
            cpf_cnpj,
            telefone,
            endereco,
            cep,
            nota_fiscal,
            data,
            vendedor,
            fotos,
            relato,
            quantidade_defeitos,
            participantes,
            preferencia_cliente,
            causa_principal,
            criado_em,
            unidade
        )

        if not sucesso:
            if eh_ajax:
                return jsonify(sucesso=False, erro=erro), 400
            return render_template("formulario.html", erro=erro, dados=request.form, relato=relato, causas=listar_causas_distintas())

        salvar_anexos(novo_id, validos)
        adicionar_produtos_defeituosos(novo_id, produtos_defeituosos)

        if eh_ajax:
            return jsonify(sucesso=True, id=novo_id)

        return redirect("/formulario")

    return render_template("formulario.html", causas=listar_causas_distintas())

@app.route("/api/nota-fiscal/<numero>")
@login_required
def api_nota_fiscal(numero):
    resultado = buscar_itens_nota_fiscal(numero)

    if resultado is None:
        return jsonify(encontrado=False)

    return jsonify(encontrado=True, cliente=resultado["cliente"], itens=resultado["itens"])

@app.route("/foto/<foto_id>")
@login_required
def servir_foto(foto_id):
    foto = buscar_foto(foto_id)
    if foto is None or not foto[3]:
        return "Foto não encontrada", 404

    _, nome_arquivo, mime_type, conteudo = foto
    return send_file(
        io.BytesIO(conteudo),
        mimetype=mime_type or "application/octet-stream",
        download_name=nome_arquivo or "foto",
    )

@app.route("/lista")
@login_required
@setor_required(SETOR_SAC)
def lista():
    rdcs = listar_RDC()
    return render_template("lista.html", rdcs=rdcs, causas=listar_causas_distintas())

@app.route('/visualizar/<int:id>', methods=["GET", "POST"])
@login_required
def visualizar_rdc(id):

    if request.method == "POST":

        sucesso, erro, fotos_atuais = processar_edicao_rdc(id)

        if not sucesso:
            rdc = buscar_RDC(id)
            return render_template(
                "visualizar.html",
                rdc=rdc,
                editando=True,
                erro=erro,
                fotos=fotos_atuais if fotos_atuais is not None else listar_fotos(id),
                acoes=listar_acoes_corretivas(id),
                causas=listar_causas_distintas(),
                produtos_defeituosos_json=montar_produtos_defeituosos_json(id)
            )

        return redirect(url_for("visualizar_rdc", id=id))

    rdc = buscar_RDC(id)

    return render_template(
        "visualizar.html",
        rdc=rdc,
        editando=False,
        fotos=listar_fotos(id),
        acoes=listar_acoes_corretivas(id),
        causas=listar_causas_distintas(),
        produtos_defeituosos_json=montar_produtos_defeituosos_json(id)
    )



@app.route("/painel-qualidade")
@login_required
@setor_required(SETOR_QUALIDADE)
def painel_qualidade():
    rdcs = listar_RDC_qualidade()

    filial_mensal = montar_dados_mensais(UNIDADE_FILIAL)
    filial_causas_meses, filial_causas_lista, filial_causas_tabela = montar_causas_por_mes(UNIDADE_FILIAL)
    filial_pareto = montar_pareto(UNIDADE_FILIAL)
    filial_pareto_por_ano = montar_pareto_por_ano_unidades([UNIDADE_FILIAL])
    filial_anos_disponiveis = sorted(
        set(filial_pareto_por_ano.keys()) | {item["mes"][:4] for item in filial_mensal},
        reverse=True
    )

    matriz_mensal = montar_dados_mensais(UNIDADE_MATRIZ)
    matriz_causas_meses, matriz_causas_lista, matriz_causas_tabela = montar_causas_por_mes(UNIDADE_MATRIZ)
    matriz_pareto = montar_pareto(UNIDADE_MATRIZ)
    matriz_pareto_por_ano = montar_pareto_por_ano_unidades([UNIDADE_MATRIZ])
    matriz_anos_disponiveis = sorted(
        set(matriz_pareto_por_ano.keys()) | {item["mes"][:4] for item in matriz_mensal},
        reverse=True
    )

    unidades_combinado = [UNIDADE_FILIAL, UNIDADE_MATRIZ]
    combinado_pareto = montar_pareto_unidades(unidades_combinado)
    combinado_pareto_por_ano = montar_pareto_por_ano_unidades(unidades_combinado)
    combinado_rdc_por_ano = montar_rdc_por_ano_completo_unidades(unidades_combinado)
    combinado_mensal = montar_mensal_combinado()
    combinado_anos_disponiveis = sorted(
        set(combinado_pareto_por_ano.keys()) | {item["mes"][:4] for item in combinado_mensal},
        reverse=True
    )

    return render_template(
        "painel_qualidade.html",
        rdcs=rdcs,
        causas=listar_causas_distintas(),
        causas_contagem=contar_rdc_por_causa(),
        sem_causa=contar_rdc_sem_causa(),
        anos_contagem=contar_rdc_por_ano(),
        filial_mensal=filial_mensal,
        filial_causas_lista=filial_causas_lista,
        filial_causas_tabela=filial_causas_tabela,
        filial_pareto=filial_pareto,
        filial_pareto_por_ano=filial_pareto_por_ano,
        filial_anos_disponiveis=filial_anos_disponiveis,
        matriz_mensal=matriz_mensal,
        matriz_causas_lista=matriz_causas_lista,
        matriz_causas_tabela=matriz_causas_tabela,
        matriz_pareto=matriz_pareto,
        matriz_pareto_por_ano=matriz_pareto_por_ano,
        matriz_anos_disponiveis=matriz_anos_disponiveis,
        combinado_pareto=combinado_pareto,
        combinado_pareto_por_ano=combinado_pareto_por_ano,
        combinado_rdc_por_ano=combinado_rdc_por_ano,
        combinado_mensal=combinado_mensal,
        combinado_anos_disponiveis=combinado_anos_disponiveis,
        rdc_por_ano_completo=montar_rdc_por_ano_completo()
    )

@app.route("/painel-qualidade/exportar-excel/<aba>")
@login_required
@setor_required(SETOR_QUALIDADE)
def exportar_dashboard_excel(aba):
    ano = request.args.get("ano") or "todos"

    if aba == "filial":
        unidade = UNIDADE_FILIAL
        nome_unidade = "Filial"
    elif aba == "matriz":
        unidade = UNIDADE_MATRIZ
        nome_unidade = "Matriz"
    elif aba != "combinado":
        return "Aba inválida", 404

    if aba in ("filial", "matriz"):
        mensal = montar_dados_mensais(unidade)
        _, causas_lista, causas_tabela = montar_causas_por_mes(unidade)
        pareto_por_ano = montar_pareto_por_ano_unidades([unidade])

        if ano != "todos":
            mensal = [item for item in mensal if item["mes"].startswith(ano)]
            causas_tabela = [linha for linha in causas_tabela if linha["mes"].startswith(ano)]
            pareto = pareto_por_ano.get(ano, [])
        else:
            pareto = montar_pareto(unidade)

        workbook = gerar_excel_unidade(
            nome_unidade, mensal, causas_lista, causas_tabela, pareto, montar_rdc_por_ano_completo(),
            ano_filtro=ano
        )
    else:
        unidades = [UNIDADE_FILIAL, UNIDADE_MATRIZ]
        causas_top, causas_ano_tabela = montar_top_causas_por_ano(unidades, limite=4)
        pareto_por_ano = montar_pareto_por_ano_unidades(unidades)
        mensal = montar_mensal_combinado()

        if ano != "todos":
            mensal = [item for item in mensal if item["mes"].startswith(ano)]
            pareto = pareto_por_ano.get(ano, [])
        else:
            pareto = montar_pareto_unidades(unidades)

        workbook = gerar_excel_combinado(
            causas_top, causas_ano_tabela, pareto,
            montar_rdc_por_ano_completo_unidades(unidades), mensal,
            ano_filtro=ano
        )

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    nome_arquivo = f"dashboard_{aba}_{ano}.xlsx"

    return send_file(
        buffer,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/painel-qualidade/<int:id>", methods=["GET", "POST"])
@login_required
@setor_required(SETOR_QUALIDADE)
def painel_qualidade_rdc(id):

    if request.method == "POST":

        sucesso, erro, fotos_atuais = processar_edicao_rdc(id)

        if not sucesso:
            rdc = buscar_RDC(id)
            return render_template(
                "painel_qualidade_rdc.html",
                rdc=rdc,
                editando=True,
                erro=erro,
                fotos=fotos_atuais if fotos_atuais is not None else listar_fotos(id),
                acoes=listar_acoes_corretivas(id),
                causas=listar_causas_distintas(),
                produtos_defeituosos_json=montar_produtos_defeituosos_json(id)
            )

        return redirect(url_for("painel_qualidade_rdc", id=id))

    rdc = buscar_RDC(id)

    if rdc is None:
        return "Rdc não encontrada", 404

    return render_template(
        "painel_qualidade_rdc.html",
        rdc=rdc,
        editando=False,
        fotos=listar_fotos(id),
        acoes=listar_acoes_corretivas(id),
        causas=listar_causas_distintas(),
        produtos_defeituosos_json=montar_produtos_defeituosos_json(id)
    )

@app.route("/painel-qualidade/<int:id>/deletar")
@login_required
@setor_required(SETOR_QUALIDADE)
def painel_qualidade_deletar_rdc(id):
    rdc = buscar_RDC(id)

    if rdc is None:
        return "Rdc não encontrada", 404

    deletar_RDC(id)

    return redirect(url_for("painel_qualidade"))

@app.route("/painel-qualidade/<int:rdc_id>/acao/<int:acao_id>/marcar", methods=["POST"])
@login_required
@setor_required(SETOR_QUALIDADE)
def marcar_acao(rdc_id, acao_id):
    concluida = 1 if request.form.get("concluida") == "1" else 0
    atualizar_acao_individual(acao_id, rdc_id, concluida)
    return redirect(url_for("painel_qualidade_rdc", id=rdc_id) + "#acoes")

@app.route("/painel-qualidade/<int:rdc_id>/acao/adicionar", methods=["POST"])
@login_required
@setor_required(SETOR_QUALIDADE)
def adicionar_acao(rdc_id):
    titulo = (request.form.get("titulo") or "").strip()
    descricao = (request.form.get("descricao") or "").strip()

    if titulo:
        adicionar_acao_corretiva(rdc_id, titulo, descricao)

    return redirect(url_for("painel_qualidade_rdc", id=rdc_id) + "#acoes")

@app.route("/painel-qualidade/<int:rdc_id>/acao/<int:acao_id>/editar", methods=["POST"])
@login_required
@setor_required(SETOR_QUALIDADE)
def editar_acao(rdc_id, acao_id):
    titulo = (request.form.get("titulo") or "").strip()
    descricao = (request.form.get("descricao") or "").strip()

    if titulo:
        editar_acao_corretiva(acao_id, rdc_id, titulo, descricao)

    return redirect(url_for("painel_qualidade_rdc", id=rdc_id) + "#acoes")

@app.route("/painel-qualidade/<int:rdc_id>/acao/<int:acao_id>/excluir", methods=["POST"])
@login_required
@setor_required(SETOR_QUALIDADE)
def excluir_acao(rdc_id, acao_id):
    remover_acao_corretiva(acao_id, rdc_id)
    return redirect(url_for("painel_qualidade_rdc", id=rdc_id) + "#acoes")

@app.route('/exportar/<int:id>')
@login_required
def exportar_rdc(id):
    rdc = buscar_RDC(id)

    if rdc is None:
        return "Rdc não encontrada", 404

    documento = gerar_ficha_reclamacao(
        rdc, listar_fotos_conteudo(id), listar_acoes_corretivas(id), listar_produtos_defeituosos(id)
    )
    nome_arquivo = f"RDC_{rdc[1]}.docx"

    return send_file(
        documento,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.route('/delete/<int:id>')
@login_required
def deletar_rdc(id):
    rdc = buscar_RDC(id)

    if rdc is None:
        return "Rdc não encontrada", 404

    deletar_RDC(id)

    return redirect ('/lista')

if __name__ == "__main__":
    app.run(debug=True)