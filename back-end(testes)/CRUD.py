import sqlite3
import os
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'exemplo.db')

SETOR_PADRAO = "Padrão"

def conectar():
    conexao = sqlite3.connect(DB_PATH)
    conexao.execute("PRAGMA foreign_keys = ON")
    return conexao

def criar_tabela():
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS rdcs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        numero TEXT NOT NULL,
        cliente TEXT NOT NULL,
        email TEXT,
        cpf_cnpj TEXT NOT NULL,
        telefone TEXT,
        endereco TEXT,
        cep TEXT,
        nota_fiscal TEXT NOT NULL,
        data TEXT NOT NULL,
        vendedor TEXT,
        fotos INTEGER DEFAULT 0
    )""")

    colunas = [linha[1] for linha in cursor.execute("PRAGMA table_info(rdcs)").fetchall()]
    if "numero" not in colunas:
        cursor.execute("ALTER TABLE rdcs ADD COLUMN numero TEXT")
    if "relato" not in colunas:
        cursor.execute("ALTER TABLE rdcs ADD COLUMN relato TEXT DEFAULT ''")
    if "quantidade_defeitos" not in colunas:
        cursor.execute("ALTER TABLE rdcs ADD COLUMN quantidade_defeitos TEXT DEFAULT ''")
    if "participantes" not in colunas:
        cursor.execute("ALTER TABLE rdcs ADD COLUMN participantes TEXT DEFAULT ''")
    if "preferencia_cliente" not in colunas:
        cursor.execute("ALTER TABLE rdcs ADD COLUMN preferencia_cliente TEXT DEFAULT ''")
    if "causa_principal" not in colunas:
        cursor.execute("ALTER TABLE rdcs ADD COLUMN causa_principal TEXT DEFAULT ''")
    if "criado_em" not in colunas:
        cursor.execute("ALTER TABLE rdcs ADD COLUMN criado_em TEXT DEFAULT ''")
    if "unidade" not in colunas:
        cursor.execute("ALTER TABLE rdcs ADD COLUMN unidade TEXT DEFAULT ''")

    cursor.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_rdcs_numero ON rdcs(numero)")

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS fotos_anexos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rdc_id INTEGER NOT NULL,
        caminho TEXT NOT NULL,
        FOREIGN KEY (rdc_id) REFERENCES rdcs(id) ON DELETE CASCADE
    )""")

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT NOT NULL,
        senha_hash TEXT NOT NULL,
        setor TEXT NOT NULL DEFAULT 'Padrão'
    )""")

    cursor.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_usuarios_usuario ON usuarios(usuario)")

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS acoes_corretivas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rdc_id INTEGER NOT NULL,
        titulo TEXT NOT NULL DEFAULT '',
        descricao TEXT NOT NULL DEFAULT '',
        concluida INTEGER DEFAULT 0,
        FOREIGN KEY (rdc_id) REFERENCES rdcs(id) ON DELETE CASCADE
    )""")

    colunas_acoes = [linha[1] for linha in cursor.execute("PRAGMA table_info(acoes_corretivas)").fetchall()]
    if "titulo" not in colunas_acoes:
        cursor.execute("ALTER TABLE acoes_corretivas ADD COLUMN titulo TEXT NOT NULL DEFAULT ''")

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS notas_fiscais_itens (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        num_docto TEXT NOT NULL,
        num_docto_norm INTEGER NOT NULL,
        serie TEXT,
        cliente_codigo TEXT,
        loja TEXT,
        cliente_nome TEXT,
        dt_emissao TEXT,
        tipo_nota TEXT,
        produto_codigo TEXT,
        descricao TEXT,
        quantidade REAL,
        valor_unitario REAL,
        valor_mercadoria REAL,
        armaz TEXT,
        cfo TEXT,
        tes TEXT,
        pedido TEXT,
        it TEXT,
        valor_ipi REAL,
        valor_icms REAL,
        valor_iss REAL,
        desp_acessorias REAL,
        total REAL
    )""")

    cursor.execute("CREATE INDEX IF NOT EXISTS idx_notas_fiscais_itens_norm ON notas_fiscais_itens(num_docto_norm)")

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS produtos_defeituosos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rdc_id INTEGER NOT NULL,
        origem TEXT NOT NULL DEFAULT 'manual',
        codigo TEXT NOT NULL DEFAULT '',
        descricao TEXT NOT NULL DEFAULT '',
        quantidade REAL NOT NULL DEFAULT 0,
        FOREIGN KEY (rdc_id) REFERENCES rdcs(id) ON DELETE CASCADE
    )""")

    conexao.commit()
    cursor.close()

def adicionar_RDC(numero, cliente, email, cpf_cnpj, telefone, endereco, cep, nota_fiscal, data, vendedor, fotos,
                   relato="", quantidade_defeitos=0, participantes="", preferencia_cliente="",
                   causa_principal="", criado_em="", unidade=""):
    conexao = conectar()
    cursor = conexao.cursor()

    sql = """
    INSERT INTO rdcs
    (numero, cliente, email, cpf_cnpj, telefone, endereco, cep, nota_fiscal, data, vendedor, fotos, relato,
     quantidade_defeitos, participantes, preferencia_cliente, causa_principal, criado_em, unidade)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    valores = (
        numero, cliente, email, cpf_cnpj, telefone, endereco, cep, nota_fiscal, data, vendedor, fotos, relato,
        quantidade_defeitos, participantes, preferencia_cliente, causa_principal, criado_em, unidade
    )

    sucesso = False
    erro = None
    novo_id = None

    try:
        cursor.execute(sql, valores)
        conexao.commit()
        sucesso = True
        novo_id = cursor.lastrowid

        print("RDC inserida com sucesso")
        print("ID:", cursor.lastrowid)

    except sqlite3.IntegrityError:
        conexao.rollback()
        erro = f"Já existe uma RDC cadastrada com o Nº \"{numero}\"."

    except sqlite3.Error as e:
        conexao.rollback()
        erro = str(e)
        print(e)

    finally:
        conexao.close()

    return sucesso, erro, novo_id

def listar_RDC():
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
    SELECT rdcs.*, COUNT(fotos_anexos.id), GROUP_CONCAT(fotos_anexos.caminho)
    FROM rdcs
    LEFT JOIN fotos_anexos ON fotos_anexos.rdc_id = rdcs.id
    GROUP BY rdcs.id
    ORDER BY rdcs.id
    """)
    rdcs = cursor.fetchall()
    conexao.close()
    return rdcs

def adicionar_foto(rdc_id, caminho):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute(
        "INSERT INTO fotos_anexos (rdc_id, caminho) VALUES (?, ?)",
        (rdc_id, caminho)
    )
    conexao.commit()
    conexao.close()

def listar_fotos(rdc_id):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("SELECT id, caminho FROM fotos_anexos WHERE rdc_id = ?", (rdc_id,))
    fotos = cursor.fetchall()
    conexao.close()
    return fotos

def buscar_foto(foto_id):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("SELECT rdc_id, caminho FROM fotos_anexos WHERE id = ?", (foto_id,))
    foto = cursor.fetchone()
    conexao.close()
    return foto

def remover_foto(foto_id):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("DELETE FROM fotos_anexos WHERE id = ?", (foto_id,))
    conexao.commit()
    conexao.close()

def buscar_RDC(id):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("SELECT * FROM rdcs WHERE id = ?", (id,))
    rdc = cursor.fetchone()
    conexao.close()
    return rdc

def deletar_RDC(id):
    fotos = listar_fotos(id)
    for _, caminho in fotos:
        caminho_absoluto = os.path.join(BASE_DIR, 'static', caminho)
        if os.path.exists(caminho_absoluto):
            os.remove(caminho_absoluto)

    conexao = conectar()
    try:
        cursor = conexao.cursor()
        cursor.execute("DELETE FROM rdcs WHERE id = ?", (id,))
        conexao.commit()
    finally:
        conexao.close()

def editar_RDC(id, numero, cliente, email, cpf_cnpj, telefone, endereco, cep, nota_fiscal, data, vendedor, fotos, relato,
                quantidade_defeitos, participantes, preferencia_cliente, causa_principal, unidade=""):
    conexao = conectar()
    cursor = conexao.cursor()

    sucesso = False
    erro = None

    try:
        cursor.execute("""
        UPDATE rdcs
        SET numero = ?, cliente = ?, email = ?, cpf_cnpj = ?, telefone = ?, endereco = ?, cep = ?, nota_fiscal = ?, data = ?, vendedor = ?, fotos = ?, relato = ?,
            quantidade_defeitos = ?, participantes = ?, preferencia_cliente = ?, causa_principal = ?, unidade = ?
        WHERE id = ?
        """, (
            numero, cliente, email, cpf_cnpj, telefone, endereco, cep, nota_fiscal, data, vendedor, fotos, relato,
            quantidade_defeitos, participantes, preferencia_cliente, causa_principal, unidade, id
        ))
        conexao.commit()
        sucesso = True

    except sqlite3.IntegrityError:
        conexao.rollback()
        erro = f"Já existe uma RDC cadastrada com o Nº \"{numero}\"."

    except sqlite3.Error as e:
        conexao.rollback()
        erro = str(e)

    finally:
        conexao.close()

    return sucesso, erro

def criar_usuario(usuario, senha, setor=SETOR_PADRAO):
    conexao = conectar()
    cursor = conexao.cursor()

    senha_hash = generate_password_hash(senha)

    sucesso = False
    erro = None

    try:
        cursor.execute(
            "INSERT INTO usuarios (usuario, senha_hash, setor) VALUES (?, ?, ?)",
            (usuario, senha_hash, setor)
        )
        conexao.commit()
        sucesso = True

    except sqlite3.IntegrityError:
        conexao.rollback()
        erro = f"Já existe um usuário com o nome \"{usuario}\"."

    except sqlite3.Error as e:
        conexao.rollback()
        erro = str(e)

    finally:
        conexao.close()

    return sucesso, erro

def buscar_usuario(usuario):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute(
        "SELECT id, usuario, senha_hash, setor FROM usuarios WHERE usuario = ?",
        (usuario,)
    )
    linha = cursor.fetchone()
    conexao.close()
    return linha

def listar_RDC_qualidade():
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
    SELECT rdcs.*,
           COUNT(DISTINCT fotos_anexos.id),
           GROUP_CONCAT(DISTINCT fotos_anexos.caminho),
           COUNT(DISTINCT acoes_corretivas.id),
           COUNT(DISTINCT CASE WHEN acoes_corretivas.concluida = 1 THEN acoes_corretivas.id END)
    FROM rdcs
    LEFT JOIN fotos_anexos ON fotos_anexos.rdc_id = rdcs.id
    LEFT JOIN acoes_corretivas ON acoes_corretivas.rdc_id = rdcs.id
    GROUP BY rdcs.id
    ORDER BY rdcs.id
    """)
    rdcs = cursor.fetchall()
    conexao.close()
    return rdcs

def adicionar_acao_corretiva(rdc_id, titulo, descricao):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute(
        "INSERT INTO acoes_corretivas (rdc_id, titulo, descricao) VALUES (?, ?, ?)",
        (rdc_id, titulo, descricao)
    )
    conexao.commit()
    conexao.close()

def listar_acoes_corretivas(rdc_id):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute(
        "SELECT id, titulo, descricao, concluida FROM acoes_corretivas WHERE rdc_id = ? ORDER BY id",
        (rdc_id,)
    )
    acoes = cursor.fetchall()
    conexao.close()
    return acoes

def atualizar_acao_individual(acao_id, rdc_id, concluida):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute(
        "UPDATE acoes_corretivas SET concluida = ? WHERE id = ? AND rdc_id = ?",
        (concluida, acao_id, rdc_id)
    )
    conexao.commit()
    conexao.close()

def editar_acao_corretiva(acao_id, rdc_id, titulo, descricao):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute(
        "UPDATE acoes_corretivas SET titulo = ?, descricao = ? WHERE id = ? AND rdc_id = ?",
        (titulo, descricao, acao_id, rdc_id)
    )
    conexao.commit()
    conexao.close()

def remover_acao_corretiva(acao_id, rdc_id):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute(
        "DELETE FROM acoes_corretivas WHERE id = ? AND rdc_id = ?",
        (acao_id, rdc_id)
    )
    conexao.commit()
    conexao.close()

def buscar_itens_nota_fiscal(numero):
    digitos = "".join(caractere for caractere in str(numero or "") if caractere.isdigit())
    if not digitos:
        return None

    numero_normalizado = int(digitos)

    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
    SELECT cliente_nome, produto_codigo, descricao, quantidade
    FROM notas_fiscais_itens
    WHERE num_docto_norm = ?
    ORDER BY id
    """, (numero_normalizado,))
    linhas = cursor.fetchall()
    conexao.close()

    if not linhas:
        return None

    return {
        "cliente": linhas[0][0],
        "itens": [
            {"codigo": codigo, "descricao": descricao, "quantidade": quantidade}
            for _, codigo, descricao, quantidade in linhas
        ]
    }

def listar_catalogo_produtos():
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
    SELECT DISTINCT produto_codigo, descricao
    FROM notas_fiscais_itens
    WHERE produto_codigo IS NOT NULL AND produto_codigo != ''
    ORDER BY descricao
    """)
    resultado = cursor.fetchall()
    conexao.close()
    return resultado

def _linhas_produtos_defeituosos(rdc_id, itens):
    return [
        (
            rdc_id,
            (item.get("origem") or "manual"),
            (item.get("codigo") or "").strip(),
            (item.get("descricao") or "").strip(),
            item.get("quantidade") or 0
        )
        for item in itens
        if (item.get("codigo") or "").strip() or (item.get("descricao") or "").strip()
    ]

def adicionar_produtos_defeituosos(rdc_id, itens):
    linhas = _linhas_produtos_defeituosos(rdc_id, itens)
    if not linhas:
        return

    conexao = conectar()
    cursor = conexao.cursor()
    cursor.executemany(
        "INSERT INTO produtos_defeituosos (rdc_id, origem, codigo, descricao, quantidade) VALUES (?, ?, ?, ?, ?)",
        linhas
    )
    conexao.commit()
    conexao.close()

def listar_produtos_defeituosos(rdc_id):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute(
        "SELECT id, origem, codigo, descricao, quantidade FROM produtos_defeituosos WHERE rdc_id = ? ORDER BY id",
        (rdc_id,)
    )
    resultado = cursor.fetchall()
    conexao.close()
    return resultado

def substituir_produtos_defeituosos(rdc_id, itens):
    linhas = _linhas_produtos_defeituosos(rdc_id, itens)

    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("DELETE FROM produtos_defeituosos WHERE rdc_id = ?", (rdc_id,))
    if linhas:
        cursor.executemany(
            "INSERT INTO produtos_defeituosos (rdc_id, origem, codigo, descricao, quantidade) VALUES (?, ?, ?, ?, ?)",
            linhas
        )
    conexao.commit()
    conexao.close()

def listar_causas_distintas():
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
    SELECT causa_principal, COUNT(*) as total
    FROM rdcs
    WHERE causa_principal IS NOT NULL AND causa_principal != ''
    GROUP BY causa_principal
    ORDER BY total DESC, causa_principal ASC
    """)
    causas = [linha[0] for linha in cursor.fetchall()]
    conexao.close()
    return causas

def contar_rdc_por_causa():
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
    SELECT causa_principal, COUNT(*) as total
    FROM rdcs
    WHERE causa_principal IS NOT NULL AND causa_principal != ''
    GROUP BY causa_principal
    ORDER BY total DESC, causa_principal ASC
    """)
    resultado = cursor.fetchall()
    conexao.close()
    return resultado

def contar_rdc_sem_causa():
    conexao = conectar()
    cursor = conexao.cursor()
    total = cursor.execute("""
    SELECT COUNT(*) FROM rdcs WHERE causa_principal IS NULL OR causa_principal = ''
    """).fetchone()[0]
    conexao.close()
    return total

def contar_rdc_por_ano():
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
    SELECT strftime('%Y', CASE WHEN criado_em IS NOT NULL AND criado_em != '' THEN criado_em ELSE data END) as ano,
           COUNT(*) as total
    FROM rdcs
    GROUP BY ano
    ORDER BY ano DESC
    """)
    resultado = cursor.fetchall()
    conexao.close()
    return resultado

def listar_reclamacoes_por_mes(unidade):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
    SELECT strftime('%Y-%m', CASE WHEN criado_em IS NOT NULL AND criado_em != '' THEN criado_em ELSE data END) as mes,
           COUNT(*) as total
    FROM rdcs
    WHERE unidade = ?
    GROUP BY mes
    ORDER BY mes ASC
    """, (unidade,))
    resultado = cursor.fetchall()
    conexao.close()
    return resultado

def listar_causas_por_mes(unidade):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
    SELECT strftime('%Y-%m', CASE WHEN criado_em IS NOT NULL AND criado_em != '' THEN criado_em ELSE data END) as mes,
           causa_principal, COUNT(*) as total
    FROM rdcs
    WHERE unidade = ? AND causa_principal IS NOT NULL AND causa_principal != ''
    GROUP BY mes, causa_principal
    ORDER BY mes ASC, causa_principal ASC
    """, (unidade,))
    resultado = cursor.fetchall()
    conexao.close()
    return resultado

def contar_causas_unidade(unidade):
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
    SELECT causa_principal, COUNT(*) as total
    FROM rdcs
    WHERE unidade = ? AND causa_principal IS NOT NULL AND causa_principal != ''
    GROUP BY causa_principal
    ORDER BY total DESC, causa_principal ASC
    """, (unidade,))
    resultado = cursor.fetchall()
    conexao.close()
    return resultado

def contar_causas_unidades(unidades):
    conexao = conectar()
    cursor = conexao.cursor()
    marcadores = ",".join("?" for _ in unidades)
    cursor.execute(f"""
    SELECT causa_principal, COUNT(*) as total
    FROM rdcs
    WHERE unidade IN ({marcadores}) AND causa_principal IS NOT NULL AND causa_principal != ''
    GROUP BY causa_principal
    ORDER BY total DESC, causa_principal ASC
    """, tuple(unidades))
    resultado = cursor.fetchall()
    conexao.close()
    return resultado

def listar_causas_por_ano_unidades(unidades):
    conexao = conectar()
    cursor = conexao.cursor()
    marcadores = ",".join("?" for _ in unidades)
    cursor.execute(f"""
    SELECT strftime('%Y', CASE WHEN criado_em IS NOT NULL AND criado_em != '' THEN criado_em ELSE data END) as ano,
           causa_principal, COUNT(*) as total
    FROM rdcs
    WHERE unidade IN ({marcadores}) AND causa_principal IS NOT NULL AND causa_principal != ''
    GROUP BY ano, causa_principal
    ORDER BY ano ASC, causa_principal ASC
    """, tuple(unidades))
    resultado = cursor.fetchall()
    conexao.close()
    return resultado

def contar_rdc_por_ano_unidades(unidades):
    conexao = conectar()
    cursor = conexao.cursor()
    marcadores = ",".join("?" for _ in unidades)
    cursor.execute(f"""
    SELECT strftime('%Y', CASE WHEN criado_em IS NOT NULL AND criado_em != '' THEN criado_em ELSE data END) as ano,
           COUNT(*) as total
    FROM rdcs
    WHERE unidade IN ({marcadores})
    GROUP BY ano
    ORDER BY ano ASC
    """, tuple(unidades))
    resultado = cursor.fetchall()
    conexao.close()
    return resultado

def autenticar_usuario(usuario, senha):
    linha = buscar_usuario(usuario)

    if linha is None:
        return None

    _, _, senha_hash, _ = linha

    if not check_password_hash(senha_hash, senha):
        return None

    return linha