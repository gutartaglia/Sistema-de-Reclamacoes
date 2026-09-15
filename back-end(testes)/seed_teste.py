import os
import random
import struct
import uuid
import zlib
from datetime import date, timedelta

from CRUD import (
    adicionar_RDC, adicionar_foto, adicionar_acao_corretiva,
    atualizar_acao_individual, conectar
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTA_UPLOADS = os.path.join(BASE_DIR, "static", "uploads")
os.makedirs(PASTA_UPLOADS, exist_ok=True)

PRIMEIROS_NOMES = [
    "Ana", "Bruno", "Carla", "Daniel", "Eduarda", "Felipe", "Gabriela", "Henrique",
    "Isabela", "Joao", "Karina", "Lucas", "Mariana", "Nicolas", "Olivia", "Pedro",
    "Queila", "Rafael", "Sofia", "Thiago", "Ursula", "Vinicius", "Wesley", "Ximena",
    "Yasmin", "Zeca", "Camila", "Diego", "Elisa", "Fabio"
]
SOBRENOMES = [
    "Silva", "Santos", "Oliveira", "Souza", "Rodrigues", "Ferreira", "Alves",
    "Pereira", "Lima", "Gomes", "Costa", "Ribeiro", "Martins", "Carvalho",
    "Almeida", "Lopes", "Soares", "Fernandes", "Vieira", "Barbosa"
]
EMPRESAS = [
    "Flora Confeccoes", "TechSol Industria", "Nordeste Uniformes", "Alfa Distribuidora",
    "Vega Comercio", "Prisma Textil", "Horizonte Log", "Bravo Suprimentos",
    "Delta Embalagens", "Omega Papelaria", "Sul Componentes", "Norte Ferragens",
    "Central Quimica", "Atlas Metalurgica", "Cristal Cosmeticos"
]
SUFIXOS_EMPRESA = ["LTDA", "LTDA-EPP", "S/A", "ME", "EIRELI"]
PRODUTOS = [
    "NONTECK DRT30", "Tecido Oxford", "Caixa Papelao 40x30", "Etiqueta Adesiva",
    "Fita Isolante", "Luva Nitrilica", "Mascara PFF2", "Rolo de Plastico Bolha",
    "Fardo de Sacaria", "Bobina Termica", "Parafuso Sextavado", "Chapa de Aco",
    "Frasco Plastico 500ml", "Tinta Acrilica", "Solvente Industrial"
]
CIDADES_UF = [
    "Sao Paulo/SP", "Brasilia/DF", "Curitiba/PR", "Salvador/BA", "Recife/PE",
    "Fortaleza/CE", "Porto Alegre/RS", "Belo Horizonte/MG", "Manaus/AM", "Goiania/GO",
    "Campinas/SP", "Florianopolis/SC", "Vitoria/ES", "Natal/RN", "Joao Pessoa/PB"
]
RUAS = ["Rua das Flores", "Av. Industrial", "Rua Sete de Setembro", "Av. Brasil",
        "Rua XV de Novembro", "Alameda Santos", "Rua do Comercio", "Av. Paulista"]
SETORES_PARTICIPANTE = ["Qualidade", "Vendedor", "Logistica", "Expedicao", "Atendimento", "Compras"]
MOTIVOS = [
    "produto chegou com avaria visivel na embalagem",
    "quantidade recebida divergente do pedido",
    "material com espessura abaixo do especificado",
    "atraso na entrega acima do prazo combinado",
    "nota fiscal com dados incorretos",
    "produto trocado por engano no momento da expedicao",
    "defeito de fabricacao identificado apos o uso",
    "cliente relatou cheiro forte incomum no material",
    "lote com coloracao diferente do padrao aprovado",
    "embalagem violada no transporte"
]
ACOES_TITULOS = [
    "Revisar processo de embalagem", "Treinar equipe de expedicao",
    "Contatar fornecedor sobre o lote", "Ajustar conferencia de pedidos",
    "Atualizar ficha tecnica do produto", "Inspecionar lote remanescente em estoque",
    "Revisar checklist de qualidade", "Reforcar treinamento de manuseio"
]


def nome_pessoa():
    return f"{random.choice(PRIMEIROS_NOMES)} {random.choice(SOBRENOMES)}"


def nome_cliente():
    if random.random() < 0.6:
        return f"{random.choice(EMPRESAS)} {random.choice(SUFIXOS_EMPRESA)}"
    return nome_pessoa()


def cpf_cnpj_aleatorio():
    if random.random() < 0.5:
        n = [random.randint(0, 9) for _ in range(11)]
        return f"{n[0]}{n[1]}{n[2]}.{n[3]}{n[4]}{n[5]}.{n[6]}{n[7]}{n[8]}-{n[9]}{n[10]}"
    n = [random.randint(0, 9) for _ in range(14)]
    return f"{n[0]}{n[1]}.{n[2]}{n[3]}{n[4]}.{n[5]}{n[6]}{n[7]}/{n[8]}{n[9]}{n[10]}{n[11]}-{n[12]}{n[13]}"


def telefone_aleatorio():
    ddd = random.randint(11, 99)
    numero = random.randint(90000000, 99999999)
    return f"({ddd}) 9{str(numero)[:4]}-{str(numero)[4:]}"


def cep_aleatorio():
    return f"{random.randint(10000, 99999)}-{random.randint(100, 999)}"


def data_aleatoria():
    inicio = date(2025, 1, 1)
    dias = random.randint(0, 610)
    return (inicio + timedelta(days=dias)).isoformat()


def email_aleatorio(nome):
    dominio = random.choice(["gmail.com", "hotmail.com", "outlook.com", "empresa.com.br", "uol.com.br"])
    base = "".join(ch for ch in nome.lower() if ch.isalnum() or ch == " ").replace(" ", ".")
    return f"{base}{random.randint(1, 999)}@{dominio}"


def relato_aleatorio(cliente):
    motivo = random.choice(MOTIVOS)
    texto = f"Cliente {cliente} relatou que o {motivo}."
    if random.random() < 0.4:
        texto = texto.replace(motivo, f"<mark>{motivo}</mark>")
    if random.random() < 0.4:
        texto += " <b>Necessita retorno urgente do setor de qualidade.</b>"
    return texto


def png_1x1(cor):
    def chunk(tipo, dados):
        return (struct.pack(">I", len(dados)) + tipo + dados +
                struct.pack(">I", zlib.crc32(tipo + dados) & 0xffffffff))

    assinatura = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    raw = b"\x00" + bytes(cor)
    idat = zlib.compress(raw)
    return assinatura + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")


def salvar_foto_dummy(rdc_id):
    cor = (random.randint(0, 255), random.randint(0, 255), random.randint(0, 255))
    conteudo = png_1x1(cor)
    nome_unico = f"{uuid.uuid4().hex}.png"
    with open(os.path.join(PASTA_UPLOADS, nome_unico), "wb") as arquivo:
        arquivo.write(conteudo)
    adicionar_foto(rdc_id, f"uploads/{nome_unico}")


def main():
    numeros_unicos = random.sample(range(300, 9999), 100)
    criados = 0

    for numero in numeros_unicos:
        cliente = nome_cliente()
        emails = [email_aleatorio(cliente)]
        if random.random() < 0.3:
            emails.append(email_aleatorio(cliente))
        email = "; ".join(emails)

        telefone = telefone_aleatorio() if random.random() < 0.85 else ""
        if not email and not telefone:
            telefone = telefone_aleatorio()

        rua = f"{random.choice(RUAS)}, {random.randint(10, 2500)} - {random.choice(CIDADES_UF)}"
        cep = cep_aleatorio()
        nota_fiscal = str(random.randint(10000, 999999))
        data_rdc = data_aleatoria()
        vendedor = nome_pessoa()
        quantidade_defeitos = random.randint(0, 25)
        participantes = f"{nome_pessoa()} - {random.choice(SETORES_PARTICIPANTE)}"
        if random.random() < 0.5:
            participantes += f", {nome_pessoa()} - {random.choice(SETORES_PARTICIPANTE)}"
        preferencia_cliente = random.choice(PRODUTOS)
        relato = relato_aleatorio(cliente)

        tem_fotos = random.random() < 0.35
        fotos_flag = 1 if tem_fotos else 0

        sucesso, erro, novo_id = adicionar_RDC(
            str(numero), cliente, email, cpf_cnpj_aleatorio(), telefone, rua, cep,
            nota_fiscal, data_rdc, vendedor, fotos_flag, relato,
            quantidade_defeitos, participantes, preferencia_cliente
        )

        if not sucesso:
            print(f"Falhou Nº {numero}: {erro}")
            continue

        criados += 1

        if tem_fotos:
            for _ in range(random.randint(1, 3)):
                salvar_foto_dummy(novo_id)

        if random.random() < 0.4:
            qtd_acoes = random.randint(1, 4)
            for _ in range(qtd_acoes):
                titulo = random.choice(ACOES_TITULOS)
                descricao = f"Acao referente a ocorrencia Nº {numero}: {titulo.lower()}."
                adicionar_acao_corretiva(novo_id, titulo, descricao)

            conexao = conectar()
            cursor = conexao.cursor()
            ids_acoes = [linha[0] for linha in cursor.execute(
                "SELECT id FROM acoes_corretivas WHERE rdc_id = ?", (novo_id,)
            ).fetchall()]
            conexao.close()

            estado = random.random()
            if estado < 0.3:
                marcar = []
            elif estado < 0.7:
                marcar = random.sample(ids_acoes, k=random.randint(1, len(ids_acoes) - 1) if len(ids_acoes) > 1 else 0)
            else:
                marcar = ids_acoes

            for acao_id in marcar:
                atualizar_acao_individual(acao_id, novo_id, 1)

    print(f"\nTotal criado: {criados}/100")


if __name__ == "__main__":
    main()