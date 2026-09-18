import os
import random
import sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from CRUD import adicionar_RDC, adicionar_acao_corretiva, atualizar_acao_individual, conectar
from seed_teste import (
    nome_cliente, nome_pessoa, cpf_cnpj_aleatorio, telefone_aleatorio, cep_aleatorio,
    email_aleatorio, relato_aleatorio, salvar_foto_dummy,
    PRODUTOS, RUAS, CIDADES_UF, SETORES_PARTICIPANTE, ACOES_TITULOS
)

UNIDADE_MATRIZ = "Matriz"
UNIDADE_FILIAL = "Filial - Mogi"

CAUSAS_PRINCIPAIS = [
    ("Defeito de fabricacao", 35),
    ("Atraso na entrega", 25),
    ("Divergencia de quantidade", 15),
    ("Embalagem avariada", 12),
    ("Erro na nota fiscal", 8),
    ("Produto fora de especificacao", 5),
]

MESES_ANO_ATUAL = [(2026, m) for m in range(4, 10)]

FAIXA_FILIAL_POR_MES = (14, 28)
FAIXA_MATRIZ_POR_MES = (4, 10)


def causa_aleatoria():
    causas = [c for c, _ in CAUSAS_PRINCIPAIS]
    pesos = [p for _, p in CAUSAS_PRINCIPAIS]
    return random.choices(causas, weights=pesos, k=1)[0]


def data_e_criado_em(ano, mes):
    ultimo_dia = 28 if mes == 2 else 30
    dia = random.randint(1, ultimo_dia)
    hora = random.randint(7, 18)
    minuto = random.randint(0, 59)
    criado_em_data = date(ano, mes, dia)
    criado_em = f"{criado_em_data.isoformat()}T{hora:02d}:{minuto:02d}:00"

    data_ocorrencia = criado_em_data - timedelta(days=random.randint(0, 4))
    return data_ocorrencia.isoformat(), criado_em


def criar_rdc(numero, unidade, ano, mes):
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
    data_rdc, criado_em = data_e_criado_em(ano, mes)
    vendedor = nome_pessoa()
    quantidade_defeitos = str(random.randint(1, 25))
    participantes = f"{nome_pessoa()} - {random.choice(SETORES_PARTICIPANTE)}"
    if random.random() < 0.5:
        participantes += f", {nome_pessoa()} - {random.choice(SETORES_PARTICIPANTE)}"
    preferencia_cliente = random.choice(PRODUTOS)
    relato = relato_aleatorio(cliente)
    causa_principal = causa_aleatoria()

    tem_fotos = random.random() < 0.35
    fotos_flag = 1 if tem_fotos else 0

    sucesso, erro, novo_id = adicionar_RDC(
        numero, cliente, email, cpf_cnpj_aleatorio(), telefone, rua, cep,
        nota_fiscal, data_rdc, vendedor, fotos_flag, relato,
        quantidade_defeitos, participantes, preferencia_cliente,
        causa_principal, criado_em, unidade
    )

    if not sucesso:
        print(f"Falhou Nº {numero}: {erro}")
        return None

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

    return novo_id


def main():
    numeros_disponiveis = random.sample(range(20000, 29999), 400)
    indice_numero = 0
    criados = 0

    for ano, mes in MESES_ANO_ATUAL:
        qtd_filial = random.randint(*FAIXA_FILIAL_POR_MES)
        qtd_matriz = random.randint(*FAIXA_MATRIZ_POR_MES)

        for _ in range(qtd_filial):
            numero = str(numeros_disponiveis[indice_numero])
            indice_numero += 1
            if criar_rdc(numero, UNIDADE_FILIAL, ano, mes) is not None:
                criados += 1

        for _ in range(qtd_matriz):
            numero = str(numeros_disponiveis[indice_numero])
            indice_numero += 1
            if criar_rdc(numero, UNIDADE_MATRIZ, ano, mes) is not None:
                criados += 1

    print(f"\nTotal criado: {criados}")


if __name__ == "__main__":
    main()
