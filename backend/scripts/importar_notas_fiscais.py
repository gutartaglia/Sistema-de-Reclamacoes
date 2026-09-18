"""Importa o faturamento (aba "2-Analitico - Itens") de notas fiscais/matr550.xlsx
para a tabela notas_fiscais_itens do banco local. Pode ser rodado de novo a qualquer
momento para recarregar os dados a partir de uma planilha atualizada (apaga e reinsere)."""

import os
import sys

import openpyxl

from CRUD import conectar, criar_tabela

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_PLANILHA = os.path.join(BASE_DIR, "notas fiscais", "matr550.xlsx")
ABA_ITENS = "2-Analitico - Itens"
PRIMEIRA_LINHA_DADOS = 3


def _texto(valor):
    if valor is None:
        return None
    texto = str(valor).strip()
    return texto or None


def _numero(valor):
    if valor is None or valor == "":
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _data(valor):
    if valor is None:
        return None
    if hasattr(valor, "isoformat"):
        return valor.isoformat()
    return str(valor)


def _linhas_da_planilha():
    workbook = openpyxl.load_workbook(CAMINHO_PLANILHA, data_only=True)
    planilha = workbook[ABA_ITENS]

    for linha in planilha.iter_rows(min_row=PRIMEIRA_LINHA_DADOS, values_only=True):
        num_docto = _texto(linha[0])
        if not num_docto:
            continue

        digitos = "".join(c for c in num_docto if c.isdigit())
        if not digitos:
            continue

        yield (
            num_docto,
            int(digitos),
            _texto(linha[1]),
            _texto(linha[2]),
            _texto(linha[3]),
            _texto(linha[4]),
            _data(linha[5]),
            _texto(linha[6]),
            _texto(linha[7]),
            _texto(linha[8]),
            _numero(linha[9]),
            _numero(linha[10]),
            _numero(linha[11]),
            _texto(linha[12]),
            _texto(linha[13]),
            _texto(linha[14]),
            _texto(linha[15]),
            _texto(linha[16]),
            _numero(linha[17]),
            _numero(linha[18]),
            _numero(linha[19]),
            _numero(linha[20]),
            _numero(linha[21]),
        )


def importar():
    criar_tabela()

    linhas = list(_linhas_da_planilha())

    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("DELETE FROM notas_fiscais_itens")
    cursor.executemany("""
    INSERT INTO notas_fiscais_itens (
        num_docto, num_docto_norm, serie, cliente_codigo, loja, cliente_nome,
        dt_emissao, tipo_nota, produto_codigo, descricao, quantidade, valor_unitario,
        valor_mercadoria, armaz, cfo, tes, pedido, it, valor_ipi, valor_icms,
        valor_iss, desp_acessorias, total
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, linhas)
    conexao.commit()

    total_linhas = cursor.execute("SELECT COUNT(*) FROM notas_fiscais_itens").fetchone()[0]
    total_notas = cursor.execute("SELECT COUNT(DISTINCT num_docto_norm) FROM notas_fiscais_itens").fetchone()[0]
    conexao.close()

    return total_linhas, total_notas


if __name__ == "__main__":
    linhas, notas = importar()
    print(f"Importação concluída: {linhas} itens de {notas} notas fiscais distintas.")
    sys.exit(0)
