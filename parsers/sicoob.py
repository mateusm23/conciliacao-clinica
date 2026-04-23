import io
import re
import pandas as pd
from pathlib import Path

# Mapeamento do campo "Tipo" do Sicoob para Forma de Pagamento
TIPO_FORMA_MAP = [
    (r'transfer[eê]ncia de pix|pix enviado', 'Pix'),
    (r'dep[oó]sito de pix|pix recebido', 'Pix'),
    (r'tarifa de entrada pix|tarifa de pix', 'Tarifa PIX'),
    (r'antecipa[cç][aã]o de receb[ií]veis|antecipa[cç][aã]o', 'Antecipação Cartão'),
    (r'liquida[cç][aã]o de cart[aã]o de d[eé]bito', 'Cartão Débito'),
    (r'pagamento de boleto|boleto', 'Boleto'),
    (r'dep[oó]sito', 'Depósito/TED'),
    (r'tarifa', 'Tarifa Bancária'),
    (r'ted|doc', 'TED/DOC'),
]


def _infer_forma(tipo_str: str) -> str:
    if not isinstance(tipo_str, str):
        return 'Outros'
    t = tipo_str.lower().strip()
    for pattern, label in TIPO_FORMA_MAP:
        if re.search(pattern, t):
            return label
    return 'Outros'


def parse(content: bytes, filename: str) -> pd.DataFrame:
    filepath = io.BytesIO(content)
    df = pd.read_excel(filepath, engine='xlrd')

    # Coluna D/C está na posição 5 — filtra linhas sem ela (ex: "Saldo do dia")
    dc_col = df.iloc[:, 5]
    df = df[dc_col.notna()].copy()

    date_raw = df.iloc[:, 0]
    tipo_raw = df.iloc[:, 2]
    desc_raw = df.iloc[:, 3]
    valor_raw = df.iloc[:, 4]
    dc_raw = df.iloc[:, 5]
    saldo_raw = df.iloc[:, 6]

    result = pd.DataFrame({
        'data': pd.to_datetime(date_raw, dayfirst=True, errors='coerce'),
        'descricao': desc_raw.astype(str).str.strip(),
        'tipo': dc_raw.map({'C': 'Entrada', 'D': 'Saída'}),
        'valor': pd.to_numeric(valor_raw, errors='coerce').abs(),
        'saldo': pd.to_numeric(saldo_raw, errors='coerce'),
        'banco_origem': Path(filename).stem,
        'forma_pagamento': tipo_raw.apply(_infer_forma),
        'categoria': '',
        'apuracao': '',
        'observacao': '',
    })

    return result.dropna(subset=['data', 'valor'])
