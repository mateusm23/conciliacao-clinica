import io
import re
import pandas as pd

# Padrões de inferência de forma de pagamento para Santander
FORMA_PATTERNS = [
    (r'PIX RECEBIDO|PIX ENVIADO|TRANSFER.*PIX|DEP.*PIX', 'Pix'),
    (r'TARIFA.*PIX|PIX.*TARIFA|TARIFA AVULSA ENVIO PIX', 'Tarifa PIX'),
    (r'PAGAMENTO CARTAO DE DEBITO|GETNET|ANTECIPACAO GETNET', 'Cartão Débito / Antecipação'),
    (r'PAGAMENTO.*BOLETO|BOLETO', 'Boleto'),
    (r'PAGAMENTO DARF|DARF|TRIBUTOS FEDERAI', 'DARF / Imposto'),
    (r'TAR LIQ COB|TARIFA MANUTENCAO|TARIFA PAGAMENTO', 'Tarifa Bancária'),
    (r'CR COB BLOQ|RECEBIMENTO', 'Cobrança / Recebimento'),
    (r'TED|DOC', 'TED/DOC'),
]


def _infer_forma(desc: str) -> str:
    if not isinstance(desc, str):
        return 'Outros'
    d = desc.upper().strip()
    for pattern, label in FORMA_PATTERNS:
        if re.search(pattern, d):
            return label
    return 'Outros'


def parse(content: bytes, filename: str) -> pd.DataFrame:
    filepath = io.BytesIO(content)

    # Santander: 2 linhas de metadado (AGENCIA/CONTA + linha vazia) antes do header real
    df = pd.read_excel(filepath, engine='openpyxl', skiprows=2, header=0)

    # Normaliza nomes de colunas (remove espaços)
    df.columns = [str(c).strip() for c in df.columns]

    # Acesso posicional para robustez contra variações de encoding
    date_col = df.columns[0]   # Data
    desc_col = df.columns[1]   # Histórico
    valor_col = df.columns[3]  # Valor (R$)
    saldo_col = df.columns[4]  # Saldo (R$)

    # Filtra linhas onde a data não é válida
    datas = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    df = df[datas.notna()].copy()
    datas = datas[datas.notna()]

    valor = pd.to_numeric(df[valor_col], errors='coerce')

    result = pd.DataFrame({
        'data': datas.values,
        'descricao': df[desc_col].astype(str).str.strip(),
        'tipo': valor.apply(lambda v: 'Entrada' if v >= 0 else 'Saída').values,
        'valor': valor.abs().values,
        'saldo': pd.to_numeric(df[saldo_col], errors='coerce').values,
        'banco_origem': 'Santander',
        'forma_pagamento': df[desc_col].apply(_infer_forma).values,
        'categoria': '',
        'apuracao': '',
        'observacao': '',
    })

    return result.dropna(subset=['data', 'valor'])
