import io
import pandas as pd
from pathlib import Path
from . import sicoob, santander


def detect_and_parse(content: bytes, filename: str) -> pd.DataFrame:
    ext = Path(filename).suffix.lower()

    if ext == '.xlsx':
        peek = io.BytesIO(content)
        df_raw = pd.read_excel(peek, header=None, nrows=3, engine='openpyxl')
        first_cell = str(df_raw.iloc[0, 0]).strip().upper()
        if first_cell == 'AGENCIA':
            return santander.parse(content, filename)

    elif ext == '.xls':
        peek = io.BytesIO(content)
        df_check = pd.read_excel(peek, nrows=1, engine='xlrd')
        if 'D/C' in df_check.columns:
            return sicoob.parse(content, filename)

    # Fallback: tenta Sicoob (formato mais comum dos arquivos do projeto)
    try:
        return sicoob.parse(content, filename)
    except Exception:
        return santander.parse(content, filename)
