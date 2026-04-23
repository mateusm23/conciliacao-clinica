import pandas as pd


def merge_and_process(dfs: list) -> pd.DataFrame:
    df = pd.concat(dfs, ignore_index=True)

    # Ordena por data decrescente (mais recente primeiro)
    df = df.sort_values('data', ascending=False).reset_index(drop=True)

    # Sinaliza possíveis duplicatas (mesmo banco + data + valor + início da descrição)
    dup_key = (
        df['banco_origem'].str.upper() + '|' +
        df['data'].astype(str) + '|' +
        df['valor'].round(2).astype(str) + '|' +
        df['descricao'].str[:30].str.upper().str.strip()
    )
    dup_mask = dup_key.duplicated(keep='first')
    df.loc[dup_mask, 'observacao'] = 'Possível duplicata'

    return df
