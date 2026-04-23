import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Paleta de cores
COR_HEADER_BG = "1B6B5A"      # verde água escuro
COR_HEADER_FG = "FFFFFF"      # branco
COR_ENTRADA_BG = "E6F4F1"     # verde água clarinho
COR_SAIDA_BG = "FDE8E8"       # vermelho rosado
COR_BORDA = "C5D8D4"

COLUNAS = [
    ("Data",               13),
    ("Descrição",          52),
    ("Tipo",               11),
    ("Valor (R$)",         15),
    ("Saldo (R$)",         15),
    ("Banco / Origem",     22),
    ("Forma de Pagamento", 22),
    ("Categoria",          18),
    ("Apuração",           15),
    ("Observação",         25),
]


def _borda():
    lado = Side(style="thin", color=COR_BORDA)
    return Border(left=lado, right=lado, top=lado, bottom=lado)


def generate_excel(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Conciliação"

    fill_header = PatternFill("solid", fgColor=COR_HEADER_BG)
    font_header = Font(bold=True, color=COR_HEADER_FG, name="Calibri", size=11)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)
    borda = _borda()

    # --- Cabeçalho ---
    for col_idx, (nome, largura) in enumerate(COLUNAS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=nome)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = align_center
        cell.border = borda
        ws.column_dimensions[get_column_letter(col_idx)].width = largura

    ws.row_dimensions[1].height = 24

    fill_entrada = PatternFill("solid", fgColor=COR_ENTRADA_BG)
    fill_saida = PatternFill("solid", fgColor=COR_SAIDA_BG)
    font_data = Font(name="Calibri", size=10)

    # --- Dados ---
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        fill = fill_entrada if row.tipo == "Entrada" else fill_saida

        data_str = row.data.strftime('%d/%m/%Y') if pd.notna(row.data) else ''
        valor = row.valor if pd.notna(row.valor) else None
        saldo = row.saldo if pd.notna(row.saldo) else None

        linha = [
            data_str,
            row.descricao,
            row.tipo,
            valor,
            saldo,
            row.banco_origem,
            row.forma_pagamento,
            row.categoria,
            row.apuracao,
            row.observacao,
        ]

        for col_idx, value in enumerate(linha, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.fill = fill
            cell.border = borda
            cell.font = font_data
            cell.alignment = align_left

            # Formatação monetária para Valor e Saldo
            if col_idx in (4, 5) and isinstance(value, (int, float)):
                cell.number_format = 'R$ #,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")

    # --- Filtro automático e linha congelada ---
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUNAS))}1"
    ws.freeze_panes = "A2"

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
