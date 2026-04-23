import io
from typing import List

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse

from parsers.detector import detect_and_parse
from pipeline import merge_and_process
from exporter import generate_excel

app = FastAPI(title="Conciliação Bancária — Clínica Meu Médico")

HTML = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Conciliação Bancária</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            background: #f0f7f5;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
        }
        .card {
            background: #ffffff;
            border-radius: 12px;
            padding: 40px 48px;
            width: 100%;
            max-width: 520px;
            box-shadow: 0 4px 24px rgba(27,107,90,0.10);
        }
        h1 { color: #1B6B5A; font-size: 22px; margin-bottom: 6px; }
        .subtitle { color: #6b8f89; font-size: 14px; margin-bottom: 32px; }
        .file-group { margin-bottom: 18px; }
        label { display: block; font-size: 13px; color: #2c5f58; font-weight: 600; margin-bottom: 6px; }
        input[type=file] {
            width: 100%;
            padding: 10px;
            border: 1.5px dashed #9ecfc7;
            border-radius: 8px;
            background: #f7fcfb;
            font-size: 13px;
            color: #3a7a71;
            cursor: pointer;
        }
        input[type=file]:hover { border-color: #1B6B5A; }
        button {
            width: 100%;
            padding: 14px;
            background: #1B6B5A;
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 15px;
            font-weight: 600;
            cursor: pointer;
            margin-top: 10px;
            transition: background 0.2s;
        }
        button:hover { background: #145247; }
        button:disabled { background: #9ecfc7; cursor: not-allowed; }
        #status {
            margin-top: 18px;
            padding: 12px 16px;
            border-radius: 8px;
            font-size: 13px;
            display: none;
        }
        .info  { background: #e6f4f1; color: #1B6B5A; }
        .error { background: #fde8e8; color: #c0392b; }
    </style>
</head>
<body>
    <div class="card">
        <h1>Conciliação Bancária</h1>
        <p class="subtitle">Clínica Meu Médico — Gerador de Planilha Financeira</p>

        <form id="form">
            <div class="file-group">
                <label>Extrato Bancário 1</label>
                <input type="file" name="arquivos" accept=".xls,.xlsx">
            </div>
            <div class="file-group">
                <label>Extrato Bancário 2</label>
                <input type="file" name="arquivos" accept=".xls,.xlsx">
            </div>
            <div class="file-group">
                <label>Extrato Bancário 3</label>
                <input type="file" name="arquivos" accept=".xls,.xlsx">
            </div>
            <button type="submit" id="btn">Gerar Planilha Excel</button>
        </form>
        <div id="status"></div>
    </div>

    <script>
        document.getElementById('form').addEventListener('submit', async function(e) {
            e.preventDefault();
            const btn = document.getElementById('btn');
            const status = document.getElementById('status');
            const inputs = document.querySelectorAll('input[type=file]');

            const formData = new FormData();
            let count = 0;
            inputs.forEach(input => {
                if (input.files[0]) {
                    formData.append('arquivos', input.files[0]);
                    count++;
                }
            });

            if (count === 0) {
                status.className = 'error';
                status.textContent = 'Selecione ao menos um arquivo.';
                status.style.display = 'block';
                return;
            }

            btn.disabled = true;
            btn.textContent = 'Processando...';
            status.className = 'info';
            status.textContent = 'Aguarde, processando os extratos...';
            status.style.display = 'block';

            try {
                const response = await fetch('/processar', { method: 'POST', body: formData });
                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.detail || 'Erro ao processar.');
                }
                const blob = await response.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'conciliacao_bancaria.xlsx';
                a.click();
                URL.revokeObjectURL(url);
                status.textContent = 'Planilha gerada com sucesso! O download iniciou automaticamente.';
            } catch(err) {
                status.className = 'error';
                status.textContent = 'Erro: ' + err.message;
            } finally {
                btn.disabled = false;
                btn.textContent = 'Gerar Planilha Excel';
            }
        });
    </script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
def index():
    return HTML


@app.post("/processar")
async def processar(arquivos: List[UploadFile] = File(...)):
    dfs = []
    erros = []

    for arquivo in arquivos:
        if not arquivo.filename:
            continue
        content = await arquivo.read()
        if not content:
            continue
        try:
            df = detect_and_parse(content, arquivo.filename)
            if df is not None and not df.empty:
                dfs.append(df)
        except Exception as e:
            erros.append(f"{arquivo.filename}: {e}")

    if not dfs:
        detalhe = "Nenhum arquivo válido processado."
        if erros:
            detalhe += " Erros: " + " | ".join(erros)
        raise HTTPException(status_code=422, detail=detalhe)

    df_final = merge_and_process(dfs)
    excel_bytes = generate_excel(df_final)

    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=conciliacao_bancaria.xlsx"},
    )
