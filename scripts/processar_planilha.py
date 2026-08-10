import os
import json
import datetime
from pathlib import Path
import openpyxl

def clean_value(val):
    if val is None:
        return None
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, str):
        return val.strip()
    return val

def processar_planilha():
    base_dir = Path(__file__).resolve().parent.parent
    data_dir = base_dir / "data"
    scripts_dir = base_dir / "scripts"

    # Procurar arquivos .xlsx na pasta data
    excel_files = list(data_dir.glob("*.xlsx")) + list(data_dir.glob("*.XLSX"))
    if not excel_files:
        raise FileNotFoundError("Nenhum arquivo Excel (.xlsx) foi encontrado na pasta 'data'.")

    file_path = excel_files[0]
    print(f"Lendo arquivo: {file_path}")

    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        print("Planilha vazia.")
        return

    headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(rows[0])]

    # Mapeamento para nomes de chaves amigáveis/normalizadas (snake_case)
    key_mapping = {
        "N° Título": "numero_titulo",
        "Tipo Título": "tipo_titulo",
        "Transação": "transacao",
        "Data Emissão": "data_emissao",
        "Data Entrada": "data_entrada",
        "Cliente": "cliente",
        "Observação": "observacao",
        "Projeto": "projeto",
        "Fase": "fase",
        "Conta Financeira": "conta_financeira",
        "Centro de Custo": "centro_custo",
        "Data de Vencimento": "data_vencimento",
        "Valor Original": "valor_original"
    }

    dados_tratados = []

    for row in rows[1:]:
        # Ignorar linhas totalmente vazias
        if not any(cell is not None for cell in row):
            continue

        item = {}
        item_original = {}

        for header, cell_val in zip(headers, row):
            val = clean_value(cell_val)
            normalized_key = key_mapping.get(header, header.lower().replace(" ", "_"))
            item[normalized_key] = val
            item_original[header] = val

        dados_tratados.append(item)

    output_json_path = data_dir / "titulos_tratados.json"
    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(dados_tratados, f, ensure_ascii=False, indent=2)

    print(f"Arquivo JSON gerado com sucesso em: {output_json_path}")
    print(f"Total de registros processados: {len(dados_tratados)}")

if __name__ == "__main__":
    processar_planilha()
