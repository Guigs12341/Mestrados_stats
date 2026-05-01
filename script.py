import pandas as pd
import re
import os

# Definições de ficheiros
files_to_process = {
    'Alameda': 'Data_Alameda.txt',
    'Tagus': 'Data_Tagus.txt'
}
output_file = 'Lista.xlsx'

def clean_value(val):
    if val:
        return val.strip().replace('\u200b', '').replace('\xa0', ' ')
    return ""

def parse_data_robust(file_path):
    data_list = []
    campos = [
        "Afinidade:", "Natureza:", "Média Final de Curso:", 
        "Duração Prevista 1º Ciclo:", "Núm Matrículas 1º Ciclo:", 
        "ECTS 2º Ciclo:", "Valoração Atividades Extracurriculares:"
    ]

    if not os.path.exists(file_path):
        print(f"AVISO: Arquivo {file_path} não encontrado.")
        return data_list

    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = clean_value(line)
            if not line or "Afinidade:" not in line:
                continue
            
            row = {}
            parte_inicial = line.split("Afinidade:")[0].strip()
            nome_match = re.match(r"^(.*?)\s*\(", parte_inicial)
            row["Nome"] = nome_match.group(1).strip() if nome_match else parte_inicial
            
            status_match = re.search(r"\((.*?):\s*([\d,]+)\)", parte_inicial)
            if status_match:
                row["Status"] = status_match.group(1).strip()
                row["Nota Colocação"] = status_match.group(2).strip()
            else:
                row["Status"] = "N/A"; row["Nota Colocação"] = "0"

            temp_line = line
            for i in range(len(campos)):
                atual = campos[i]
                if i + 1 < len(campos):
                    proximo = campos[i+1]
                    padrao = f"{re.escape(atual)}(.*?){re.escape(proximo)}"
                    match = re.search(padrao, temp_line)
                    valor = match.group(1) if match else "0"
                else:
                    valor = temp_line.split(atual)[-1]
                
                col_name = atual.replace(":", "").strip()
                row[col_name] = clean_value(valor)

            data_list.append(row)
    return data_list

# --- Lógica de Escrita ---

if os.path.exists(output_file):
    writer_mode = {'mode': 'a', 'if_sheet_exists': 'replace', 'engine': 'openpyxl'}
else:
    writer_mode = {'mode': 'w', 'engine': 'openpyxl'}

with pd.ExcelWriter(output_file, **writer_mode) as writer:
    for sheet_name, file_path in files_to_process.items():
        results = parse_data_robust(file_path)
        if results:
            df = pd.DataFrame(results)
            
            # --- CONVERSÃO E LIMPEZA DE NÚMEROS ---
            # Lista de colunas que devem ser numéricas (com ponto em vez de vírgula)
            cols_to_fix = ['Nota Colocação', 'Média Final de Curso']
            
            for col in cols_to_fix:
                if col in df.columns:
                    # Substitui vírgula por ponto e converte para número (float)
                    df[col] = df[col].str.replace(',', '.').astype(float)
            
            # --- ORDENAÇÃO ---
            # Agora que é numérico, a ordenação é nativa e precisa
            df = df.sort_values(by='Nota Colocação', ascending=False)
            # -----------------

            df.to_excel(writer, index=False, sheet_name=sheet_name)
            print(f"Concluído: Aba '{sheet_name}' atualizada (Médias com '.' e Ordenadas).")

print(f"\nFicheiro atualizado com sucesso: {output_file}")