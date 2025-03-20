import re
import pandas as pd

def parse_line(line):
    data = {
        'Nº Beneficiário': '',
        'Beneficiário': '',
        'CPF': '',
        'Plano': '',
        'Tp.': '',
        'Id.': '',
        'Dependência': '',
        'Dt Inclusão': '',
        'Rubrica': '',
        'Valor': '',
        'Valor Total': ''
    }

    # 1. Capturar Nº Beneficiário (formato XXX.XXXXXXX)
    n_beneficiario_match = re.match(r'^(\d+\.\d+)', line)
    if n_beneficiario_match:
        data['Nº Beneficiário'] = n_beneficiario_match.group(1)
        line = line.replace(n_beneficiario_match.group(1), '', 1).strip()

    # 2. Capturar Nome (até encontrar CPF ou Plano)
    nome_split = re.split(r'(?=\d{11}|AMBULATORIAL)', line)
    if nome_split:
        data['Beneficiário'] = nome_split[0].strip()
        line = line.replace(nome_split[0], '', 1).strip()

    # 3. Capturar CPF (11 dígitos)
    cpf_match = re.search(r'(\d{3}\.?\d{3}\.?\d{3}-?\d{2})', line)
    if cpf_match:
        data['CPF'] = cpf_match.group(1)
        line = line.replace(cpf_match.group(1), '', 1).strip()

    # 4. Capturar Plano (AMBULATORIAL I)
    plano_match = re.search(r'(AMBULATORIAL I?)', line)
    if plano_match:
        data['Plano'] = plano_match.group(1)
        line = line.replace(plano_match.group(1), '', 1).strip()

    # 5. Capturar Tp. Id. (T ou D)
    tp_id_match = re.search(r'\b([TD])\b', line)
    if tp_id_match:
        data['Tp.'] = tp_id_match.group(1)
        line = line.replace(tp_id_match.group(1), '', 1).strip()

    # 6. Capturar Id. (número após T/D)
    id_match = re.search(r'(\d+)', line)
    if id_match and data['Tp.']:
        data['Id.'] = id_match.group(1)
        line = line.replace(id_match.group(1), '', 1).strip()

    # 7. Capturar Dependência (texto entre Id. e Data Limite)
    dependencia_split = re.split(r'(\d{2}/\d{2}/\d{4})', line)
    if len(dependencia_split) >= 2:
        # Tudo antes da data é a "Dependência"
        data['Dependência'] = dependencia_split[0].strip()
        line = line.replace(dependencia_split[0], '', 1).strip()

    # 8. Capturar Data Inclusao (dd/mm/aaaa)
    data_inclusao_match = re.search(r'(\d{2}/\d{2}/\d{4})', line)
    if data_inclusao_match:
        data['Dt Inclusão'] = data_inclusao_match.group(1)
        line = line.replace(data_inclusao_match.group(1), '', 1).strip()

    # 9. Capturar Rubrica (texto após "Mensalidade")
    rubrica_match = re.search(r'(Mensalidade.*?)(\d+,\d+)', line)
    if rubrica_match:
        data['Rubrica'] = rubrica_match.group(1).strip()
        line = line.replace(rubrica_match.group(1), '', 1).strip()

    # 10. Capturar Valor e Valor Total
    valores = re.findall(r'(\d+,\d+)', line)
    if valores:
        data['Valor'] = valores[0]
        data['Valor Total'] = valores[1] if len(valores) >= 2 else ''

    return data


def organize_pdf_content(input_excel, output_excel):
    """Processa o arquivo Excel bruto e gera a saída formatada"""
    try:
        print(f"[DEBUG] Lendo arquivo: {input_excel}")
        df = pd.read_excel(input_excel)
        organized_data = []

        for index, row in df.iterrows():
            # Processar todas as páginas
            lines = row['Conteúdo'].split('\n')
            
            # Encontrar início (após "N° Beneficiário") e fim ("ANS - nº")
            start_index = None
            end_index = None
            
            for i, line in enumerate(lines):
                # Detectar início da tabela
                if re.search(r'N[°º]\s*Beneficiário', line, re.IGNORECASE):
                    start_index = i + 1  # Linha após o cabeçalho
                
                # Detectar fim da tabela
                if re.search(r'ANS\s*-\s*n[°º]', line, re.IGNORECASE):
                    end_index = i
                    break  # Parar ao encontrar o fim
            
            # Extrair linhas válidas
            if start_index is not None and end_index is not None:
                data_lines = lines[start_index:end_index]
                
                for line in data_lines:
                    parsed_data = parse_line(line)
                    if parsed_data['Nº Beneficiário']:
                        organized_data.append(parsed_data)

        # Criar DataFrame e salvar
        if organized_data:
            df_organized = pd.DataFrame(organized_data)
            df_organized.to_excel(output_excel, index=False, engine='openpyxl')
            print(f"[DEBUG] Dados salvos em: {output_excel}")
            return True
        else:
            print("[DEBUG] Nenhum dado válido encontrado para processar")
            return False

    except Exception as e:
        print(f"[DEBUG] Erro no processamento: {str(e)}")
        raise Exception(f"Erro no processamento dos dados: {str(e)}")