import json
import PyPDF2
import os
import re
import pandas as pd

def extract_information(text):
    """
    Extrai informações específicas do texto do PDF.
    """
    info = {
        "Substância": None,
        "Número ONU": None,
        "Número de Risco": None,
        "Classe": None,
        "Risco Subsidiário": None,
        "Primeiros Socorros": None,
        "Medidas de Combate ao Incêndio": None,
        "Medidas a Tomar em Caso de Fugas Acidentais": None
    }

    # Extração de padrões específicos
    substance_match = re.search(r"Nome do produto\s*:\s*(.+)", text)
    un_number_match = re.search(r"Número ONU\s*:\s*(\d+)", text)
    risk_number_match = re.search(r"Número de Risco\s*:\s*(\d+)", text)
    class_match = re.search(r"Classe\s*:\s*(\d+)", text)
    subsidiary_risk_match = re.search(r"Risco Subsidiário\s*:\s*(.+)", text)
    first_aid_match = re.search(r"4\. PRIMEIROS SOCORROS(.+?)(\d+\. |\Z)", text, re.DOTALL)
    fire_measures_match = re.search(r"5\. MEDIDAS DE COMBATE A INCÊNDIO(.+?)(\d+\. |\Z)", text, re.DOTALL)
    accidental_measures_match = re.search(r"6\. MEDIDAS A TOMAR EM CASO DE FUGAS ACIDENTAIS\s*(.+?)(\d+\. |\Z)", text, re.DOTALL)

    # Atribuindo valores encontrados
    info["Substância"] = substance_match.group(1).strip() if substance_match else None
    info["Número ONU"] = un_number_match.group(1).strip() if un_number_match else None
    info["Número de Risco"] = risk_number_match.group(1).strip() if risk_number_match else None
    info["Classe"] = class_match.group(1).strip() if class_match else None
    info["Risco Subsidiário"] = subsidiary_risk_match.group(1).strip() if subsidiary_risk_match else None
    info["Primeiros Socorros"] = first_aid_match.group(1).strip() if first_aid_match else None
    info["Medidas de Combate ao Incêndio"] = fire_measures_match.group(1).strip() if fire_measures_match else None
    info["Medidas a Tomar em Caso de Fugas Acidentais"] = accidental_measures_match.group(1).strip() if accidental_measures_match else None

    return info

def pdf_to_json_with_extraction(pdf_path, json_path):
    try:
        # Resolvendo os caminhos absolutos
        pdf_path = os.path.abspath(pdf_path)
        json_path = os.path.abspath(json_path)

        # Abrindo o arquivo PDF
        with open(pdf_path, 'rb') as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)

            # Extraindo texto de cada página
            extracted_data = {
                "pages": [],
                "extracted_information": {}
            }

            full_text = ""
            for i, page in enumerate(reader.pages):
                text = page.extract_text()
                extracted_data["pages"].append({"page_number": i + 1, "text": text.strip()})
                full_text += text + "\n"

            # Extraindo informações específicas
            extracted_data["extracted_information"] = extract_information(full_text)

        # Salvando os dados em formato JSON
        with open(json_path, 'w', encoding='utf-8') as json_file:
            json.dump(extracted_data, json_file, indent=4, ensure_ascii=False)

        print(f"Arquivo JSON salvo com sucesso em: {json_path}")

        # Gerando a tabela a partir do JSON extraído
        create_table_from_json(json_path)

    except Exception as e:
        print(f"Erro ao processar o PDF: {e}")

def create_table_from_json(json_path):
    """
    Lê o arquivo JSON e cria uma tabela com as informações extraídas.
    """
    try:
        # Lendo o arquivo JSON
        with open(json_path, 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)

        # Extraindo as informações para a tabela
        extracted_info = data.get("extracted_information", {})

        # Criando o DataFrame
        df = pd.DataFrame([extracted_info])

        # Definindo o caminho do arquivo Excel
        table_path = os.path.splitext(json_path)[0] + ".xlsx"

        # Salvando a tabela em formato Excel
        df.to_excel(table_path, index=False, engine='openpyxl')
        print(f"Tabela salva com sucesso em: {table_path}")

    except Exception as e:
        print(f"Erro ao criar a tabela: {e}")

# Caminhos fixos fornecidos
pdf_path = "C:/Users/mauri/OneDrive/Área de Trabalho/extraindoDados/4-Acetilbenzonitrila.pdf"
json_path = "C:/Users/mauri/OneDrive/Área de Trabalho/extraindoDados/4-Acetilbenzonitrila.json"

# Chamada da função
pdf_to_json_with_extraction(pdf_path, json_path)
