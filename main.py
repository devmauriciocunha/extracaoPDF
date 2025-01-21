import os
import re
import json
import pandas as pd
from PyPDF2 import PdfReader

def extract_information_by_section(text):
    """
    Extrai informações específicas do texto do PDF considerando seções variáveis.
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

    # Extração de padrões por seção com tolerância a variações
    substance_match = re.search(r"(Nome do produto|Substância)\s*:\s*(.+)", text, re.IGNORECASE)
    un_number_match = re.search(r"(Número ONU|ONU)\s*:\s*(\d+)", text, re.IGNORECASE)
    risk_number_match = re.search(r"(Número de Risco|Risco)\s*:\s*(\d+)", text, re.IGNORECASE)
    class_match = re.search(r"(Classe)\s*:\s*(\d+)", text, re.IGNORECASE)
    subsidiary_risk_match = re.search(r"(Risco Subsidiário|Subsidiário)\s*:\s*(.+)", text, re.IGNORECASE)
    first_aid_match = re.search(r"(4\.\s*PRIMEIROS SOCORROS|PRIMEIROS SOCORROS)(.+?)(\d+\. |\Z)", text, re.DOTALL | re.IGNORECASE)
    fire_measures_match = re.search(r"(5\.\s*MEDIDAS DE COMBATE A INCÊNDIO|COMBATE A INCÊNDIO)(.+?)(\d+\. |\Z)", text, re.DOTALL | re.IGNORECASE)
    accidental_measures_match = re.search(r"(6\.\s*MEDIDAS A TOMAR EM CASO DE FUGAS ACIDENTAIS|FUGAS ACIDENTAIS)(.+?)(\d+\. |\Z)", text, re.DOTALL | re.IGNORECASE)

    # Atribuindo valores encontrados
    info["Substância"] = substance_match.group(2).strip() if substance_match else None
    info["Número ONU"] = un_number_match.group(2).strip() if un_number_match else None
    info["Número de Risco"] = risk_number_match.group(2).strip() if risk_number_match else None
    info["Classe"] = class_match.group(2).strip() if class_match else None
    info["Risco Subsidiário"] = subsidiary_risk_match.group(2).strip() if subsidiary_risk_match else None
    info["Primeiros Socorros"] = first_aid_match.group(2).strip() if first_aid_match else None
    info["Medidas de Combate ao Incêndio"] = fire_measures_match.group(2).strip() if fire_measures_match else None
    info["Medidas a Tomar em Caso de Fugas Acidentais"] = accidental_measures_match.group(2).strip() if accidental_measures_match else None

    return info

def process_all_pdfs_in_directory(input_dir, output_dir, consolidated_excel_name="consolidated_data.xlsx"):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Inicializando o DataFrame vazio
    consolidated_df = pd.DataFrame()
    results = {}

    for file_name in os.listdir(input_dir):
        if file_name.endswith('.pdf'):
            pdf_path = os.path.join(input_dir, file_name)

            # Processa cada PDF e extrai informações
            try:
                with open(pdf_path, 'rb') as pdf_file:
                    reader = PdfReader(pdf_file)
                    pages = []
                    full_text = ""
                    for page_number, page in enumerate(reader.pages):
                        text = page.extract_text()
                        full_text += text + "\n"
                        pages.append({"page_number": page_number + 1, "text": text})

                    # Extraindo informações
                    extracted_info = extract_information_by_section(full_text)

                    # Adicionando ao DataFrame consolidado
                    extracted_info["Arquivo"] = file_name  # Inclui o nome do arquivo para referência
                    consolidated_df = pd.concat(
                        [consolidated_df, pd.DataFrame([extracted_info])],
                        ignore_index=True
                    )

                    # Criando JSON completo
                    full_json = {
                        "pages": pages,
                        "extracted_information": extracted_info
                    }

                    # Salvando JSON individual
                    json_path = os.path.join(output_dir, f"{os.path.splitext(file_name)[0]}.json")
                    with open(json_path, 'w', encoding='utf-8') as json_file:
                        json.dump(full_json, json_file, indent=4, ensure_ascii=False)

                    results[file_name] = {"status": "Success", "json_path": json_path}
            except Exception as e:
                results[file_name] = {"status": "Error", "error": str(e)}

    # Salvando o DataFrame consolidado em um único arquivo Excel
    consolidated_excel_path = os.path.join(output_dir, consolidated_excel_name)
    consolidated_df.to_excel(consolidated_excel_path, index=False, engine='openpyxl')

    # Retornando o caminho do arquivo consolidado e os resultados de cada PDF
    return {"consolidated_excel_path": consolidated_excel_path, "results": results}

# Caminhos fixos fornecidos
input_dir = r"C:/Users/mauri/OneDrive/Área de Trabalho/extraindoDados/"
output_dir = input_dir

# Processando todos os PDFs e salvando tudo na mesma tabela
process_results = process_all_pdfs_in_directory(input_dir, output_dir)

# Exibindo os resultados
print(json.dumps(process_results, indent=4, ensure_ascii=False))
