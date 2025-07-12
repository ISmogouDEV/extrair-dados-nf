import PyPDF2
import re
import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
import time

def log(mensagem):
    """Função para exibir mensagens na tela e simular um loading"""
    print(mensagem)
    time.sleep(0.1)  # Simula um pequeno delay para facilitar a leitura do progresso

def extract_text_from_pdf(pdf_path):
    log(f"Lendo PDF: {os.path.basename(pdf_path)}")
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
    return text

def extract_info(text, pdf_path):
    log(f"Extraindo informações de: {os.path.basename(pdf_path)}")

    info = {
        "Arquivo": os.path.basename(pdf_path),
        "NF": "",
        "CNPJ Prestador": "",
        "CNPJ Tomador": "",
        "Município": "",
        "UF": "",
        "Código de Verificação": "",
        "PIS": "",
        "COFINS": "",
        "IR(R$)": "",
        "INSS(R$)": "",
        "CSLL(R$)": "",
        "(-) ISS Retido": "",
        "Valor dos Serviços R$": "",
        "Data da Emissão": "",
        "Hora da Emissão": ""
    }
    
    patterns = {
        "Município": r"Local da Prestação[\s:]*([\w\s]+)-\s*([A-Z]{2})",
        "NF": r"Número da\s*NFS-e[\s:]*([\d]+)",
        "CNPJ Prestador": r"CNPJ/CPF[\s:]*([\d./-]+)",
        "CNPJ Tomador": r"CNPJ/CPF[\s:]*([\d./-]+)",
        "Código de Verificação": r"Código de Verificação[\s:]*([\d]+)",
        "PIS": r"PIS[\s:]*([\d.,]+)",
        "COFINS": r"COFINS[\s:]*([\d.,]+)",
        "IR(R$)": r"IR\(R\$\)[\s:]*([\d.,]+)",
        "INSS(R$)": r"INSS\(R\$\)[\s:]*([\d.,]+)",
        "CSLL(R$)": r"CSLL\(R\$\)[\s:]*([\d.,]+)",
        "(-) ISS Retido": r"\(-\) ISS Retido[\s:]*([\d.,]+)",
        "Valor dos Serviços R$": r"Valor dos Serviços R\$[\s:]*([\d.,]+)",
        "Data e Hora da Emissão": r"Data e Hora da Emissão\s*([\d/]+)\s*([\d:]+)"
    }
    
    for key, pattern in patterns.items():
        matches = re.findall(pattern, text, re.DOTALL)
        if matches:
            if key == "CNPJ Prestador":
                for match in matches:
                    if match.startswith("15.040"):
                        info[key] = match.strip()
                        break
            elif key == "CNPJ Tomador":
                for match in matches:
                    if match.startswith(("06.626", "04.899")):
                        info[key] = match.strip()
                        break
            elif key == "Município":
                info["Município"] = matches[0][0].strip()
                info["UF"] = matches[0][1].strip()
            elif key == "Data e Hora da Emissão":
                info["Data da Emissão"] = matches[0][0].strip()
                info["Hora da Emissão"] = matches[0][1].strip()
            else:
                info[key] = matches[0].strip()
        else:
            info[key] = "Não encontrado"
    
    return info

def calcular_cc(cnpj_tomador):
    if cnpj_tomador.startswith("06.626"):
        parte_cnpj = int(cnpj_tomador.split("/")[1][:4])
        cc = 10000 + parte_cnpj
    elif cnpj_tomador.startswith("04.899"):
        parte_cnpj = int(cnpj_tomador.split("/")[1][:4])
        cc = (10000 + parte_cnpj) + 200000000
    else:
        cc = "Não definido"

    return cc

def select_folder():
    root = tk.Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
    return folder_selected

if __name__ == "__main__":
    log("Iniciando extração de dados...")
    
    folder_path = select_folder()
    if folder_path:
        log(f"Pasta selecionada: {folder_path}")
        extracted_data = []
        
        arquivos = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
        total_arquivos = len(arquivos)

        if total_arquivos == 0:
            log("Nenhum PDF encontrado na pasta.")
        else:
            log(f"Total de PDFs encontrados: {total_arquivos}")

        for i, filename in enumerate(arquivos, 1):
            pdf_path = os.path.join(folder_path, filename)
            text = extract_text_from_pdf(pdf_path)
            extracted_info = extract_info(text, pdf_path)
            extracted_data.append(extracted_info)
            log(f"({i}/{total_arquivos}) Processado: {filename}")

        if extracted_data:
            log("Convertendo para Excel...")

            df = pd.DataFrame(extracted_data)
            df["Centro de Custo"] = df["CNPJ Tomador"].apply(calcular_cc)

            df = df[["Arquivo", "NF", "Centro de Custo", "Município", "UF", "CNPJ Prestador", "CNPJ Tomador", 
                     "Código de Verificação", "PIS", "COFINS", "IR(R$)", "INSS(R$)", "CSLL(R$)", 
                     "(-) ISS Retido", "Valor dos Serviços R$", "Data da Emissão", "Hora da Emissão"]]

            df.drop(columns=["Data e Hora da Emissão"], errors="ignore", inplace=True)
            excel_path = os.path.join(folder_path, "dados_extraidos_com_cc.xlsx")
            df.to_excel(excel_path, index=False)

            log(f"✅ Arquivo Excel salvo em: {excel_path}")
        else:
            log("❌ Nenhum dado extraído.")

    log("Processo concluído!")
