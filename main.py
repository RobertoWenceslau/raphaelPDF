import pdfplumber
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

# Carregar a tabela de CIDs
cid_mapping = pd.read_csv("cid_mapping.csv", sep=";", encoding="ISO-8859-1")
cid_dict = dict(zip(cid_mapping["SUBCAT"], cid_mapping["DESCRICAO"]))


def corrigir_cid(cid):
    """
    Corrige possíveis erros no formato do CID.
    Exemplo: "1200" → "I200".
    Retorna o CID corrigido e uma mensagem de correção (se houver).
    """
    cid_original = cid
    if len(cid) == 4 and cid[0].isdigit():
        # Se o primeiro caractere for um número, substitua por "I"
        cid = "I" + cid[1:]
        mensagem_correcao = f"CID: {cid_original} → {cid}"
    else:
        mensagem_correcao = ""
    return cid, mensagem_correcao


def extrair_dados(pdf_path):
    dados = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if not texto:
                    continue  # Ignora páginas sem texto

                # Extrair dados do beneficiário
                beneficiario = ""
                if "Beneficiário(a)" in texto:
                    beneficiario = texto.split("Beneficiário(a)")[1].split("\n")[0].strip()

                # Extrair nome do prestador
                prestador = "PRESTADOR NÃO INFORMADO"
                if "Executora" in texto:
                    prestador_pdf = texto.split("Executora")[1].split("\n")[0].strip()
                    if prestador_pdf == "UNIMED FLORIANOPOLIS":
                        prestador = "UNIMED FLORIANOPOLIS"
                    else:
                        prestador = f"UNIMED FLORIANOPOLIS ({prestador_pdf})"

                # Extrair CID Principal
                cid = ""
                motivo = "MOTIVO NÃO INFORMADO"
                if "CID Principal" in texto:
                    cid = texto.split("CID Principal")[1].split("\n")[0].strip()
                    if cid:  # Só corrigir e mapear se o CID não estiver vazio
                        cid, mensagem_correcao = corrigir_cid(cid)
                        motivo = cid_dict.get(cid, f"CID NÃO ENCONTRADO: {cid}")
                        if mensagem_correcao:
                            motivo = f"{mensagem_correcao} - {motivo}"

                # Extrair datas de internação e alta
                data_admissao = "DATA NÃO INFORMADA"
                if "Atendimento" in texto:
                    try:
                        data_admissao = texto.split("Atendimento")[1].split("\n")[0].strip().split()[0]
                        datetime.strptime(data_admissao, "%d/%m/%y")  # Valida o formato da data
                    except (IndexError, ValueError):
                        data_admissao = "DATA NÃO INFORMADA"

                data_alta = "DATA NÃO INFORMADA"
                if "Alta" in texto:
                    try:
                        data_alta = texto.split("Alta")[1].split("\n")[0].strip().split()[0]
                        datetime.strptime(data_alta, "%d/%m/%y")  # Valida o formato da data
                    except (IndexError, ValueError):
                        data_alta = "DATA NÃO INFORMADA"

                # Extrair todas as diárias e datas de acomodação
                diarias = []
                for line in texto.split("\n"):
                    if "DIÁRIA DE" in line:
                        partes = line.split()
                        if len(partes) >= 4:  # Garantir que a linha tenha dados suficientes
                            tipo_diaria = " ".join(partes[2:-2])  # Extrai o tipo de diária (ex: "UTI ADULTO GERAL")
                            data_diaria = partes[-2]  # Extrai a data da diária
                            diarias.append({"tipo": tipo_diaria, "data": data_diaria})

                # Determinar o CARÁTER com base nas datas das diárias
                carater = "ELETIVO"  # Valor padrão
                if diarias:
                    # Converter as datas para objetos datetime para comparação
                    diarias_com_datas = []
                    for diaria in diarias:
                        try:
                            data_diaria = datetime.strptime(diaria["data"], "%d/%m/%y")
                            diarias_com_datas.append({"tipo": diaria["tipo"], "data": data_diaria})
                        except ValueError:
                            continue  # Ignora diárias com datas inválidas

                    # Verificar se há alguma UTI com data anterior ao apartamento
                    datas_uti = [d["data"] for d in diarias_com_datas if "UTI" in d["tipo"]]
                    datas_apartamento = [d["data"] for d in diarias_com_datas if "APARTAMENTO" in d["tipo"]]

                    if datas_uti and datas_apartamento:
                        # Se a menor data de UTI for anterior à menor data de apartamento, é URGÊNCIA
                        if min(datas_uti) < min(datas_apartamento):
                            carater = "URGENCIA"
                    elif datas_uti:
                        # Se houver apenas UTI, é URGÊNCIA
                        carater = "URGENCIA"
                    # Caso contrário, mantém "ELETIVO"

                # Inferir TIPO DE INTERNAÇÃO com base nos procedimentos
                tipo_internacao = "CLINICA"
                if "CATETERISMO" in texto or "CIRURGIA" in texto or "RESSECÇÃO" in texto or "TAXA DE SALA CIRÚRGICA" in texto:
                    tipo_internacao = "CIRURGICA"

                # Adicionar dados à lista
                dados.append({
                    "UF": "SC",
                    "TIPO DE INTERNAÇÃO": tipo_internacao,
                    "CARÁTER": carater,
                    "Beneficiário Atendido": beneficiario,
                    "ACOMODAÇÃO": "APARTAMENTO",  # Acomodação fixa
                    "PRESTADOR": prestador,
                    "CNPJ": "77.658.611/0001-08",
                    "DATA DA ADMISSÃO": data_admissao,
                    "DATA DA ALTA": data_alta,
                    "MOTIVO DA INTERNAÇÃO": motivo
                })

        # Salvar dados em um DataFrame
        return pd.DataFrame(dados)

    except Exception as e:
        print(f"Erro ao processar o arquivo PDF: {e}")
        return pd.DataFrame()  # Retorna um DataFrame vazio em caso de erro

def selecionar_arquivo():
    root =tk.Tk()
    root.withdraw() # Ocultar Janela principal
    caminho_arquivo = filedialog.askopenfilename(title="Selecione o arquivo PDF", filetypes=(("Arquivos PDF", ".pdf"), ("Todos os Arquivos", ".*")))
    return caminho_arquivo


# Exemplo de uso
pdf_path = selecionar_arquivo()
tabela = extrair_dados(pdf_path)
if not tabela.empty:
    tabela.to_excel("censo_internados.xlsx", index=False)
else:
    print("Nenhum dado foi extraído do PDF.")