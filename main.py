import tkinter as tk
from tkinter import filedialog
import os
import ajuste_largura_coluna
import busca_dados

print("\n\033[34m - Extração de informações de VÁRIOS arquivos PDFs em uma pasta - \033[m\n")
print("\033[31mATENÇÃO - VERIFICAR SE O ARQUIVO cid_mapping.csv ESTÁ LOCALIZADO NA MESMA PASTA DO EXECUTÁVEL\033[m")
print()
print("\n\033[33mSELECIONE E PASTA COM OS ARQUIVOS PDFs: \033[m\n")



def carregar_pdfs_da_pasta(pasta):
    # Lista para armazenar os caminhos dos arquivos PDF
    dados_pdfs = []

    # Percorre os arquivos na pasta
    for arquivo in os.listdir(pasta):
        # Obtem o caminho completo do arquivo
        caminho_arquivo = os.path.join(pasta, arquivo)

        # Verifica se é um arquivo (e não uma pasta) e se a extensão é .pdf
        if os.path.isfile(caminho_arquivo) and arquivo.lower().endswith(".pdf"):
            dados_pdfs.append(caminho_arquivo)

    return dados_pdfs


# Função para selecionar a pasta
def selecionar_pasta():
    pasta_selecionada = filedialog.askdirectory()
    if pasta_selecionada:
        print(f"\n\033[32mPasta selecionada: \033[m{pasta_selecionada}")
        dados = carregar_pdfs_da_pasta(pasta_selecionada)

        # Exibir os dados extraídos
        print("\033[33m\nArquivos identificados na Pasta informada: \033[m\n")
        for item in dados:
            print(f"Arquivo: {item}")
            # print(f"Texto: {item['Texto'][:100]}...")  # Exibe apenas os primeiros 100 caracteres
            print("-" * 50)

    return dados, pasta_selecionada

# Criar a janela principal do tkinter
root = tk.Tk()
root.withdraw()  # Esconder a janela principal

# Abrir a caixa de diálogo para selecionar a pasta
lista_de_pdfs, caminho_pasta = selecionar_pasta()

# Loop para iterar pelos arquivos PDFs na pasta selecionada e extrair os dados
primeira_iteracao = True
for pdf in lista_de_pdfs:
    busca_dados.dados_pdf(pdf, primeira_iteracao, caminho_pasta)
    primeira_iteracao = False  # Após a primeira iteração, define como False

arquivo_excel = f'{caminho_pasta}/censo_internados.xlsx'

# Ajustar larguda das colunas do arquivo Excel Gerado
ajuste_largura_coluna.formatar_excel(arquivo_excel)

# Abre o arquivo no aplicativo padrão (Excel) configurado no sistema
os.startfile(arquivo_excel)
