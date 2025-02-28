import tkinter as tk
import pandas as pd
import os
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
    
def geraPlanilha():
    # Caminho da pasta que você deseja analisar
    pasta = 'X:/10-USUARIOS/arthur/NFs Pendetes'

    # Listar arquivos na pasta
    arquivos = os.listdir(pasta)

    # Criar listas para armazenar os dados
    nomes_documentos = []
    formatos_documentos = []

    # Preencher as listas com os nomes e formatos dos arquivos
    for arquivo in arquivos:
        if arquivo == ".sync":  # Ignorar o arquivo .sync
            continue
        nome, formato = os.path.splitext(arquivo)
        nomes_documentos.append(nome)
        formatos_documentos.append(formato)

    # Criar um DataFrame com os dados
    df = pd.DataFrame({
        'Nome documento': nomes_documentos,
        'Formato documento': formatos_documentos
    })

    # Salvar o DataFrame em uma planilha Excel
    df.to_excel('planilha_documentos.xlsx', index=False)

def comparaPlanilhas():
    def preencher_dados_diretamente(xlsx_origem, xlsx_destino):
        # 1. Ler a planilha de origem e verificar os nomes das colunas
        df_origem = pd.read_excel(xlsx_origem, dtype=str)  # Garante que os valores sejam strings
        
        colunas_disponiveis = df_origem.columns.tolist()
        print("Colunas encontradas na planilha de origem:", colunas_disponiveis)

        if "E" in colunas_disponiveis:
            coluna_dados = "E"
        else:
            if len(colunas_disponiveis) >= 5:
                coluna_dados = colunas_disponiveis[4]  # Índice 4 é a 5ª coluna
            else:
                raise ValueError("A coluna E não foi encontrada e há menos de 5 colunas disponíveis.")

        # Criar um conjunto de valores únicos, removendo espaços e convertendo para string
        valores_origem = set(df_origem[coluna_dados].dropna().astype(str).str.strip())

        # 2. Abrir ou criar a planilha de destino
        try:
            wb = load_workbook(xlsx_destino)
        except FileNotFoundError:
            wb = Workbook()

        # 3. Pegar a aba principal (ativa) e processar os dados
        ws = wb.active  # Primeira aba da planilha

        # Verificar qual é a última linha com dados na coluna A
        last_row = ws.max_row

        # 4. Preencher os valores diretamente na coluna B
        for row in range(2, last_row + 1):  # Começa da linha 2 (ignorando cabeçalho)
            valor_celula = ws[f"A{row}"].value

            if valor_celula is not None:
                valor_celula = str(valor_celula).strip()  # Garantir que é string e remover espaços
                if valor_celula in valores_origem:
                    ws[f"B{row}"] = "Encontrado"
                else:
                    ws[f"B{row}"] = "Não Encontrado"

        # 5. Salvar a planilha
        wb.save(xlsx_destino)
        print(f"Planilha '{xlsx_destino}' foi atualizada com os valores processados.")

    # Exemplo de uso
    preencher_dados_diretamente("lancamentos-fevereiro.xlsx", "planilha_documentos.xlsx")

def deletar_documentos_repetidos():
    import os
    import pandas as pd

    # Definição do caminho da planilha e da pasta de arquivos
    planilha_nome = "planilha_documentos.xlsx"  # Nome do arquivo .XLSX
    pasta_arquivos = r"C:\Users\Cliente\Desktop\NFs Pendetes"

    # Carregar a planilha
    try:
        df = pd.read_excel(planilha_nome, dtype=str)
        print("Planilha carregada")
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        exit()

    # Verificar se as colunas esperadas existem
    if "Nome documento" not in df.columns or "Formato documento" not in df.columns:
        print("Erro: A planilha não contém as colunas esperadas.")
        exit()

    # Filtrar documentos marcados como "Encontrado", ignorando arquivos que começam com "."
    arquivos_para_excluir = df[df["Formato documento"].str.strip().str.lower() == "encontrado"]["Nome documento"].str.strip()

    # Iterar sobre os arquivos da pasta
    for arquivo in os.listdir(pasta_arquivos):
        nome_base, extensao = os.path.splitext(arquivo)  # Separar nome e extensão
        print("Nome base: ",nome_base)
        if nome_base in arquivos_para_excluir.values:
            caminho_arquivo = os.path.join(pasta_arquivos, arquivo)
            try:
                os.remove(caminho_arquivo)
                print(f"Arquivo deletado: {arquivo}")
            except Exception as e:
                print(f"Erro ao deletar {arquivo}: {e}")

    print("Processo concluído.")

root = tk.Tk()
root.title("Comparar planilhas")

label = tk.Label(root, text="Clique para gerar planilha")
label.pack()

button = tk.Button(root, text="Gerar planilha", command=geraPlanilha)
button.pack()

#---

label = tk.Label(root, text="Clique para comparar planilhas")
label.pack()

button = tk.Button(root, text="Comparar planilhas", command=comparaPlanilhas)
button.pack()

#---

label = tk.Label(root, text="Clique para deletar notas já lançadas")
label.pack()

button = tk.Button(root, text="Deletar notas lançadas", command=deletar_documentos_repetidos)

button.pack()

root.mainloop()
