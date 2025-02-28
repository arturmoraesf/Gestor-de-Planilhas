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
    def copiar_coluna_e_aplicar_procv(xlsx_origem, xlsx_destino):
        # 1. Ler a planilha de origem e verificar os nomes das colunas
        df_origem = pd.read_excel(xlsx_origem)
        
        # Verificar se a coluna "E" existe pelo nome ou posição
        colunas_disponiveis = df_origem.columns.tolist()
        print("Colunas encontradas na planilha de origem:", colunas_disponiveis)
        
        if "E" in colunas_disponiveis:
            coluna_dados = "E"
        else:
            # Se "E" não for encontrada, pegar a 5ª coluna (índice 4, pois começa do 0)
            if len(colunas_disponiveis) >= 5:
                coluna_dados = colunas_disponiveis[4]
            else:
                raise ValueError("A coluna E não foi encontrada na planilha de origem e há menos de 5 colunas disponíveis.")
        
        df_coluna = df_origem[[coluna_dados]]
        
        # 2. Abrir ou criar a planilha de destino
        try:
            wb = load_workbook(xlsx_destino)
        except FileNotFoundError:
            wb = Workbook()
        
        # 3. Criar a aba 2 (se não existir) e colar a coluna nela
        sheet_name = "Aba2"
        if sheet_name not in wb.sheetnames:
            wb.create_sheet(sheet_name)
        
        ws2 = wb[sheet_name]
        for idx, value in enumerate(df_coluna[coluna_dados], start=1):
            ws2[f"A{idx}"] = value
        
        # 4. Aplicar PROCV na aba principal
        main_sheet = wb.active  # Primeira aba da planilha
        last_row = len(df_coluna) + 1  # Última linha preenchida
        lookup_formula = f"=IF(ISNA(VLOOKUP(A2, Aba2!A:A, 1, FALSE)), \"Não Encontrado\", \"Encontrado\")"
        
        # Escrever a fórmula na coluna B a partir da linha 2
        for row in range(2, last_row):
            main_sheet[f"B{row}"] = lookup_formula.replace("A2", f"A{row}")
        
        # 5. Salvar a planilha
        wb.save(xlsx_destino)
        
        # 6. Reabrir para imprimir os resultados
        wb = load_workbook(xlsx_destino, data_only=True)
        main_sheet = wb.active
        print("Resultados do PROCV:")
        for row in range(2, last_row):
            print(f"Linha {row}: {main_sheet[f'B{row}'].value}")
        
    # Exemplo de uso
    copiar_coluna_e_aplicar_procv("lancamentos-fevereiro.xlsx", "planilha_documentos.xlsx")

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

    # Filtrar documentos marcados como "Encontrado"
    arquivos_para_excluir = df[df["Formato documento"].str.strip().str.lower() == "encontrado"]["Nome documento"].str.strip()
    print(arquivos_para_excluir)

    # Iterar sobre os arquivos da pasta
    for arquivo in os.listdir(pasta_arquivos):
        nome_base, extensao = os.path.splitext(arquivo)  # Separar nome e extensão
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
