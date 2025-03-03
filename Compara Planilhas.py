import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from openpyxl import load_workbook, Workbook

# Função para gerar a planilha
def geraPlanilha():
    pasta = filedialog.askdirectory(title="Selecione a pasta com os documentos")
    if not pasta:
        messagebox.showwarning("Aviso", "Nenhuma pasta selecionada.")
        return
    
    try:
        arquivos = [f for f in os.listdir(pasta) if not f.startswith('.')]
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao acessar a pasta: {e}")
        return
    
    nomes_documentos = []
    formatos_documentos = []
    for arquivo in arquivos:
        nome, formato = os.path.splitext(arquivo)
        nomes_documentos.append(nome)
        formatos_documentos.append(formato)
    
    df = pd.DataFrame({
        'Nome documento': nomes_documentos,
        'Formato documento': formatos_documentos
    })
    
    df.to_excel('planilha_documentos.xlsx', index=False)
    messagebox.showinfo("Sucesso", "Planilha gerada com sucesso!")

# Função para comparar planilhas
def comparaPlanilhas():
    xlsx_origem = filedialog.askopenfilename(title="Selecione a planilha para comparar", filetypes=[("Arquivos Excel", "*.xlsx")])
    if not xlsx_origem:
        messagebox.showwarning("Aviso", "Nenhuma planilha selecionada.")
        return
    
    xlsx_destino = "planilha_documentos.xlsx"
    
    try:
        df_origem = pd.read_excel(xlsx_origem, dtype=str)
        valores_origem = set(df_origem.iloc[:, 4].dropna().astype(str).str.strip())
        wb = load_workbook(xlsx_destino)
        ws = wb.active
        
        for row in range(2, ws.max_row + 1):
            valor_celula = str(ws[f"A{row}"].value).strip()
            ws[f"B{row}"] = "Encontrado" if valor_celula in valores_origem else "Não Encontrado"
        
        wb.save(xlsx_destino)
        messagebox.showinfo("Sucesso", f"Planilha '{xlsx_destino}' foi atualizada.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar as planilhas: {e}")

# Função para deletar documentos repetidos
def deletar_documentos_repetidos():
    planilha_nome = "planilha_documentos.xlsx"
    pasta_arquivos = filedialog.askdirectory(title="Selecione a pasta com os arquivos")
    if not pasta_arquivos:
        messagebox.showwarning("Aviso", "Nenhuma pasta selecionada.")
        return
    
    try:
        df = pd.read_excel(planilha_nome, dtype=str)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")
        return
    
    if "Nome documento" not in df.columns or "Formato documento" not in df.columns:
        messagebox.showerror("Erro", "A planilha não contém as colunas esperadas.")
        return
    
    arquivos_para_excluir = df[df["Formato documento"].str.strip().str.lower() == "encontrado"]["Nome documento"].str.strip()
    
    for arquivo in os.listdir(pasta_arquivos):
        nome_base, _ = os.path.splitext(arquivo)
        if nome_base in arquivos_para_excluir.values:
            caminho_arquivo = os.path.join(pasta_arquivos, arquivo)
            try:
                os.remove(caminho_arquivo)
                print(f"Arquivo deletado: {arquivo}")
            except Exception as e:
                print(f"Erro ao deletar {arquivo}: {e}")
    
    messagebox.showinfo("Sucesso", "Processo concluído. Arquivos duplicados removidos.")

# Criar a interface gráfica
root = tk.Tk()
root.title("Gestor de Planilhas")
root.geometry("400x350")
root.resizable(False, False)

frame = ttk.Frame(root, padding=20)
frame.pack(fill="both", expand=True)

label = ttk.Label(frame, text="Escolha uma ação:", font=("Arial", 12, "bold"))
label.pack(pady=10)

btn_gerar = ttk.Button(frame, text="Gerar Planilha", command=geraPlanilha)
btn_gerar.pack(pady=5, fill="x")

btn_comparar = ttk.Button(frame, text="Comparar Planilhas", command=comparaPlanilhas)
btn_comparar.pack(pady=5, fill="x")

btn_deletar = ttk.Button(frame, text="Deletar Notas Lançadas", command=deletar_documentos_repetidos)
btn_deletar.pack(pady=5, fill="x")

btn_sair = ttk.Button(frame, text="Sair", command=root.quit)
btn_sair.pack(pady=20, fill="x")

root.mainloop()
