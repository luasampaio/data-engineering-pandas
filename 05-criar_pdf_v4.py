import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def copiar_pdfs_por_agrupamento(diretorio_base, caminho_excel, ano):
    try:
        if not os.path.exists(diretorio_base):
            messagebox.showerror("Erro", "O diretório base não existe.")
            return

        if not os.path.exists(caminho_excel):
            messagebox.showerror("Erro", "O arquivo Excel não existe.")
            return

        # Lê o arquivo Excel
        df = pd.read_excel(caminho_excel)

        # Filtrar apenas entradas válidas
        df = df.dropna(subset=['NOME_DIRETOR'])
        df['NOME_DIRETOR'] = df['NOME_DIRETOR'].astype(str).str.strip()
        df['NOME_AGRUPAMENTO'] = df['NOME_AGRUPAMENTO'].fillna('').astype(str).str.strip()

        pastas_criadas = set()

        for _, linha in df.iterrows():
            nome_diretor = linha['NOME_DIRETOR']
            nome_agrupamento = linha['NOME_AGRUPAMENTO']

            pdf_diretor = f"{nome_diretor}.pdf"
            caminho_pdf = os.path.join(diretorio_base, pdf_diretor)
            if not os.path.exists(caminho_pdf):
                print(f"Arquivo {pdf_diretor} não encontrado no diretório base. Pulando...")
                continue

            pasta_diretor_com_ano = os.path.join(diretorio_base, nome_diretor, str(ano))

            if pasta_diretor_com_ano not in pastas_criadas:
                os.makedirs(pasta_diretor_com_ano, exist_ok=True)
                pastas_criadas.add(pasta_diretor_com_ano)
                print(f"Pasta do diretor {pasta_diretor_com_ano} criada.")

            shutil.move(caminho_pdf, pasta_diretor_com_ano)
            print(f"Arquivo {pdf_diretor} movido para {pasta_diretor_com_ano}.")

            if nome_agrupamento:
                pasta_agrupamento_com_ano = os.path.join(diretorio_base, nome_agrupamento, str(ano))
                if pasta_agrupamento_com_ano not in pastas_criadas:
                    os.makedirs(pasta_agrupamento_com_ano, exist_ok=True)
                    pastas_criadas.add(pasta_agrupamento_com_ano)
                    print(f"Pasta do agrupamento {pasta_agrupamento_com_ano} criada.")

                shutil.copytree(pasta_diretor_com_ano, pasta_agrupamento_com_ano, dirs_exist_ok=True)
                print(f"Arquivos copiados de {pasta_diretor_com_ano} para {pasta_agrupamento_com_ano}.")
            else:
                print(f"Sem agrupamento para o diretor {nome_diretor}. Apenas a pasta do diretor foi criada.")

        messagebox.showinfo("Sucesso", "Processo concluído com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def selecionar_diretorio():
    diretorio = filedialog.askdirectory(title="Selecione a pasta de PDFs")
    entry_diretorio_base.delete(0, tk.END)
    entry_diretorio_base.insert(0, diretorio)

def selecionar_excel():
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
    )
    entry_excel.delete(0, tk.END)
    entry_excel.insert(0, arquivo)

def executar_script():
    diretorio_base = entry_diretorio_base.get().strip()
    caminho_excel = entry_excel.get().strip()
    ano = entry_ano.get().strip()

    if not diretorio_base or not caminho_excel or not ano:
        messagebox.showwarning("Aviso", "Por favor, preencha todos os campos.")
        return

    if not ano.isdigit():
        messagebox.showwarning("Aviso", "O ano deve ser um número válido.")
        return

    copiar_pdfs_por_agrupamento(diretorio_base, caminho_excel, ano)

# Interface Gráfica
root = tk.Tk()
root.title("Organizador de PDFs - Itaú")

# Configuração de cores
cor_primaria = "#EC7000"
cor_secundaria = "#FFFFFF"
root.configure(bg=cor_secundaria)

# Configuração da janela
root.geometry("550x200")

# Labels e Entradas
tk.Label(root, text="Pasta de PDFs:", bg=cor_secundaria, fg=cor_primaria, font=("Arial", 10, "bold")).grid(row=1, column=0, padx=10, pady=10, sticky="e")
entry_diretorio_base = tk.Entry(root, width=40)
entry_diretorio_base.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar", command=selecionar_diretorio, bg=cor_primaria, fg=cor_secundaria).grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Arquivo Excel:", bg=cor_secundaria, fg=cor_primaria, font=("Arial", 10, "bold")).grid(row=2, column=0, padx=10, pady=10, sticky="e")
entry_excel = tk.Entry(root, width=40)
entry_excel.grid(row=2, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar", command=selecionar_excel, bg=cor_primaria, fg=cor_secundaria).grid(row=2, column=2, padx=10, pady=10)

tk.Label(root, text="Ano do Relatório:", bg=cor_secundaria, fg=cor_primaria, font=("Arial", 10, "bold")).grid(row=3, column=0, padx=10, pady=10, sticky="e")
entry_ano = tk.Entry(root, width=40)
entry_ano.grid(row=3, column=1, padx=10, pady=10)

# Botão de Executar
tk.Button(root, text="Executar", command=executar_script, bg=cor_primaria, fg=cor_secundaria, font=("Arial", 10, "bold")).grid(row=4, column=1, pady=20)

# Loop da Interface
root.mainloop()
