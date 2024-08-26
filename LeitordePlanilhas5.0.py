import os
import json
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.ttk import Progressbar
import pandas as pd

# Configuração do logging
logging.basicConfig(
    filename=os.path.join(os.environ['USERPROFILE'], 'Desktop', 'logfile.log'),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def carregar_configuracoes():
    caminho_config = os.path.join(os.environ['USERPROFILE'], 'Desktop', 'configuracoes.json')
    if os.path.exists(caminho_config):
        with open(caminho_config, 'r') as file:
            configuracoes = json.load(file)
    else:
        configuracoes = {
            "diretorio_padrao": os.path.join(os.environ['USERPROFILE'], 'Desktop'),
            "formato_data": "%Y-%m-%d %H:%M:%S",
            "banco_preferido": "mysql"
        }
        salvar_configuracoes(configuracoes)
    return configuracoes

def salvar_configuracoes(configuracoes):
    caminho_config = os.path.join(os.environ['USERPROFILE'], 'Desktop', 'configuracoes.json')
    with open(caminho_config, 'w') as file:
        json.dump(configuracoes, file, indent=4)

def carregar_mapeamento():
    caminho_mapeamento = os.path.join(os.environ['USERPROFILE'], 'Desktop', 'mapeamento.json')
    if os.path.exists(caminho_mapeamento):
        with open(caminho_mapeamento, 'r') as file:
            return json.load(file)
    else:
        return {}

def salvar_mapeamento(mapeamento_colunas):
    caminho_mapeamento = os.path.join(os.environ['USERPROFILE'], 'Desktop', 'mapeamento.json')
    with open(caminho_mapeamento, 'w') as file:
        json.dump(mapeamento_colunas, file, indent=4)

def abrir_configuracoes_gui():
    global configuracoes
    config_window = tk.Toplevel(root)
    config_window.title("Configurações")

    tk.Label(config_window, text="Diretório Padrão:").grid(row=0, column=0, padx=10, pady=5)
    diretorio_entry = tk.Entry(config_window, width=50)
    diretorio_entry.grid(row=0, column=1, padx=10, pady=5)
    diretorio_entry.insert(0, configuracoes['diretorio_padrao'])

    tk.Label(config_window, text="Formato de Data:").grid(row=1, column=0, padx=10, pady=5)
    data_entry = tk.Entry(config_window, width=50)
    data_entry.grid(row=1, column=1, padx=10, pady=5)
    data_entry.insert(0, configuracoes['formato_data'])

    tk.Label(config_window, text="Banco Preferido:").grid(row=2, column=0, padx=10, pady=5)
    banco_entry = tk.Entry(config_window, width=50)
    banco_entry.grid(row=2, column=1, padx=10, pady=5)
    banco_entry.insert(0, configuracoes['banco_preferido'])

    def salvar_alteracoes():
        configuracoes['diretorio_padrao'] = diretorio_entry.get()
        configuracoes['formato_data'] = data_entry.get()
        configuracoes['banco_preferido'] = banco_entry.get()
        salvar_configuracoes(configuracoes)
        messagebox.showinfo("Configurações", "Configurações salvas com sucesso!")
        config_window.destroy()

    tk.Button(config_window, text="Salvar", command=salvar_alteracoes).grid(row=3, column=1, pady=10)

def abrir_mapeamento_gui():
    global janela_mapeamento, entries_mapeamento
    janela_mapeamento = tk.Toplevel(root)
    janela_mapeamento.title("Mapeamento de Colunas")

    tk.Label(janela_mapeamento, text="Mapeie as colunas do arquivo para as colunas do banco de dados").pack(padx=10, pady=10)

    if df is not None:
        entries_mapeamento = {}
        for coluna in df.columns:
            tk.Label(janela_mapeamento, text=f"{coluna}:").pack(anchor="w", padx=10)
            entry = tk.Entry(janela_mapeamento, width=50)
            entry.pack(padx=10, pady=5)
            entries_mapeamento[coluna] = entry
        
        tk.Button(janela_mapeamento, text="Salvar Mapeamento", command=salvar_mapeamento).pack(pady=10)
    else:
        tk.messagebox.showerror("Erro", "Nenhum arquivo carregado para mapear colunas.")

def salvar_mapeamento():
    global mapeamento_colunas
    mapeamento_colunas = {coluna: entry.get().strip() for coluna, entry in entries_mapeamento.items()}
    salvar_mapeamento(mapeamento_colunas)
    tk.messagebox.showinfo("Mapeamento", "Mapeamento salvo com sucesso!")
    janela_mapeamento.destroy()

def processar_arquivo_gui():
    global df
    try:
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not caminho_arquivo:
            return

        progress_bar['value'] = 10
        root.update_idletasks()

        df = pd.read_excel(caminho_arquivo)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

        progress_bar['value'] = 50
        root.update_idletasks()

        df = validar_e_limpar_dados(df)
        if df is None:
            messagebox.showerror("Erro", "Erro na validação e limpeza de dados.")
            return
        
        progress_bar['value'] = 80
        root.update_idletasks()

        mostrar_dados(df)

        progress_bar['value'] = 100
        root.update_idletasks()

    except Exception as e:
        log_error(f"Erro ao processar arquivo: {str(e)}")
        messagebox.showerror("Erro", f"Erro ao processar arquivo: {str(e)}")

def mostrar_dados(df):
    for i in treeview.get_children():
        treeview.delete(i)

    treeview["column"] = list(df.columns)
    treeview["show"] = "headings"

    for col in treeview["columns"]:
        treeview.heading(col, text=col)

    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        treeview.insert("", "end", values=row)

def filtrar_dados_gui():
    global df
    if df is None:
        tk.messagebox.showerror("Erro", "Nenhum arquivo foi carregado!")
        return
    
    colunas = df.columns.tolist()
    coluna_selecionada = simpledialog.askstring("Filtro de Dados", f"Escolha a coluna para filtrar:\n{', '.join(colunas)}")
    
    if coluna_selecionada not in colunas:
        tk.messagebox.showerror("Erro", "Coluna inválida!")
        return
    
    valor_filtro = simpledialog.askstring("Filtro de Dados", f"Digite o valor para filtrar na coluna '{coluna_selecionada}':")
    
    if valor_filtro is None or valor_filtro == '':
        tk.messagebox.showerror("Erro", "Valor de filtro não pode ser vazio!")
        return
    
    df_filtrado = df[df[coluna_selecionada].astype(str).str.contains(valor_filtro, case=False, na=False)]
    
    if df_filtrado.empty:
        tk.messagebox.showinfo("Filtro de Dados", "Nenhum dado encontrado com o filtro aplicado.")
    else:
        tk.messagebox.showinfo("Filtro de Dados", f"{len(df_filtrado)} registros encontrados.")
        mostrar_dados(df_filtrado)

def exportar_dados_gui():
    global df
    if df is None:
        tk.messagebox.showerror("Erro", "Nenhum arquivo foi carregado!")
        return
    
    caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if caminho_arquivo:
        try:
            df.to_excel(caminho_arquivo, index=False)
            tk.messagebox.showinfo("Exportação", f"Dados exportados com sucesso para {caminho_arquivo}")
        except Exception as e:
            log_error(f"Erro ao exportar dados: {str(e)}")
            tk.messagebox.showerror("Erro", f"Erro ao exportar dados: {str(e)}")

def criar_inserts_gui():
    global entry_tabela
    global janela_inserts
    janela_inserts = tk.Toplevel(root)
    janela_inserts.title("Criar Inserts SQL")

    tk.Label(janela_inserts, text="Nome da Tabela:").pack(padx=10, pady=5)
    entry_tabela = tk.Entry(janela_inserts, width=50)
    entry_tabela.pack(padx=10, pady=5)

    tk.Button(janela_inserts, text="Gerar Inserts", command=gerar_inserts).pack(pady=10)

def gerar_inserts():
    global df, mapeamento_colunas, entry_tabela
    if df is None:
        tk.messagebox.showwarning("Aviso", "Nenhum dado disponível para criar os inserts.")
        return

    nome_tabela = entry_tabela.get().strip()
    if not nome_tabela:
        tk.messagebox.showwarning("Aviso", "Por favor, insira o nome da tabela.")
        return

    if not mapeamento_colunas:
        tk.messagebox.showwarning("Aviso", "Mapeamento de colunas não encontrado. Configure o mapeamento antes de gerar os inserts.")
        return

    inserts = []
    for index, row in df.iterrows():
        colunas = []
        valores = []
        for coluna, coluna_banco in mapeamento_colunas.items():
            colunas.append(coluna_banco)
            valor = row[coluna]
            if pd.isna(valor):
                valores.append("NULL")
            else:
                if isinstance(valor, str):
                    valor_escapado = valor.replace("'", "''")
                    valores.append(f"'{valor_escapado}'")
                elif isinstance(valor, (int, float)):
                    valores.append(str(valor))
                elif isinstance(valor, pd.Timestamp):
                    valor_formatado = valor.strftime("'%Y-%m-%d %H:%M:%S'")
                    valores.append(valor_formatado)
        
        insert_query = f"INSERT INTO {nome_tabela} ({', '.join(colunas)}) VALUES ({', '.join(valores)});"
        inserts.append(insert_query)

    file_path = filedialog.asksaveasfilename(
        defaultextension=".sql",
        filetypes=[("SQL files", "*.sql")],
        title="Salvar Script SQL"
    )

    if file_path:
        try:
            with open(file_path, "w") as file:
                for insert in inserts:
                    file.write(insert + "\n")
            tk.messagebox.showinfo("Sucesso", f"Script SQL criado com sucesso em {file_path}")
        except Exception as e:
            tk.messagebox.showerror("Erro", f"Erro ao criar o script SQL: {str(e)}")
        finally:
            janela_inserts.destroy()

def log_error(message):
    logging.error(message)
    messagebox.showerror("Erro", message)

def validar_e_limpar_dados(df):
    try:
        df = df.dropna(how='all')
        return df
    except Exception as e:
        log_error(f"Erro ao validar e limpar dados: {str(e)}")
        return None

def editar_colunas_gui():
    global df
    if df is None:
        tk.messagebox.showerror("Erro", "Nenhum arquivo foi carregado!")
        return
    
    editor_window = tk.Toplevel(root)
    editor_window.title("Editar Colunas")

    tk.Label(editor_window, text="Selecione as colunas a remover:").pack(padx=10, pady=10)

    coluna_var = tk.Variable(value=df.columns.tolist())
    lista_colunas = tk.Listbox(editor_window, listvariable=coluna_var, selectmode=tk.MULTIPLE, height=10)
    lista_colunas.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    def remover_colunas():
        colunas_remover = [lista_colunas.get(i) for i in lista_colunas.curselection()]
        if not colunas_remover:
            tk.messagebox.showerror("Erro", "Nenhuma coluna selecionada para remover!")
            return
        
        global df
        df = df.drop(columns=colunas_remover)
        mostrar_dados(df)
        tk.messagebox.showinfo("Edição de Colunas", "Colunas removidas com sucesso!")
        editor_window.destroy()

    tk.Button(editor_window, text="Remover Colunas", command=remover_colunas).pack(pady=10)

def criar_menu():
    menu_bar = tk.Menu(root)
    root.config(menu=menu_bar)

    file_menu = tk.Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Arquivo", menu=file_menu)
    file_menu.add_command(label="Abrir Arquivo", command=processar_arquivo_gui)
    file_menu.add_command(label="Exportar Dados", command=exportar_dados_gui)
    file_menu.add_separator()
    file_menu.add_command(label="Sair", command=root.quit)

    edit_menu = tk.Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Editar", menu=edit_menu)
    edit_menu.add_command(label="Configurações", command=abrir_configuracoes_gui)
    edit_menu.add_command(label="Mapeamento de Colunas", command=abrir_mapeamento_gui)
    edit_menu.add_command(label="Filtrar Dados", command=filtrar_dados_gui)
    edit_menu.add_command(label="Editar Colunas", command=editar_colunas_gui)
    edit_menu.add_command(label="Criar Inserts SQL", command=criar_inserts_gui)

# Configuração da interface gráfica
root = tk.Tk()
root.title("Processador de Dados")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10, fill="both", expand=True)

treeview = ttk.Treeview(frame)
treeview.pack(side="left", fill="both", expand=True)

scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=treeview.xview)
scroll_x.pack(side="bottom", fill="x")
treeview.configure(xscrollcommand=scroll_x.set)

scroll_y = ttk.Scrollbar(frame, orient="vertical", command=treeview.yview)
scroll_y.pack(side="right", fill="y")
treeview.configure(yscrollcommand=scroll_y.set)

progress_bar = Progressbar(root, orient="horizontal", length=100, mode="determinate")
progress_bar.pack(pady=10, fill="x")

df = None
mapeamento_colunas = carregar_mapeamento()
configuracoes = carregar_configuracoes()

criar_menu()

root.mainloop()
