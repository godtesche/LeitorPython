import os
import json
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.ttk import Progressbar
import pandas as pd

# Configuração do logging
logging.basicConfig(filename=os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'logfile.log'),
                    level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Carregar ou criar configurações padrão
def carregar_configuracoes():
    caminho_config = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'configuracoes.json')
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
    caminho_config = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'configuracoes.json')
    with open(caminho_config, 'w') as file:
        json.dump(configuracoes, file, indent=4)

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

def processar_arquivo_gui():
    global df  # Usar a variável global df
    try:
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not caminho_arquivo:
            return

        progress_bar['value'] = 10
        root.update_idletasks()

        df = pd.read_excel(caminho_arquivo)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]  # Remove colunas "Unnamed"

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

    treeview["columns"] = list(df.columns)
    treeview["show"] = "headings"

    for col in treeview["columns"]:
        treeview.heading(col, text=col)

    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        treeview.insert("", "end", values=row)

def criar_menu():
    menu_bar = tk.Menu(root)
    
    # Menu Arquivo
    arquivo_menu = tk.Menu(menu_bar, tearoff=0)
    arquivo_menu.add_command(label="Abrir Arquivo", command=processar_arquivo_gui)
    arquivo_menu.add_command(label="Configurações", command=abrir_configuracoes_gui)
    arquivo_menu.add_separator()
    arquivo_menu.add_command(label="Sair", command=root.quit)
    menu_bar.add_cascade(label="Arquivo", menu=arquivo_menu)
    
    # Menu Processar
    processar_menu = tk.Menu(menu_bar, tearoff=0)
    processar_menu.add_command(label="Filtrar Dados", command=filtrar_dados_gui)
    processar_menu.add_command(label="Exportar Dados", command=exportar_dados_gui)
    processar_menu.add_command(label="Criar Inserts SQL", command=criar_inserts_gui)
    menu_bar.add_cascade(label="Processar", menu=processar_menu)
    
    root.config(menu=menu_bar)

def filtrar_dados_gui():
    global df  # Usar a variável global df
    if df is None:
        tk.messagebox.showerror("Erro", "Nenhum arquivo foi carregado!")
        return

    colunas = df.columns.tolist()

    # Interface para selecionar a coluna
    coluna_selecionada = simpledialog.askstring("Filtro de Dados", f"Escolha a coluna para filtrar:\n{', '.join(colunas)}")

    if coluna_selecionada not in colunas:
        tk.messagebox.showerror("Erro", "Coluna inválida!")
        return

    # Interface para inserir o valor de filtragem
    valor_filtro = simpledialog.askstring("Filtro de Dados", f"Digite o valor para filtrar na coluna '{coluna_selecionada}':")

    if valor_filtro is None or valor_filtro == '':
        tk.messagebox.showerror("Erro", "Valor de filtro não pode ser vazio!")
        return

    # Aplicando o filtro
    df_filtrado = df[df[coluna_selecionada].astype(str).str.contains(valor_filtro, case=False, na=False)]

    # Verificando se o filtro retornou resultados
    if df_filtrado.empty:
        tk.messagebox.showinfo("Filtro de Dados", "Nenhum dado encontrado com o filtro aplicado.")
    else:
        tk.messagebox.showinfo("Filtro de Dados", f"{len(df_filtrado)} registros encontrados.")
        exibir_dados(df_filtrado)  # Mostrar dados filtrados

def exibir_dados(dados):
    # Verifica se a janela principal já foi criada
    if 'root' not in globals():
        tk.messagebox.showerror("Erro", "Interface gráfica não está disponível.")
        return

    # Cria uma nova janela para exibir os dados
    janela_dados = tk.Toplevel(root)
    janela_dados.title("Dados Filtrados")

    # Configura o frame para a Treeview
    frame = tk.Frame(janela_dados)
    frame.pack(fill=tk.BOTH, expand=True)

    # Configura o widget Treeview
    tree = ttk.Treeview(frame, columns=dados.columns, show='headings')
    tree.pack(fill=tk.BOTH, expand=True)

    # Configura as colunas
    for col in dados.columns:
        tree.heading(col, text=col)
        tree.column(col, anchor=tk.W)

    # Insere os dados na Treeview
    for _, row in dados.iterrows():
        tree.insert('', 'end', values=tuple(row))

    # Adiciona uma barra de rolagem vertical
    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Adiciona uma barra de rolagem horizontal
    scrollbar_x = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
    tree.configure(xscroll=scrollbar_x.set)
    scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

    # Botão de fechar
    button_close = tk.Button(janela_dados, text="Fechar", command=janela_dados.destroy)
    button_close.pack(pady=10)

    janela_dados.geometry("800x600")  # Define o tamanho da janela

def exportar_dados_gui():
    global df  # Usar a variável global df
    if df is None:
        tk.messagebox.showwarning("Aviso", "Nenhum dado para exportar.")
        return

    # Cria uma nova janela para exportar os dados
    janela_exportacao = tk.Toplevel(root)
    janela_exportacao.title("Exportar Dados")

    # Função para salvar os dados
    def salvar_dados(formato):
        filetypes = {
            "CSV": [("CSV files", "*.csv")],
            "Excel": [("Excel files", "*.xlsx")],
            "TXT": [("Text files", "*.txt")]
        }
        file_ext = {
            "CSV": ".csv",
            "Excel": ".xlsx",
            "TXT": ".txt"
        }
        file_path = filedialog.asksaveasfilename(
            defaultextension=file_ext[formato],
            filetypes=filetypes[formato],
            title="Salvar Arquivo"
        )

        if file_path:
            try:
                if formato == "CSV":
                    df.to_csv(file_path, index=False)
                elif formato == "Excel":
                    df.to_excel(file_path, index=False)
                elif formato == "TXT":
                    df.to_csv(file_path, sep='\t', index=False)
                
                tk.messagebox.showinfo("Sucesso", f"Dados exportados com sucesso para {file_path}")
            except Exception as e:
                tk.messagebox.showerror("Erro", f"Erro ao exportar os dados: {str(e)}")
            finally:
                janela_exportacao.destroy()

    # Criação dos botões para exportação
    label = tk.Label(janela_exportacao, text="Escolha o formato para exportar os dados:")
    label.pack(pady=10)

    btn_csv = tk.Button(janela_exportacao, text="Exportar como CSV", command=lambda: salvar_dados("CSV"))
    btn_csv.pack(pady=5)

    btn_excel = tk.Button(janela_exportacao, text="Exportar como Excel", command=lambda: salvar_dados("Excel"))
    btn_excel.pack(pady=5)

    btn_txt = tk.Button(janela_exportacao, text="Exportar como TXT", command=lambda: salvar_dados("TXT"))
    btn_txt.pack(pady=5)

    # Botão para cancelar
    btn_cancelar = tk.Button(janela_exportacao, text="Cancelar", command=janela_exportacao.destroy)
    btn_cancelar.pack(pady=10)

    janela_exportacao.geometry("300x200")  # Define o tamanho da janela

def criar_inserts_gui():
    global df  # Usar a variável global df
    if df is None:
        tk.messagebox.showwarning("Aviso", "Nenhum dado disponível para criar os inserts.")
        return

    # Cria uma nova janela para o mapeamento das colunas
    janela_inserts = tk.Toplevel(root)
    janela_inserts.title("Criar Inserts SQL")

    # Label para o nome da tabela
    label_tabela = tk.Label(janela_inserts, text="Nome da tabela no banco de dados:")
    label_tabela.pack(pady=5)

    # Entrada para o nome da tabela
    entry_tabela = tk.Entry(janela_inserts)
    entry_tabela.pack(pady=5)

    # Frame para o mapeamento das colunas
    frame_colunas = tk.Frame(janela_inserts)
    frame_colunas.pack(pady=10)

    # Dicionário para armazenar os widgets de mapeamento
    colunas_mapping = {}

    # Loop para criar widgets de mapeamento das colunas
    for i, coluna in enumerate(df.columns):
        # Label para a coluna do arquivo
        label_coluna_arquivo = tk.Label(frame_colunas, text=f"Coluna do arquivo: {coluna}")
        label_coluna_arquivo.grid(row=i, column=0, padx=10, pady=5)

        # Entrada para o nome da coluna no banco de dados
        entry_coluna_banco = tk.Entry(frame_colunas)
        entry_coluna_banco.grid(row=i, column=1, padx=10, pady=5)

        # Armazena no dicionário para uso posterior
        colunas_mapping[coluna] = entry_coluna_banco

    # Função para gerar o script SQL
    def gerar_inserts():
        nome_tabela = entry_tabela.get().strip()
        if not nome_tabela:
            tk.messagebox.showwarning("Aviso", "Por favor, insira o nome da tabela.")
            return
        
        inserts = []
        for index, row in df.iterrows():
            # Construir a parte das colunas e dos valores do insert
            colunas = []
            valores = []
            for coluna_arquivo, entry_banco in colunas_mapping.items():
                coluna_banco = entry_banco.get().strip()
                if coluna_banco:
                    colunas.append(coluna_banco)
                    valor = row[coluna_arquivo]
                    # Adiciona tratamento para valores NULL
                    if pd.isna(valor):
                        valores.append("NULL")
                    else:
                        # Adiciona aspas em valores string e converte os outros para string
                        if isinstance(valor, str):
                            valores.append(f"'{valor}'")
                        else:
                            valores.append(str(valor))
            
            # Criar a query de insert
            insert_query = f"INSERT INTO {nome_tabela} ({', '.join(colunas)}) VALUES ({', '.join(valores)});"
            inserts.append(insert_query)

        # Salvar o script em um arquivo SQL
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

    # Botão para gerar os inserts
    btn_gerar = tk.Button(janela_inserts, text="Gerar Inserts", command=gerar_inserts)
    btn_gerar.pack(pady=10)

    # Botão para cancelar
    btn_cancelar = tk.Button(janela_inserts, text="Cancelar", command=janela_inserts.destroy)
    btn_cancelar.pack(pady=10)

    janela_inserts.geometry("400x300")  # Define o tamanho da janela

def log_error(message):
    logging.error(message)

def log_info(message):
    logging.info(message)

def validar_e_limpar_dados(df):
    try:
        df = df.fillna('NULL')
        df = df.drop_duplicates()
        log_info("Validação e limpeza de dados concluídas.")
        return df
    except Exception as e:
        log_error(f"Erro ao validar e limpar dados: {str(e)}")
        return None

# Configurações globais
configuracoes = carregar_configuracoes()
df = None  # Inicializa a variável global df

# Interface gráfica
root = tk.Tk()
root.title("Processamento de Dados")

# Adicionando uma barra de progresso
progress_bar = Progressbar(root, orient="horizontal", length=400, mode='determinate')
progress_bar.pack(pady=10)

# Criando Treeview para exibir os dados
treeview = ttk.Treeview(root)
treeview.pack(padx=20, pady=20)

# Criando o menu
criar_menu()

root.mainloop()
