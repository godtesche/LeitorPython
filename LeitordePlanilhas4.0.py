import os
import pandas as pd
import logging
import json
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, Scrollbar

# Configuração do logging
logging.basicConfig(filename=os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'logfile.log'),
                    level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

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

def selecionar_colunas(df, colunas):
    try:
        if colunas == 'todas':
            return df
        else:
            colunas_selecionadas = [col.strip() for col in colunas.split(',')]
            return df[colunas_selecionadas]
    except Exception as e:
        log_error(f"Erro ao selecionar colunas: {str(e)}")
        return None

def aplicar_filtro(df, coluna_filtro, valor_filtro):
    try:
        df_filtrado = df[df[coluna_filtro] == valor_filtro]
        return df_filtrado
    except Exception as e:
        log_error(f"Erro ao aplicar filtro: {str(e)}")
        return None

def salvar_arquivo(df, nome_arquivo, formato):
    try:
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        caminho_completo = os.path.join(desktop, f"{nome_arquivo}.{formato}")

        if formato == 'csv':
            df.to_csv(caminho_completo, index=False, sep=';')
        elif formato == 'xlsx':
            df.to_excel(caminho_completo, index=False)
        elif formato == 'txt':
            df.to_csv(caminho_completo, index=False, sep='\t')
        else:
            raise ValueError("Formato de exportação inválido.")
        
        log_info(f"Arquivo salvo em: {caminho_completo}")
        return caminho_completo
    except Exception as e:
        log_error(f"Erro ao salvar arquivo: {str(e)}")
        return None

def criar_insert(df, nome_tabela, tipo_banco):
    try:
        colunas_tabela = [col.strip() for col in df.columns]
        nome_arquivo_sql = f"{nome_tabela}.sql"
        caminho_arquivo_sql = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', nome_arquivo_sql)
        
        with open(caminho_arquivo_sql, 'w') as file:
            for index, row in df.iterrows():
                valores = [formatar_valor(value, tipo_banco) for value in row]
                insert_sql = f"INSERT INTO {nome_tabela} ({', '.join(colunas_tabela)}) VALUES ({', '.join(valores)});\n"
                file.write(insert_sql)
        
        log_info(f"Arquivo SQL salvo em: {caminho_arquivo_sql}")
        return caminho_arquivo_sql
    except Exception as e:
        log_error(f"Erro ao criar comandos INSERT: {str(e)}")
        return None

def formatar_valor(value, tipo_banco):
    if pd.isna(value):
        return "NULL"
    elif isinstance(value, str):
        return f"'{value.replace("'", "''")}'"
    elif isinstance(value, pd.Timestamp):
        return formatar_data(value, tipo_banco)
    else:
        return str(value)

def formatar_data(value, tipo_banco):
    if tipo_banco in ['mysql', 'postgresql', 'sqlite']:
        return f"'{value.strftime('%Y-%m-%d %H:%M:%S')}'"
    elif tipo_banco == 'oracle':
        return f"TO_DATE('{value.strftime('%Y-%m-%d %H:%M:%S')}', 'YYYY-MM-DD HH24:MI:SS')"
    else:
        raise ValueError("Tipo de banco de dados não suportado para datas.")

def exportar_configuracoes(configuracoes, nome_arquivo):
    try:
        caminho_arquivo = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', f"{nome_arquivo}.json")
        with open(caminho_arquivo, 'w') as file:
            json.dump(configuracoes, file, indent=4)
        log_info(f"Configurações exportadas para: {caminho_arquivo}")
        return caminho_arquivo
    except Exception as e:
        log_error(f"Erro ao exportar configurações: {str(e)}")
        return None

def importar_configuracoes(nome_arquivo):
    try:
        caminho_arquivo = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', f"{nome_arquivo}.json")
        with open(caminho_arquivo, 'r') as file:
            configuracoes = json.load(file)
        log_info(f"Configurações importadas de: {caminho_arquivo}")
        return configuracoes
    except Exception as e:
        log_error(f"Erro ao importar configurações: {str(e)}")
        return {}

def gerar_relatorio(df):
    try:
        num_registros = len(df)
        colunas = list(df.columns)
        messagebox.showinfo("Relatório", f"Número de registros: {num_registros}\nColunas: {', '.join(colunas)}")
        log_info(f"Relatório gerado: {num_registros} registros, Colunas: {', '.join(colunas)}")
    except Exception as e:
        log_error(f"Erro ao gerar relatório: {str(e)}")
        messagebox.showerror("Erro", f"Erro ao gerar relatório: {str(e)}")

def mostrar_dados(df, text_widget):
    try:
        text_widget.delete(1.0, tk.END)
        df_str = df.to_string(index=False)
        text_widget.insert(tk.END, df_str)
    except Exception as e:
        log_error(f"Erro ao mostrar dados: {str(e)}")
        messagebox.showerror("Erro", f"Erro ao mostrar dados: {str(e)}")

def aplicar_filtro_gui(df):
    try:
        coluna_filtro = simpledialog.askstring("Filtro", "Informe o nome da coluna para aplicar o filtro:")
        if coluna_filtro not in df.columns:
            messagebox.showerror("Erro", "Coluna não encontrada.")
            return None
        
        valores_distintos = df[coluna_filtro].dropna().unique()
        valores_distintos = sorted(set(valores_distintos))
        
        # Criar uma nova janela para a seleção do filtro
        filtro_window = tk.Toplevel(root)
        filtro_window.title("Selecionar Valor de Filtro")

        # Criar a Listbox e o Scrollbar
        listbox = tk.Listbox(filtro_window, selectmode=tk.SINGLE, height=15, width=50)
        scrollbar = Scrollbar(filtro_window, orient=tk.VERTICAL, command=listbox.yview)
        
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)
        
        # Adicionar valores distintos à Listbox
        for valor in valores_distintos:
            listbox.insert(tk.END, valor)
        
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        def aplicar():
            selecionado = listbox.get(tk.ACTIVE)
            if selecionado:
                filtro_window.destroy()
                return selecionado
            else:
                messagebox.showerror("Erro", "Nenhum valor selecionado.")
                return None
        
        apply_button = tk.Button(filtro_window, text="Aplicar Filtro", command=lambda: aplicar())
        apply_button.pack(pady=10)

        # Aguardar o fechamento da janela
        root.wait_window(filtro_window)
        
        return aplicar()
    
    except Exception as e:
        log_error(f"Erro ao aplicar filtro com GUI: {str(e)}")
        messagebox.showerror("Erro", f"Erro ao aplicar filtro com GUI: {str(e)}")
        return None

def processar_arquivo_gui():
    try:
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not caminho_arquivo:
            return
        
        df = pd.read_excel(caminho_arquivo)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]  # Remove colunas "Unnamed"
        df = validar_e_limpar_dados(df)
        if df is None:
            messagebox.showerror("Erro", "Erro na validação e limpeza de dados.")
            return
        
        mostrar_dados(df, text_area)

        aplicar_filtro_decisao = messagebox.askyesno("Aplicar Filtro", "Deseja aplicar um filtro nos dados?")
        if aplicar_filtro_decisao:
            valor_filtro = aplicar_filtro_gui(df)
            if valor_filtro:
                df = aplicar_filtro(df, coluna_filtro, valor_filtro)
                if df is None:
                    messagebox.showerror("Erro", "Erro ao aplicar filtro.")
                    return
                mostrar_dados(df, text_area)

        selecionar_colunas_decisao = messagebox.askyesno("Selecionar Colunas", "Deseja selecionar colunas específicas?")
        if selecionar_colunas_decisao:
            colunas = simpledialog.askstring("Selecionar Colunas", "Informe os nomes das colunas separados por vírgula ou 'todas' para selecionar todas:")
            df = selecionar_colunas(df, colunas)
            if df is None:
                messagebox.showerror("Erro", "Erro ao selecionar colunas.")
                return
            mostrar_dados(df, text_area)

        gerar_relatorio(df)

        acao = simpledialog.askstring("Ação", "Você deseja criar um insert SQL, exportar as colunas ou exportar/importar configurações? (insert/exportar/configuracoes):").strip().lower()

        if acao == 'insert':
            tipo_banco = simpledialog.askstring("Tipo de Banco", "Informe o tipo de banco de dados (mysql/postgresql/sqlite/oracle):").strip().lower()
            if tipo_banco not in ['mysql', 'postgresql', 'sqlite', 'oracle']:
                messagebox.showerror("Erro", "Tipo de banco de dados não suportado.")
                return
            nome_tabela = simpledialog.askstring("Nome da Tabela", "Informe o nome da tabela:")
            caminho_arquivo_sql = criar_insert(df, nome_tabela, tipo_banco)
            if caminho_arquivo_sql:
                messagebox.showinfo("Sucesso", f"Arquivo SQL criado com sucesso: {caminho_arquivo_sql}")
        
        elif acao == 'exportar':
            formato = simpledialog.askstring("Formato", "Informe o formato de exportação (csv/xlsx/txt):").strip().lower()
            if formato not in ['csv', 'xlsx', 'txt']:
                messagebox.showerror("Erro", "Formato de exportação não suportado.")
                return
            nome_arquivo = simpledialog.askstring("Nome do Arquivo", "Informe o nome do arquivo de exportação:")
            caminho_arquivo = salvar_arquivo(df, nome_arquivo, formato)
            if caminho_arquivo:
                messagebox.showinfo("Sucesso", f"Arquivo exportado com sucesso: {caminho_arquivo}")

        elif acao == 'configuracoes':
            operacao = simpledialog.askstring("Configurações", "Deseja exportar ou importar configurações? (exportar/importar):").strip().lower()
            nome_arquivo = simpledialog.askstring("Nome do Arquivo", "Informe o nome do arquivo de configurações:")
            
            if operacao == 'exportar':
                configuracoes = {
                    "colunas_selecionadas": list(df.columns),
                    "filtros_aplicados": "Coluna e valor de filtro atual não implementados",
                }
                caminho_arquivo_json = exportar_configuracoes(configuracoes, nome_arquivo)
                if caminho_arquivo_json:
                    messagebox.showinfo("Sucesso", f"Configurações exportadas com sucesso: {caminho_arquivo_json}")
            
            elif operacao == 'importar':
                configuracoes = importar_configuracoes(nome_arquivo)
                if configuracoes:
                    messagebox.showinfo("Configurações Importadas", json.dumps(configuracoes, indent=4))

    except Exception as e:
        log_error(f"Erro durante o processamento do arquivo: {str(e)}")
        messagebox.showerror("Erro", f"Erro durante o processamento do arquivo: {str(e)}")

# Configurar a GUI
root = tk.Tk()
root.title("Processador de Dados")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Área de texto para exibir o DataFrame
text_area = tk.Text(frame, wrap='none', height=20, width=100)
text_area.pack(side='left')

scrollbar = Scrollbar(frame, orient='vertical', command=text_area.yview)
scrollbar.pack(side='right', fill='y')

text_area.config(yscrollcommand=scrollbar.set)

botao_processar = tk.Button(root, text="Processar Arquivo", command=processar_arquivo_gui)
botao_processar.pack(pady=10)

root.mainloop()
