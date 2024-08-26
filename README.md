# LeitorPython
Leitor de Dados 4.0
---

# Processamento de Dados - README

## Descrição

Este aplicativo GUI em Python permite processar arquivos Excel, com funcionalidades para carregar, filtrar, exportar e gerar scripts SQL a partir dos dados. A interface gráfica é construída usando a biblioteca `tkinter`, e os dados são manipulados com a biblioteca `pandas`.

## Funcionalidades

- **Abrir Arquivo**: Carrega um arquivo Excel (.xlsx ou .xls) e exibe seus dados em uma tabela.
- **Configurações**: Permite configurar o diretório padrão, formato de data e banco de dados preferido.
- **Filtrar Dados**: Filtra os dados carregados com base em uma coluna e valor especificados.
- **Exportar Dados**: Exporta os dados filtrados para formatos CSV, Excel ou TXT.
- **Criar Inserts SQL**: Gera um script SQL para inserir dados em uma tabela de banco de dados.

## Requisitos

- Python 3.x
- Bibliotecas: `tkinter`, `pandas`, `openpyxl` (para suporte a arquivos Excel), `json`, `logging`

## Instalação

1. **Clone o repositório ou baixe o código**:
   ```bash
   git clone <URL_DO_REPOSITORIO>
   ```

2. **Instale as dependências**:
   Execute o seguinte comando para instalar as bibliotecas necessárias:
   ```bash
   pip install pandas openpyxl
   ```

## Uso

1. **Executar o aplicativo**:
   Navegue até o diretório onde o código está localizado e execute:
   ```bash
   python seu_script.py
   ```

2. **Abrir Arquivo**:
   - Vá para o menu "Arquivo" e selecione "Abrir Arquivo".
   - Escolha um arquivo Excel para carregar. Os dados serão exibidos na tabela.

3. **Configurações**:
   - Vá para o menu "Arquivo" e selecione "Configurações".
   - Modifique as configurações conforme necessário e salve as alterações.

4. **Filtrar Dados**:
   - Vá para o menu "Processar" e selecione "Filtrar Dados".
   - Escolha a coluna e o valor para filtrar os dados carregados.

5. **Exportar Dados**:
   - Vá para o menu "Processar" e selecione "Exportar Dados".
   - Escolha o formato desejado (CSV, Excel ou TXT) e salve o arquivo.

6. **Criar Inserts SQL**:
   - Vá para o menu "Processar" e selecione "Criar Inserts SQL".
   - Forneça o nome da tabela e mapeie as colunas do arquivo para as colunas do banco de dados.
   - O script SQL será salvo em um arquivo .sql.

## Estrutura do Código

- **Funções Principais**:
  - `carregar_configuracoes()`: Carrega as configurações do arquivo JSON ou cria configurações padrão.
  - `salvar_configuracoes(configuracoes)`: Salva as configurações no arquivo JSON.
  - `abrir_configuracoes_gui()`: Interface gráfica para modificar as configurações.
  - `processar_arquivo_gui()`: Carrega e exibe um arquivo Excel, e valida e limpa os dados.
  - `mostrar_dados(df)`: Exibe os dados carregados em uma tabela (Treeview).
  - `filtrar_dados_gui()`: Interface gráfica para filtrar dados.
  - `exportar_dados_gui()`: Interface gráfica para exportar dados em diferentes formatos.
  - `criar_inserts_gui()`: Interface gráfica para gerar e salvar um script SQL.

- **Logging**:
  - Logs de erros e informações são salvos em um arquivo `logfile.log` na área de trabalho do usuário.

## Contribuição

Se você deseja contribuir para o projeto, por favor, siga estas etapas:
1. Faça um fork do repositório.
2. Crie uma nova branch (`git checkout -b feature/nova-funcionalidade`).
3. Faça suas alterações e adicione commits (`git commit -am 'Adiciona nova funcionalidade'`).
4. Faça um push para a branch (`git push origin feature/nova-funcionalidade`).
5. Crie um Pull Request.

## Licença

Este projeto está licenciado sob a Licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

---

