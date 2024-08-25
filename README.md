# LeitorPython
Leitor de Dados 4.0

Claro! Aqui está um exemplo de um README para o código fornecido:

---

# Processador de Arquivos Excel

Este é um aplicativo GUI desenvolvido em Python que permite processar arquivos Excel, aplicar filtros, selecionar colunas específicas, gerar relatórios e exportar dados em vários formatos. O aplicativo usa a biblioteca `tkinter` para a interface gráfica e `pandas` para manipulação de dados.

## Funcionalidades

1. **Carregar Arquivo Excel**
   - Selecione um arquivo Excel para processar.

2. **Validação e Limpeza de Dados**
   - Substitui valores `NaN` por `NULL`.
   - Remove duplicatas.

3. **Exibir Dados**
   - Mostra o DataFrame carregado em uma área de texto dentro da interface.

4. **Aplicar Filtros**
   - Permite a seleção de uma coluna para aplicar um filtro.
   - Exibe uma lista vertical com valores distintos da coluna selecionada para escolher o valor de filtro.

5. **Selecionar Colunas**
   - Escolha colunas específicas para exibir ou processe todas as colunas.

6. **Gerar Relatório**
   - Gera um relatório com o número de registros e nomes das colunas.

7. **Exportar Dados**
   - Exporte os dados em formatos `CSV`, `XLSX` ou `TXT`.

8. **Criar Comandos INSERT SQL**
   - Gera um arquivo SQL com comandos `INSERT` para diferentes tipos de banco de dados (MySQL, PostgreSQL, SQLite, Oracle).

9. **Exportar e Importar Configurações**
   - Salve e carregue configurações em arquivos JSON.

## Requisitos

- Python 3.x
- Bibliotecas:
  - `pandas`
  - `tkinter` (geralmente incluído com Python)
  - `openpyxl` (para trabalhar com arquivos Excel)

## Instalação

1. **Clone o Repositório**

   ```bash
   git clone https://github.com/seu_usuario/processador-arquivos-excel.git
   ```

2. **Instale as Dependências**

   Crie um ambiente virtual e instale as dependências:

   ```bash
   python -m venv env
   source env/bin/activate  # No Windows use `env\Scripts\activate`
   pip install pandas openpyxl
   ```

## Uso

1. **Executar o Aplicativo**

   Execute o script `main.py` para iniciar o aplicativo:

   ```bash
   python main.py
   ```

2. **Carregar e Processar Arquivo**

   - Clique em "Processar Arquivo" para abrir um diálogo de seleção de arquivos.
   - Escolha um arquivo Excel e aguarde o processamento.

3. **Aplicar Filtros**

   - Selecione se deseja aplicar um filtro.
   - Escolha a coluna e o valor desejado a partir da lista vertical exibida.

4. **Selecionar Colunas**

   - Escolha se deseja selecionar colunas específicas ou todas.

5. **Gerar Relatório**

   - O relatório será exibido com o número de registros e as colunas.

6. **Exportar Dados**

   - Escolha o formato de exportação e o nome do arquivo para salvar.

7. **Criar Comandos INSERT SQL**

   - Informe o tipo de banco de dados e o nome da tabela para gerar o arquivo SQL.

8. **Exportar e Importar Configurações**

   - Salve e carregue configurações em arquivos JSON conforme necessário.

## Notas

- Certifique-se de que o arquivo Excel não contenha colunas com nomes "Unnamed" não desejados, pois elas serão removidas automaticamente.
- Assegure-se de que os tipos de dados estejam corretos para evitar erros na criação do SQL.

## Contribuição

Se você deseja contribuir para este projeto, sinta-se à vontade para abrir um pull request ou relatar problemas. Agradecemos suas contribuições!

## Licença

Este projeto é licenciado sob a [Licença MIT](https://opensource.org/licenses/MIT). Veja o arquivo LICENSE para mais detalhes.

--
