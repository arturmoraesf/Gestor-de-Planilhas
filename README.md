# Gestor de Planilhas

## Descrição
Este projeto consiste em um aplicativo desktop desenvolvido em Python com a biblioteca Tkinter, que permite gerar, comparar e gerenciar planilhas Excel de documentos. Ele oferece funcionalidades para criar uma planilha a partir dos arquivos de uma pasta, comparar planilhas e deletar documentos duplicados.

## Funcionalidades
- **Gerar Planilha**: Lê os arquivos de uma pasta e gera uma planilha Excel contendo seus nomes e formatos.
- **Comparar Planilhas**: Compara a planilha gerada com outra planilha fornecida pelo usuário, marcando os documentos encontrados.
- **Deletar Documentos Repetidos**: Remove arquivos duplicados da pasta com base na comparação entre as planilhas.

## Tecnologias Utilizadas
- Python 3
- Tkinter (Interface Gráfica)
- Pandas (Manipulação de Dados)
- OpenPyXL (Manipulação de Arquivos Excel)

## Instalação
1. Instale as dependências necessárias:
   ```sh
   pip install pandas openpyxl
   ```

## Como Usar
1. Execute o script principal:
   ```sh
   python script.py
   ```
2. Escolha a funcionalidade desejada na interface:
   - **Gerar Planilha**: Selecione uma pasta para criar uma planilha com os documentos listados.
   - **Comparar Planilhas**: Selecione um arquivo Excel para comparação.
   - **Deletar Notas Lançadas**: Escolha a pasta de arquivos e remova os duplicados.
