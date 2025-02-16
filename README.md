# PROJETO-6-ANALISE-DE-FATURAMENTO

# Relatório de Vendas por Loja

Este repositório contém um script Python que gera um relatório detalhado de vendas por loja a partir de uma base de dados do Excel e envia os resultados por e-mail utilizando o Microsoft Outlook.

## Funcionalidades

- **Leitura de dados**: Importa os dados de vendas de um arquivo Excel.
- **Análises de vendas**:
  - Calcula o faturamento total por loja.
  - Determina a quantidade de produtos vendidos por loja.
  - Calcula o ticket médio por produto em cada loja.
- **Envio de relatório por e-mail**:
  - Gera um relatório HTML com as análises realizadas.
  - Envia o relatório automaticamente via Microsoft Outlook.

## Tecnologias Utilizadas

- **Python**: Linguagem de programação utilizada para processar os dados.
- **Pandas**: Biblioteca utilizada para manipulação e análise de dados.
- **Win32com**: Biblioteca utilizada para integração com o Microsoft Outlook e envio de e-mails.

## Requisitos

Antes de executar o script, certifique-se de que seu ambiente está configurado corretamente:

1. **Python 3.x instalado**: [Baixar Python](https://www.python.org/downloads/)
2. **Bibliotecas necessárias**:
   - Instale as dependências com o comando:
     ```bash
     pip install pandas pywin32
     ```
3. **Microsoft Outlook configurado**:
   - O script utiliza o Outlook para enviar e-mails, então o cliente de e-mail deve estar configurado corretamente no seu computador.
4. **Arquivo Excel (`Vendas.xlsx`)**:
   - O arquivo `Vendas.xlsx` deve estar na mesma pasta que o script e conter as seguintes colunas:
     - `ID Loja`: Identificação da loja.
     - `Valor Final`: Valor total da venda.
     - `Quantidade`: Quantidade de produtos vendidos.

## Como Usar

1. Coloque o arquivo `Vendas.xlsx` na mesma pasta que o script.
2. Edite o e-mail de destino no script para o endereço desejado:
   ```python
   mail.To = 'fotonovaemail@gmail.com'

