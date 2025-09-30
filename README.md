# AutomacaoForecast
Projeto de automação para formatação de planilhas

# Ferramenta de Automação de Relatórios Financeiros

![Python Version](https://img.shields.io/badge/python-3.9%2B-blue)

Uma aplicação web desenvolvida em Python para automatizar a consolidação e o processamento de relatórios, transformando um fluxo de trabalho manual e demorado em um processo de poucos cliques.

## 📜 Sobre o Projeto

Este projeto foi criado para resolver um desafio comum em departamentos financeiros: a necessidade de consolidar dados de múltiplas fontes (planilhas Excel) em um único relatório mestre. O processo manual era repetitivo, propenso a erros e consumia um tempo valioso que poderia ser usado para análises mais estratégicas.

A solução é uma aplicação web interna que oferece uma interface simples para que usuários não-técnicos possam fazer o upload dos arquivos brutos e receber, em segundos, o relatório final consolidado e formatado.

## ✨ Funcionalidades

* **Interface Web Amigável:** Permite o uso da ferramenta sem necessidade de conhecimento técnico.
* **Upload de Múltiplos Arquivos:** Suporte para o envio simultâneo dos arquivos de entrada (DRE, Forecasts, etc.).
* **Processamento Automatizado:** Toda a lógica de negócio, incluindo limpeza, transformação, cálculos e consolidação dos dados, é executada automaticamente no backend com a biblioteca Pandas.
* **Geração de Relatório Final:** Cria um arquivo Excel (`.xlsx`) consolidado e pronto para análise.
* **Download Direto:** O usuário pode baixar o resultado diretamente pela interface ao final do processo.

## 🛠️ Tecnologias Utilizadas

O projeto foi construído utilizando as seguintes tecnologias:

* **Backend:**
    * [Python](https://www.python.org/)
    * [Flask](https://flask.palletsprojects.com/): Micro-framework para a criação do servidor web.
    * [Pandas](https://pandas.pydata.org/): Biblioteca para manipulação e análise dos dados.
    * [Openpyxl](https://openpyxl.readthedocs.io/): Dependência do Pandas para manipulação de arquivos `.xlsx`.
* **Frontend:**
    * HTML5
    * CSS3

## 📁 Estrutura de Pastas

O repositório está organizado da seguinte forma para garantir a separação de responsabilidades:

```
/Automacao_Web/
|
|-- app.py                # Servidor web (Backend Flask)
|-- processador.py        # Módulo com toda a lógica de processamento de dados
|-- requirements.txt      # Lista de dependências do Python
|
|-- /templates/           # Arquivos HTML da interface
|   |-- index.html
|   |-- sucesso.html
|
|-- /uploads/             # Pasta temporária para os arquivos enviados
|-- /output/              # Pasta onde o relatório final é salvo
|-- /arquivados_forecast/ # Pasta para arquivar os arquivos já processados
|
`-- README.md             # Documentação do projeto
```


## 📖 Como Usar

1.  Acesse a aplicação pelo navegador.
2.  Clique nos botões "Escolher arquivo" para selecionar os arquivos DRE e Forecast (Irrestrito e Restrito) do seu computador.
3.  Clique no botão **"Processar Arquivos"**.
4.  Aguarde o processamento. Você será redirecionado para uma página de sucesso.
5.  Clique no botão **"Baixar Banco
