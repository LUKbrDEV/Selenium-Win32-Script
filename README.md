Automação com Selenium e Outlook

Este repositório contém dois scripts em Python que automatizam tarefas comuns: abertura de chamados em sistema interno via Selenium e extração de anexos PDF de e-mails no Outlook.

📌 Código 1: Automação de Chamados com Selenium

Este script utiliza Selenium WebDriver e Pandas para automatizar o processo de abertura de chamados em um sistema interno.

Funcionalidades

Carrega dados de um arquivo CSV (chamados.csv).

Preenche automaticamente os campos do formulário de abertura de chamado.

Anexa arquivos PDF relacionados.

Envia os chamados e registra no console o fornecedor correspondente.

Tecnologias utilizadas

Python

Pandas

Selenium WebDriver

Como usar

Instale as dependências:

pip install pandas selenium

Configure o Chrome com depuração remota (127.0.0.1:9222).

Ajuste o caminho do arquivo CSV no código.

Execute o script:

python automacao_chamados.py

📌 Código 2: Automação de Extração de PDFs no Outlook

Este script utiliza win32com para acessar o Outlook e salvar anexos PDF de e-mails não lidos em uma pasta local.

Funcionalidades

Conecta ao Outlook via MAPI.

Localiza uma pasta específica de e-mails (ImpressCN).

Filtra mensagens não lidas.

Salva anexos PDF em uma pasta local (C:\temp).

Marca os e-mails como lidos após o processamento.

Tecnologias utilizadas

Python

win32com.client

Outlook MAPI

Como usar

Instale a biblioteca necessária:

pip install pywin32

Configure o nome da pasta e store do Outlook no código.

Execute o script:

python extracao_outlook.py

📂 Estrutura do Repositório

automacao-outlook-selenium/
├── automacao_chamados.py
├── extracao_outlook.py
├── chamados.csv
└── README.md

⚠️ Observações

Os scripts são exemplos práticos e podem precisar de ajustes conforme o ambiente.

Certifique-se de ter permissões adequadas para acessar o Outlook e o sistema interno.
