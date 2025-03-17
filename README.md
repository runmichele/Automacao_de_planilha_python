# Automação de Planilhas com Python e Envio por E-mail

Este projeto consiste em uma automação feita com Python utilizando as bibliotecas pandas e openpyxl. O objetivo é importar uma planilha em Excel, realizar tratamento dos dados, adicionar gráficos, aplicar fórmulas e, por fim, enviar o arquivo final por e-mail de forma automática.

# 📁 Estrutura do Projeto

📂 data/
  ├── 📄 VendaCarros.xlsx
  ├─ 1-import_data.py
  ├── 2-pivot_table.py
  ├─ 3-sheet_ready.py
  ├─ 4-add_charty.py
  ├─ 5-formulas.py
  ├─ 6-email.py
  ├─ barchart.xlsx
  ├─ pivot_table.xlsx
  ├─ arquivo.xlsx
  ├─ senha.txt
  └─ test.xlsx

# ⚙️ Detalhes dos scripts

📌 1-import_data.py

Importa e trata os dados iniciais da planilha Excel utilizando pandas.

# 2-pivot_table.py

Cria tabelas dinâmicas (pivot tables) com os dados tratados.

# 3-sheet_ready.py

Ajusta e prepara as planilhas para visualização e análise.

# 4-add_charty.py

Adiciona gráficos aos dados no arquivo Excel utilizando openpyxl.

# 5-formulas.py

Insere fórmulas no Excel utilizando Python (openpyxl).

# 6-email.py

Envia o arquivo final como anexo por e-mail, utilizando SMTP.

Bibliotecas: smtplib, ssl, mimetypes, email.

Carrega uma senha através de um arquivo de texto (senha.txt).

Envia o arquivo final anexado via Gmail.

📌 Bibliotecas e tecnologias utilizadas:

Python

Pandas

Openpyxl

SMTPLib

SSL

EmailMessage

#  📂 Estrutura de Arquivos

data/
- 1-import_data.py
- 2-pivot_table.py
- 3-sheet_ready.py
- 4-add_charty.py
- 5-formulas.py
- 6-email.py
- barchart.xlsx
- pivot_table.xlsx
- senha.txt
- VendaCarros.xlsx
- arquivo.xlsx
- test.xlsx


# Como utilizar o projeto?

Clone este repositório:

git clone https://github.com/runmichele/Automacao_de_planilha_python.git

Instale as dependências necessárias:

pip install pandas openpyxl

Execute cada arquivo Python na ordem indicada pelo número do arquivo.

Configure suas credenciais de e-mail no arquivo senha.txt e no script de e-mail antes de executar.

#  🎯 Resultado esperado

* Dados tratados e analisados em tabelas dinâmicas.

* Gráficos gerados automaticamente.

* Planilha pronta com fórmulas aplicadas.

* Envio automatizado por e-mail com o arquivo resultante.
