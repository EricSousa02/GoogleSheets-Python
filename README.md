Este é um script Python que automatiza o cálculo do status do aluno e da nota para aprovação final em uma planilha do Google Sheets. Antes de usar o script, é necessário configurar a API do Google Sheets e instalar as dependências necessárias. Siga o tutorial abaixo:

# Configuração da API do Google Sheets
- [Acesse o Tutorial da API do Google Sheets em Python](https://developers.google.com/sheets/api/quickstart/python?hl=pt-br)
- Siga as instruções na seção "Passos Iniciais" para criar um projeto no Console de Desenvolvedores do Google e ativar a API do Google Sheets.
- Baixe o arquivo de credenciais JSON (credentials.json) e salve-o no mesmo diretório do script.
- Execute o seguinte comando no terminal para instalar as bibliotecas necessárias:

- pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

Execute o script Python usando o seguinte comando:

- python main.py

Certifique-se de que o arquivo token.json seja gerado no mesmo diretório para armazenar as credenciais de acesso.

# Detalhes do Script
- O script utiliza a API do Google Sheets para ler e escrever dados na planilha.
- O identificador da planilha (SPREADSHEET_ID) e o intervalo de células (RANGE_NAME) são definidos no início do script e podem ser ajustados conforme necessário.
- A função calculate_status determina o status do aluno com base nas notas e na presença.
- A função calculate_naf calcula a nota para aprovação final em caso de exame final.
- O script adiciona colunas para o status do aluno e a nota para aprovação final na planilha.
- Lembre-se de que o script realiza alterações na planilha, então use-o com cautela e faça backup dos dados, se necessário. Em caso de erro, mensagens de erro serão exibidas no console.

Observação: Certifique-se de que a conta associada ao projeto da API do Google Sheets tenha acesso de edição à planilha especificada.

Aviso: O script está configurado para executar em ambiente local, pois utiliza o fluxo de autorização do aplicativo instalado. Certifique-se de ter as bibliotecas necessárias instaladas e siga as instruções do tutorial para configurar corretamente a API.
