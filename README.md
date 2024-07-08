# email_handler

# português

O script Python fornecido (main.py) utiliza a biblioteca imap_tools para interagir com um servidor de email IMAP (Outlook neste caso) e realiza as seguintes tarefas:

Definição da Classe (ManipuladorDeEmails):
Inicializa uma conexão com a caixa de correio IMAP utilizando as credenciais do usuário.
Fornece métodos para:
Buscar emails com base em critérios de remetente e destinatário (buscar_emails_por_remetente_e_destinatario).
Buscar emails com anexos específicos (buscar_emails_com_anexo_e_sem_anexo).
Estes métodos coletam metadados do email como assunto, conteúdo do texto, remetente, detalhes do anexo e salvam os anexos localmente.

Funções:
verificar_arquivos_baixados: Verifica se todos os anexos foram baixados para uma pasta especificada.
verificar_planilha_escrita: Verifica se a planilha do Excel foi escrita com sucesso.

Execução Principal:
Conecta-se à caixa de correio e recupera emails com critérios específicos.
Armazena os dados coletados em um DataFrame do pandas.
Tentativas contínuas para salvar este DataFrame em um arquivo Excel até ser bem-sucedido.
Exibe uma mensagem ao finalizar com sucesso.
Instalação das Dependências:
Para executar este script, certifique-se de ter as seguintes dependências instaladas:

imap_tools: Gerencia interações IMAP.

Instalação via pip:
pip install imap-tools

pandas: Gerencia dados no formato de DataFrame.
Instalação via pip:
pip install pandas

tqdm: Exibe barras de progresso durante o processamento de emails e anexos.
Instalação via pip:
pip install tqdm

Executando o Script:
Salve o script Python fornecido (main.py) em seu ambiente local.
Certifique-se de ter Python instalado juntamente com o pip para gerenciar pacotes.
Instale as dependências listadas acima utilizando o pip.
Atualize o script com suas credenciais de email (usuario, senha), email do remetente (remetente) e título desejado para o anexo (titulo_anexo).

Execute o script (python main.py) e acompanhe as mensagens no console.
O script continuará executando até que todos os anexos sejam baixados e o arquivo Excel seja escrito com sucesso.
Este setup garante que você possa gerenciar e processar efetivamente emails com anexos de um servidor IMAP utilizando Python.

# inglish

The provided Python script (main.py) utilizes the imap_tools library to interact with an IMAP email server (Outlook in this case) and perform the following tasks:

Class Definition (ManipuladorDeEmails):

Initializes an IMAP mailbox connection with the user's credentials.

Provides methods to:
Fetch emails based on sender and recipient criteria (buscar_emails_por_remetente_e_destinatario).
Fetch emails with specific attachments (buscar_emails_com_anexo_e_sem_anexo).
These methods gather email metadata such as subject, text content, sender, attachment details, and save attachments locally.

Functions:
verificar_arquivos_baixados: Checks if all attachments have been downloaded to a specified folder.
verificar_planilha_escrita: Checks if the Excel spreadsheet has been successfully written.

Main Execution:
Connects to the mailbox and retrieves emails with specific criteria.
Stores fetched data in a pandas DataFrame.
Continuously attempts to save this DataFrame into an Excel file until successful.
Displays a message upon successful completion.
Dependencies Installation:
To run this script, ensure you have the following dependencies installed:

imap_tools: Handles IMAP interactions.

Install via pip:
pip install imap-tools

pandas: Manages data in DataFrame format.

Install via pip:
pip install pandas
tqdm: Displays progress bars during email and attachment processing.

Install via pip:
pip install tqdm

Running the Script:
Save the provided Python script (main.py) to your local environment.
Make sure you have Python installed along with pip for package management.
Install the dependencies listed above using pip.
Update the script with your email credentials (usuario, senha), sender email (remetente), and desired attachment title (titulo_anexo).
Run the script (python main.py) and monitor the console for progress messages.

The script will continue executing until all attachments are downloaded and the Excel file is successfully written.
This setup ensures you can effectively manage and process emails with attachments from an IMAP server using Python.