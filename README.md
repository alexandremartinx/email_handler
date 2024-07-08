# email_handler

Description:
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