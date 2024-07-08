from imap_tools import MailBox, AND
import pandas as pd
import os
from tqdm import tqdm
from time import sleep

class ManipuladorDeEmails:
    def __init__(self, usuario, senha, servidor='imap-mail.outlook.com'):
        self.usuario = usuario
        self.senha = senha
        self.servidor = servidor
        self.caixa_de_entrada = MailBox(self.servidor).login(self.usuario, self.senha)

    def buscar_emails_por_remetente_e_destinatario(self, remetente, destinatario):
        emails = list(self.caixa_de_entrada.fetch(AND(from_=remetente, to=destinatario)))
        lista_emails = []
        for email in tqdm(emails, desc="Buscando emails por remetente e destinatário"):
            lista_emails.append({
                "Remetente": remetente,
                "Assunto": email.subject,
                "Texto": email.text
            })
        return lista_emails

    def buscar_emails_com_anexo_e_sem_anexo(self, remetente, titulo_anexo, pasta_salvar):
        emails = list(self.caixa_de_entrada.fetch(AND(from_=remetente)))
        lista_dados_emails = []
        if not os.path.exists(pasta_salvar):
            os.makedirs(pasta_salvar)
        
        for email in tqdm(emails, desc="Buscando emails"):
            dados_email = {
                "Remetente": remetente,
                "Assunto": email.subject,
                "Texto": email.text,
                "Caminho do Anexo": None,
                "Tipo de Conteúdo": None
            }

            if len(email.attachments) > 0:
                for anexo in tqdm(email.attachments, desc="Processando anexos", leave=False):
                    if titulo_anexo in anexo.filename:
                        caminho_arquivo = os.path.join(pasta_salvar, anexo.filename)
                        with open(caminho_arquivo, 'wb') as arquivo:
                            arquivo.write(anexo.payload)
                        dados_email["Caminho do Anexo"] = caminho_arquivo
                        dados_email["Tipo de Conteúdo"] = anexo.content_type

            lista_dados_emails.append(dados_email)

        return lista_dados_emails

def verificar_arquivos_baixados(pasta_anexos, esperar=10):
    while True:
        if os.path.exists(pasta_anexos) and os.listdir(pasta_anexos):
            return True
        else:
            print(f"Aguardando arquivos serem baixados...")
            sleep(esperar)

def verificar_planilha_escrita(arquivo_planilha, esperar=10):
    while True:
        if os.path.exists(arquivo_planilha):
            return True
        else:
            print(f"Aguardando planilha ser escrita...")
            sleep(esperar)

if __name__ == "__main__":
    usuario = "seuemail@hotmail.com"
    senha = "sua senha"
    remetente = 'email do remente'
    pasta_salvar_anexos = "anexos_baixados"
    arquivo_planilha = "resultado_emails.xlsx"

    manipulador_de_emails = ManipuladorDeEmails(usuario, senha)

    while True:
        try:
            dados_emails = manipulador_de_emails.buscar_emails_com_anexo_e_sem_anexo(remetente, "curriculo", pasta_salvar_anexos)
            break
        except Exception as e:
            print(f"Erro ao buscar emails: {e}")
            print("Tentando novamente...")

    df_dados_emails = pd.DataFrame(dados_emails)

    while True:
        try:
            with pd.ExcelWriter(arquivo_planilha) as writer:
                df_dados_emails.to_excel(writer, sheet_name='Emails', index=False)
            break
        except Exception as e:
            print(f"Erro ao salvar planilha: {e}")
            print("Tentando novamente...")

    print(f"Resultados salvos em '{arquivo_planilha}'")
