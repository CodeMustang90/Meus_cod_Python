# IMAP - Protocolo de acesso a e-mails
from imap_tools import MailBox, AND
import time

 
def extrair_anexos(usuario, senha_email, destinatario):
    
    with MailBox('imap.gmail.com').login(usuario, senha_email) as meu_email:
        meu_email.folder.set('[Gmail]/E-mails enviados')
        for i, email in enumerate(meu_email.fetch(AND(to=destinatario))):
                if len(email.attachments) > 0:
                    for anexo in email.attachments:
                        with open(f'Jupyter\\Integração E-mail com Python\\arquivos\\E-mail {i+1} - {anexo.filename}', 'wb') as arquivo:
                            arquivo.write(anexo.payload)
                        print(anexo.filename)

def mein(): 
    while True: 
        usuario = input('Digite seu e-mail: ')
        senha_email = input('Digite sua senha: ')
        destinatario = input('Digite o destinatário: ')
        if usuario and senha_email and destinatario:
            print('Extraindo anexos...')
            time.sleep(1)
            extrair_anexos(usuario, senha_email, destinatario)
            print('Anexos extraídos!')
            break
        else:
            print('Informaçõs inválidas')
            continue
if __name__ == '__main__':
    mein()
