import win32com.client as win32
import os
from dotenv import load_dotenv
load_dotenv()

# Abre o Outlook
outlook = win32.Dispatch('Outlook.Application')

# Cria um novo e-mail
mail = outlook.CreateItem(0)

# Configurações básicas
EMAIL_TO = os.getenv("EMAIL_TO")
mail.To = EMAIL_TO 
mail.Subject = "Teste de envio automático - Python"
mail.Body = (
    "Olá,\n\n"
    "Este é um e-mail de teste enviado automaticamente pelo Python "
    "utilizando o Outlook Desktop.\n\n"
    "Se você recebeu esta mensagem, o envio está funcionando corretamente.\n\n"
    "Atenciosamente,\n"
    "Automação Python"
)

# Envia o e-mail
mail.Send()

print("E-mail enviado com sucesso!")
