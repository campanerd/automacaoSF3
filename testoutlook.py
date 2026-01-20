import win32com.client as win32

# Abre o Outlook
outlook = win32.Dispatch('Outlook.Application')

# Cria um novo e-mail
mail = outlook.CreateItem(0)

# Configurações básicas
mail.To = "davi.fernandes@icrcobranca.com.br"      
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
