import win32com.client as win32

# Criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# Criar um email
email = outlook.CreateItem(0)

# Configurar as informações do email
email.To = "dedlop@gmail.com"
email.Subject = "teste email automatico python"
email.HTMLBody = """
<p>Este é um teste.</p>

<p>Controle de demanda.</>

<p>Segue as tasks.</p>
"""

email.Send()
print('Email enviado com sucesso!!!')