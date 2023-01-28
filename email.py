import win32com.client as win32

def enviarEmail():
     # Criar a integração com o outlook
     outlook = win32.Dispatch('outlook.application')

     # Criar um email
     email = outlook.CreateItem(0)

     # Configurar as informações do email
     email.To = "andrelopes.ads.dev@gmail.com"
     email.Subject = "teste email automatico python"
     email.HTMLBody = """
     <p>Este é um teste.</p>     
     <p>Controle de demanda.</>     
     <p>Segue as tasks.</p>
     """
     email.Send()
     print('Email enviado com sucesso!!!')

#Função para receber as demandas via "assunto" do email.
def receberDemanda():
     outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
     inbox = outlook.GetDefaultFolder(6)
     messages = inbox.Items
     unread_messages = messages.Restrict("[UnRead] = true")

     subjects = []
     for message in unread_messages:
         subjects.append(message.Subject)

     for subject in subjects:
         subject_words = subject.split()
         print(subject_words)







