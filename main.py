from tkinter import *
import tkinter as tk
import pandas as pd
from tkintertable import TableCanvas
from tkinter import ttk
import win32com.client as win32

def callback(*args):
    print(var.get())
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

dados = pd.read_excel('testeExcel.xlsx')
dados_dict = dados.to_dict()

janela = tk.Tk()
janela.title("Astro - Controle de Demanda")
janela.geometry("1000x500")
var = StringVar()
var.set("Option 1")


barra_menu = tk.Menu(janela)
janela.config(menu=barra_menu)
barra_menu.config(bd=1, relief="solid")

arquivo_menu = tk.Menu(barra_menu)
barra_menu.add_cascade(label="Arquivo ", menu=arquivo_menu)
sobre_menu = tk.Menu(barra_menu)
barra_menu.add_cascade(label="Sobre ", menu=sobre_menu)

arquivo_menu.add_command(label="Abrir")
arquivo_menu.add_command(label="Salvar")

label = Label(janela, text="Controle de Demanda", font=("Verdana", 16), pady=10)
label.pack()

caixa = tk.Frame(janela, bd=2, relief="sunken")
caixa.pack(padx=5, pady=5)

tabela = TableCanvas(caixa, data=dados_dict)
tabela.config(width=500, height=100)
tabela.show()

janela.update_idletasks()
largura = janela.winfo_width()
altura = janela.winfo_height()
x = (janela.winfo_screenwidth() // 2) - (largura // 2)
y = (janela.winfo_screenheight() // 2) - (altura // 2)
janela.geometry("{}x{}+{}+{}".format(largura, altura, x, y))

notebook = ttk.Notebook(janela)
notebook.pack(fill="both", expand="yes")

aba1 = ttk.Frame(notebook)
notebook.add(aba1, text="Enviar e Receber Demandas")

btnReceberEmail = tk.Button(aba1, text="Receber Demandas", width=20, height=1, command=receberDemanda)
btnReceberEmail.pack(side= 'left', padx=20)

label2 = Label(aba1, text="Enviar email para: ")
label2.pack(side="left", padx=50)

option = OptionMenu(aba1, var, "email1", "email 2", "email3", command=callback())
option.pack(side="left", padx=10)

btnEnviarEmail = tk.Button(aba1, text="Disparar Email", width=20, height=1)
btnEnviarEmail.pack(side="left", padx=10)

aba2 = ttk.Frame(notebook)
notebook.add(aba2, text="Gerenciamento de Equipe")

label2 = tk.Label(aba2, text="Conteúdo da aba 2")
label2.pack()

notebook.pack(expand=1, fill='both')

janela.mainloop()




