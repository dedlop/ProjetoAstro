from tkinter import *
import tkinter as tk
import pandas as pd
from tkintertable import TableCanvas

dados = pd.read_excel('testeExcel.xlsx')
dados_dict = dados.to_dict()

janela = tk.Tk()
janela.title("Astro - Controle de Demanda")
janela.geometry("1200x600")

barra_menu = tk.Menu(janela)
janela.config(menu=barra_menu)
barra_menu.config(bd=1, relief="solid")

arquivo_menu = tk.Menu(barra_menu)
barra_menu.add_cascade(label="Arquivo ", menu=arquivo_menu)
sobre_menu = tk.Menu(barra_menu)
barra_menu.add_cascade(label="Sobre ", menu=sobre_menu)

arquivo_menu.add_command(label="Abrir")
arquivo_menu.add_command(label="Salvar")

label = Label(janela, text="Controle de Demanda", font=("Verdana", 16), pady=40)
label.pack()

caixa = tk.Frame(janela, bd=2, relief="sunken")
caixa.pack(padx=5, pady=5)

tabela = TableCanvas(caixa, data=dados_dict)
tabela.config(width=1000, height=500)
tabela.show()

janela.update_idletasks()
largura = janela.winfo_width()
altura = janela.winfo_height()
x = (janela.winfo_screenwidth() // 2) - (largura // 2)
y = (janela.winfo_screenheight() // 2) - (altura // 2)
janela.geometry("{}x{}+{}+{}".format(largura, altura, x, y))


janela.mainloop()


