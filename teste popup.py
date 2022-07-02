from pathlib import Path

import pandas as pd
import tkinter as tk
import sys
from tkinter import ttk
from tkinter.messagebox import showinfo
from tkinter import filedialog 
import os
#
# define parâmetros
data = 0
Spectrum_list = []
#
# Classe da Interface com Usuário
class Tela(ttk.Frame):
    def __init__(self, master):
        #
        #ttk.Frame.__init__(self, master) 
        self.master = master
        self.title = "SmartSpec3000"
        #
        # define tamanho da janela
        #
        self.master.geometry('400x100')
        #
        # Botões e Texto no MasterFrame
        #
        self.caixa1 = tk.Frame(master, borderwidth=2, relief='raised')
        self.caixa1.grid(column=1, row=0)
        self.texto1 = tk.Label(self.caixa1, text='---------SmartSpec3000-----------')
        self.texto1.grid(column=2, row=1)
        self.texto2 = tk.Label(master, text='Nome do arquivo Excel') # a ser criado - sem extensão')
        self.texto2.grid(column=1, row=3)
        self.nome=tk.Entry(master)
        self.nome.grid(column=2,row=3)
        self.botaoSair = tk.Button(master, text='     Sair    ', command=master.destroy)
        self.botaoSair.grid(column=3, row=6)
        self.botaoLeitura = tk.Button(master, text='  Obter Dados   ', command=self.leitura)
        self.botaoLeitura.grid(column=1, row=6)
        self.botaoDiretorio = tk.Button(master, text=' Selecionar Diretório  ', command=self.diretorio)
        self.botaoDiretorio.grid(column=2, row=6)
        #
        self.master.mainloop()
        #
    def diretorio(self):
        # Allow user to select a directory and store it in global var
        # called folder_path
        global folder_path
        filename = filedialog.askdirectory()
        folder_path.set(filename)
        
        # definir folder_path como diretorio ativo
        sourcePath = folder_path.get()
        os.chdir(sourcePath)  # Provide the path here
        print(filename)

    def popup_showinfo(self):
        #
        #  Tratamento de Erro:        
        print('chegou em popup_showinfo')
        showinfo("ATENÇÃO", "Definir o nome do arquivo para armazenar o espectro!!")
        #
          
    def leitura(self):
        nome_arquivo=self.nome.get()
        print(Path.cwd())
        if (Path(nome_arquivo).exists() and nome_arquivo!=''):
           showinfo("ATENÇÃO", "Esse nome já existe!!")
           print("EPA")
        #
        # avalia erros associados ao nome do arquivo
        if (nome_arquivo==''):
            print("sem nome do arquivo")
            showinfo("ATENÇÃO", "Definir o nome do arquivo para armazenar o espectro!!")
            #self.popup_showinfo()
        print (nome_arquivo)

#---------------------------------
if __name__ == '__main__':
    window = tk.Tk()
    folder_path = tk.StringVar()
    app = Tela(window)