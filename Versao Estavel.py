import serial
import pandas as pd
import tkinter as tk
import sys
from tkinter import ttk

# import tksheet
#
# configura a porta serial
serialPort = serial.Serial(port="COM9", baudrate=9600, bytesize=8, timeout=2, stopbits=serial.STOPBITS_ONE)
serialString = ""  # Used to hold data coming over UART
# define parâmetros
data = 0
Spectrum_list = []
#
# Classe da Interface com Usuário
class Tela:
    def __init__(self, master):
        #

        self.master = master
        self.title = "SmartSpec3000"
        #
        # define tamanho da janela
        #
        self.master.geometry('350x100')
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
        self.botaoSair.grid(column=2, row=6)
        self.botaoLeitura = tk.Button(master, text='  Obter Dados   ', command=self.leitura)
        self.botaoLeitura.grid(column=1, row=6)
        #
        self.master.mainloop()
        #
    def popup_showinfo(self):
        #
        #  Tratamento de Erro:
        win = tk.Toplevel()
        showinfo("ATENÇÃO", "Definir o nome do arquivo para armazenar o espectro!!")
        #
    def leitura(self):
        nome_arquivo=self.nome.get()
        #
        # avalia erros associados ao nome do arquivo
        if (nome_arquivo==''):
            print("sem nome do arquivo")
            self.button_showinfo = ttk.Button(self, text="Show Info", command=popup_showinfo)
            self.button_showinfo.pack()
        print (nome_arquivo)
        LeituraRS(nome_arquivo)

class LeituraRS():
    def __init__(self,nome_arquivo):

        self.leitura(nome_arquivo)

    def leitura(self,nome):
        print("inicio da coleta")
        Spectrum_list = []
        linha = 0
        print("nome do arquivo:  ",nome)
        while (linha < 1000):
            # Wait until there is data waiting in the serial buffer
            if (serialPort.in_waiting > 0):

                # Read data out of the buffer until a carraige return / new line is found
                serialString = serialPort.readline()

                # Print the contents of the serial data
                print(serialString.decode('Ascii'))
                # recebe a string da linha e Save line in list Spectrum_list
                Asc_Spectrum = serialString.decode('Ascii')
                Spectrum_list.append(Asc_Spectrum)
                # salva em um dataframe do Pandas
                df_Spectrum = pd.DataFrame(Spectrum_list)  # , columns=['lambda','absorbance'])
                #
                # identifica fim dos dados numéricos
                if ("start" in Asc_Spectrum):
                    # split dados coletados (delimiter= ':') em duas colunas
                    df3 = df_Spectrum[0].str.split(pat=": ", n=1, expand=True)
                    # Converter para numerico - renomear as colunas
                    df3.rename(columns={0: 'Lambda',1:'Absorbância', 2:'Excesso'}, inplace=True)
                    df3['Lambda'] = pd.to_numeric(df3['Lambda'], errors='coerce')
                    df3['Absorbância'] = pd.to_numeric(df3['Absorbância'], errors='coerce')
                    #df3.rename(columns={0: 'Lambda',1:'Absorbância', 2:'Excesso'}, inplace=True)
                    # salvar com arquivo Excel
                    df3.to_excel(f"{nome}.xlsx")
                    print("arquivo criado")
                    print(df3)
                    sys.exit()
                # Tell the device connected over the serial port that we recevied the data!
                # The b at the beginning is used to indicate bytes!
                serialPort.write(b"Thank you for sending data \r\n")
                #
                # Identifica o EOF
                if (Asc_Spectrum==''):
                    linha = 1001
                # Apagar após teste
                if linha == 609:
                    # linha = 0
                    print(df_Spectrum)
                    linha = 1001
                # conta o número de linhas importadas da RS232
                linha = linha + 1

    def LeituraSerial(self):
        LeituraRS.leitura(self)
#------------------------------------------------------------------
if __name__ == '__main__':
    window = tk.Tk()
    app = Tela(window)

