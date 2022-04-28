from tkinter import *
from tkinter.ttk import Combobox
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
from datetime import datetime
import pandas as pd
import numpy as np
import requests

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()
lista_coin = list(dicionario_moedas.keys())


def pegar_cotacao():
    moeda = cbox1.get()
    data_cotacao = calendario1.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}_date={ano}{mes}{dia}"
    requisicao_moedas = requests.get(link)
    cotacao = requisicao_moedas.json()
    valor_moeda = cotacao[0]['bid']
    msg3['text'] = f"A cotação da {moeda} no dia {data_cotacao} foi de R${valor_moeda}"


def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title="Selecione o Arquivo de Moeda")
    var_caminhoarquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivoselecionado['text'] = f"Arquivo Selecionado: {caminho_arquivo}"


# noinspection PyArgumentList
def atualizar_cotacoes():
    try:
        df = pd.read_excel(var_caminhoarquivo.get())
        moedas = df.iloc[:, 0]
        data_inicial = calendario3.get()
        data_final = calendario3.get()

        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]

        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        for moeda in moedas:
            link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?" \
                   f"start_date={ano_inicial}{mes_inicial}{dia_inicial}_date={ano_final}{mes_final}{dia_final}"

            requisicao_moedas = requests.get(link)
            cotacoes = requisicao_moedas.json()
            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')
                if data not in df:
                    print(data)
                    df[data] = np.nan

                df.loc[df.iloc[:, 0] == moeda, data] = bid
            df.to_excel("Moedas.xlsx")
            label_atualizarcotacoes['text'] = "Arquivo Atualizado com Sucesso"
    except:
        label_atualizarcotacoes['text'] = "Selecione um Arquivo Excel no Formato Correto "


jan = Tk()
jan.title('Sistema de Cotação de Moedas')

# Titulo
titulo1 = Label(text="Cotação de 1 Moeda Especica", borderwidth=2, relief='solid')
titulo1.grid(row=0, column=0, padx=10, pady=10, sticky='NSWE', columnspan=3)

# Mensagem de seleção e combobox.
msg1 = Label(text="Selecionar Moeda", anchor='e')
msg1.grid(row=1, column=0, padx=10, pady=10, sticky='NSWE', columnspan=2)

cbox1 = Combobox(values=lista_coin)
cbox1.grid(row=1, column=2, padx=10, pady=10, sticky='NSWE')

# Mensagem de selecionar o dia e calendario.
msg2 = Label(text="Selecionar o dia que deseja pegar a cotação", anchor='e')
msg2.grid(row=2, column=0, padx=10, pady=10, sticky='NSWE', columnspan=2)

calendario1 = DateEntry(year=2021, locale='pt_br')
calendario1.grid(row=2, column=2, padx=10, pady=10, sticky='NSWE')

# Mensagem com o resultado da cotação com botão de buscar.
msg3 = Label(text=f"")
msg3.grid(row=3, column=0, padx=10, pady=10, sticky='NSWE', columnspan=2)

botao1 = Button(text='Pegar Cotação', command=pegar_cotacao)
botao1.grid(row=3, column=2, padx=10, pady=10, sticky='NSWE')

# Cotação de várias moedas.
titulo2 = Label(text="Cotação de Múltiplas Moedas", borderwidth=2, relief='solid')
titulo2.grid(row=4, column=0, padx=10, pady=10, sticky='NSWE', columnspan=3)

msg4 = Label(text="Selecione um arquivo em Excel com Moedas na Coluna A")
msg4.grid(row=5, column=0, padx=10, pady=10, sticky='NSWE', columnspan=2)

var_caminhoarquivo = StringVar()

botao2 = Button(text='Clique para Selecionar', command=selecionar_arquivo)
botao2.grid(row=5, column=2, padx=10, pady=10, sticky='NSWE')

label_arquivoselecionado = Label(text="Nenhum Arquivo Selecionado", anchor='e')
label_arquivoselecionado.grid(row=6, column=0, padx=10, pady=10, sticky='NSWE', columnspan=3)

# Data inicial e data Final com entry do calendario.
msg6 = Label(text="Data Inicial", anchor='e')
msg6.grid(row=7, column=0, padx=10, pady=10, sticky='NSWE')

calendario2 = DateEntry(year=2021, locale='pt_br')
calendario2.grid(row=7, column=1, padx=10, pady=10, sticky='NSWE')

msg7 = Label(text="Data Final", anchor='e')
msg7.grid(row=8, column=0, padx=10, pady=10, sticky='NSWE')

calendario3 = DateEntry(year=2021, locale='pt_br')
calendario3.grid(row=8, column=1, padx=10, pady=10, sticky='NSWE')

# Botão atualizar cotações.
botao3 = Button(text='Atualizar Cotações', command=atualizar_cotacoes)
botao3.grid(row=9, column=0, padx=10, pady=10, sticky='NSWE')

label_atualizarcotacoes = Label(text=f"")
label_atualizarcotacoes.grid(row=9, column=1, padx=10, pady=10, sticky='NSWE', columnspan=2)

# Botão de fechar.
botao4 = Button(text="Fechar", command=jan.quit)
botao4.grid(row=10, column=2, padx=10, pady=10, sticky='nswe')

jan.mainloop()
