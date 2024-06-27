import PyPDF2
import csv
from tkinter import *
from tkinter import filedialog
import os
import re


def converterParaCSV():
    # Open the PDF file in read-binary mode
    with open(arquivo, 'rb') as file:
        # Create a PDF file reader object
        pdf_reader = PyPDF2.PdfReader(file)
        nome_arquivo_csv()
        populaCSV(pdf_reader)

# Loop through each page in the PDF file


def populaCSV(arquivoPDF):
    cabecalhos = ('CONTA', 'VALOR ORIGINAL', 'VALOR ATUALIZADO',
                  'DEPREC. NO MES', 'DEPREC. NO EXERC', 'DEPREC. ACUMULADA')
    # dicionario = []
    # Create a CSV file writer object
    csv_writer = csv.writer(
        open(arquivo_CSV+'.csv', 'w', newline=''), dialect='excel', delimiter=';')
    output = []
    output.append(cabecalhos)

    for page_num in range(1, len(arquivoPDF.pages)):
        # Get the page object
        page = arquivoPDF.pages[page_num]
        # Extract the text from the page
        text = page.extract_text()
        dados = []
        linhas = text.split("\n")
        for linha in linhas:

            linha_conta = str(linha).startswith("Conta")
            linha_rs = "R$" in linha
            linha_valor = "Valor" in linha

            if (linha_conta or linha_rs) and not linha_valor:
                dados.append(linha)

        dados = dados[:-1]

        pairs = list(zip(dados[::2], dados[1::2]))

        for pair in pairs:
            conta = re.sub(r'\s+', ' ', str(pair[0]).split(" ", 1)[1].strip())
            valores = pair[1].split(" ")
            valor_original = float(
                valores[1].replace('.', '').replace(',', '.'))
            valor_atualizado = float(
                valores[2].replace('.', '').replace(',', '.'))
            valor_deprec_mes = float(
                valores[3].replace('.', '').replace(',', '.'))
            valor_deprec_exerc = float(
                valores[4].replace('.', '').replace(',', '.'))
            valor_deprec_acumul = float(
                valores[5].replace('.', '').replace(',', '.'))

            linha = tuple([conta, valor_original, valor_atualizado,
                          valor_deprec_mes, valor_deprec_exerc, valor_deprec_acumul])
            output.append(linha)

    csv_writer.writerows(output)


def abrir_arquivo():
    global arquivo
    caminho = filedialog.askopenfilename(initialdir=os.getcwd(
    ), title="identifique o arquivo de contatos",  filetypes=(("text files", "*.PDF"), ("all files", "*.*")))
    caminhoParaLista = caminho.split('/')
    arquivo = caminhoParaLista[-1]


def nome_arquivo_csv():
    global arquivo_CSV
    arquivo_CSV_temp = arquivo_saida.get('1.0', END)
    arquivo_CSV = arquivo_CSV_temp.strip()


janela = Tk()
janela.title("Disparo de mensagens - Whatsapp Busness")
janela.geometry("500x180")
janela.configure(background="#dde")
# ----------------------------------------------------
Button(janela, text="Sair", command=janela.quit).place(x=400, y=10)
Button(janela, text="Selecionar arquivo PDF ",
       command=abrir_arquivo).place(x=50, y=10)

Label(janela, text="Nome do arquivo de saída").place(x=50, y=50)
arquivo_saida = Text(janela, height=1, width=50, autoseparators=True)
arquivo_saida.place(x=50, y=80, )

# Label(janela, text="Link para postagem").place(x=50, y=210)
# link = Entry(janela)
# link.place(x=50, y=230, width=400)

# ------------------------------------------------------
btnEnviar = Button(janela, text="Converter Arquivo", command=converterParaCSV)
btnEnviar.place(x=180, y=120)

janela.mainloop()
