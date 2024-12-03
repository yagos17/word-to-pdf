from os import path, listdir
from docx2pdf import convert
from tkinter import filedialog
import PySimpleGUI as sg

layout = [
    [sg.Text('Converta Word p/ PDF')],
    [sg.Button('Converter')]
]

janela = sg.Window('Conversor', layout)

while True:
    event, values = janela.read()
    if event == sg.WIN_CLOSED:
        break
    elif event == 'Converter':
        try:
            pasta = filedialog.askdirectory()
            if not pasta:
                sg.popup('Nenhuma pasta selecionada.')
                continue

            caminho_pasta = path.join(pasta)
            lista_arquivos = listdir(caminho_pasta)

            print(lista_arquivos)
            print(caminho_pasta)

            if lista_arquivos:
                for arquivo in lista_arquivos:
                    if arquivo.endswith('.docx'):
                        convert(path.join(caminho_pasta, arquivo))
                sg.popup('Conversão Concluída.')
            else:
                sg.popup('Nenhum arquivo .docx encontrado na pasta.')

        except FileNotFoundError:
            sg.popup('Erro: Pasta não encontrada.')
        except Exception as e:
            sg.popup(f'Erro: {e}')

janela.close()