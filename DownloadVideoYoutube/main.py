import PySimpleGUI as sg
from pytube import YouTube

layout = [[sg.Text("Link do video "), sg.InputText()],
          [sg.Text('Informe caminho para Download'), sg.InputText(), sg.FolderBrowse()],
         [sg.Button('Baixar'), sg.Button('cancelar')]
         ]
janela = sg.Window("Donload de videos do Youtube",layout)
while True:
  event, values = janela.read()
  if event == 'cancelar' or event == sg.WIN_CLOSED:
    break
  elif event == 'baixar':
    video = YouTube(values[0])
    video.streams.get_highest_resolution().download(output_path=values[1])
    sg.popup_ok("Download Concluido com Exito")
janela.close()
