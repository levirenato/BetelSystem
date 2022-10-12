# Import libraries
import os
import time

#get anexo, to the Function script. Its the file just created
def imprimir(anexo):
# Insert the directory path in here
    path = 'Historico'
    file = anexo
# Send to impress
    file_path = f'{path}\\{file}'
    os.startfile(file_path, 'print')
    print(f'Printing {file_path}')
