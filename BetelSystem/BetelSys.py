import PySimpleGUI as sg
import Function
from pyrsistent import v

# variables
# transform the .txt file in a dicionary
cod = {}
with open('Pigmento.txt') as f:
    cod = dict(x.rstrip().split(None, 1) for x in f)  # type: ignore
# List of itens to show in combo function
itens = [
    '700g',
    '83g',
    '650g',
    '15,7g',
    'Tampa 5L',
    'Tampa 20L',
    'Alça',
    '',
]
# List of sopro's itens to show in combo function
sopro_itens = [
    '20L', '5L', '10L'
]
# .txt file with contain the list of emails to send the file
email_list = open('Email.txt', 'r').readlines()
# GUITheme
sg.theme('PythonPlus')
# GUI of injeção
injecao = [
    [sg.Text('*Máquina:'), sg.Input('', size=(10, 30), key='-Maquina-'), sg.Text('*Produto: '),
     sg.Combo(itens, size=(15), key='-Itens-')],
    [sg.Text('*Quant.  :'), sg.Input('', size=(20), key='-Quantidade-')],
    [sg.Text('*Cliente:  '), sg.Input('', size=20, key='-Cliente-')],
    [sg.Text('*Cor:       '), sg.Input('', size=(20, 50), key='-Cor-')],
    [sg.Text('COD:      '), sg.Combo(list(cod), size=(20, 50), key='-COD-')],
    [sg.Text('OBS:      '), sg.Input('', size=(30, 2), key='-OBS-')],

    [sg.Text('Embalagem:')], [sg.Radio(group_id='1', text='Caixa', default=True, key='-Emb_caixa-'),
                              sg.Radio(group_id='1', text='Bag', key='-Emb_bag-'),
                              sg.Radio(group_id='1', text='Saco', key='-Emb_saco-')],
    [sg.Text('Ações:')], [sg.Checkbox('Imprimir', default=True, key='-Imprimir-'),
                          sg.Checkbox('Enviar Email', default=True, key='-Send_email-')],
    [sg.Button('Enviar', button_color=(sg.YELLOWS[0], sg.GREENS[0]), expand_x=True),
     sg.Button('Sair', button_color=(sg.YELLOWS[0], 'Red'), expand_x=True)]

]
# GUI of sopro
sopro = [
    [sg.Text('*Máquina:    '), sg.Input('', size=(10, 30), key='-sMaquina-')],
    [sg.Text('*Produto:     '), sg.Combo(sopro_itens, size=(20), key='-sItens-')],
    [sg.Text('*Quantidade:'), sg.Input('', size=(20), key='-sQuantidade-')],
    [sg.Text('*Cliente:      '), sg.Input('', size=20, key='-sCliente-')],
    [sg.Text('*Cor:           '), sg.Input('', size=(20, 50), key='-sCor-')],
    [sg.Text(' OBS:         '), sg.Input('', size=(30, 2), key='-sOBS-')],
    [sg.Text('Ações:')], [sg.Checkbox('Imprimir', default=True, key='-sImprimir-'),
                          sg.Checkbox('Enviar Email', default=True, key='-sSend_email-')],
    [sg.Button('Enviar', button_color=(sg.YELLOWS[0], sg.GREENS[0]), expand_x=True, k='-Enviar-'),
     sg.Button('Sair', button_color=(sg.YELLOWS[0], 'Red'), expand_x=True, k='-Sair-')]
]
# GUI log
log = [
    sg.Output(size=(40, 20))
]
# Final layot
layout = [
    [sg.TabGroup(
        [[
            sg.Tab('Injeção', injecao),
            sg.Tab('Sopro', sopro)]]),
        sg.Output(size=(30, 17), key='-Output-')
    ],

    [sg.Text('Criado por Levi Renato', font='italic')],

]

# var to creat the window
play = sg.Window('BetelSystem', layout)


# Event Loop to process "events" and get the "values" of the inputs
def start_program():
    while True:
        event, values = play.read()  # type: ignore
        if event == sg.WIN_CLOSED or event == 'Sair' or event == '-Sair-':  # if user closes window or clicks 'Sair'
            break
        # save values in variables
        numero_maquina = values['-Maquina-']
        quantidade = values['-Quantidade-']
        cliente = values['-Cliente-']
        produto = values['-Itens-']
        cor = values['-Cor-']
        Cod = values['-COD-']
        obs = values['-OBS-']
        # checks what´s is packing choice
        vol = ''
        if values['-Emb_caixa-']:
            vol = 'Cx'
        elif values['-Emb_bag-']:
            vol = 'Bag'
        elif values['-Emb_saco-']:
            vol = 'Saco'
        else:
            print('Error')

        # After the event click 'Enviar', verify if the current fields is empty
        if event == 'Enviar':
            if not numero_maquina and numero_maquina == '':
                sg.popup_error('Não deixe os campos com * vazio!')
                start_program()
            elif not produto and produto == '':
                sg.popup_error('Não deixe os campos com * vazio!')
                start_program()
            elif not cor and cor == '':
                sg.popup_error('Não deixe os campos com * vazio!')
                start_program()
            # quantidade
            elif not quantidade and quantidade == '':
                sg.popup_error('Não deixe os campos com * vazio!')
                start_program()
            # cliente
            elif not cliente and cliente == '':
                sg.popup_error('Não deixe os campos com * vazio!')
                start_program()
            else:
                # after cheks the filds, put the values in current function to write in sheet
                Function.injecao_title(str(numero_maquina))
                Function.injecao_prod(str(produto))
                Function.injecao_cor(str(cor))
                Function.injecao_cli(str(cliente))
                Function.injecao_COD(cod.get(Cod))
                try:
                    Function.injecao_qnt_total(int(quantidade), produto, vol)
                except:
                    sg.popup_error('Digite uma quantidade válida!')
                    start_program()

                Function.injecao_data()
                Function.injecao_OBs(str(obs))
                Function.start()
                if values['-Imprimir-']:
                    Function.imprimir()
                    print('Arquivo enviado p/ impressão')
                else:
                    pass
                if values['-Send_email-']:
                    for i in email_list:
                        if i == '':
                            break
                        try:
                            Function.send_email(i)
                            print(f'Email enviado p/ {i}')
                        except:
                            print('Falha ao enviar email')
                else:
                    pass
                print('==========================')
                sg.popup('Programação Feita!')

        ##########        Sopro- page 2       #################################
        # save values in variables
        snumero_maquina = values['-sMaquina-']
        squantidade = values['-sQuantidade-']
        scliente = values['-sCliente-']
        sproduto = values['-sItens-']
        scor = values['-sCor-']
        sobs = values['-sOBS-']
        # After the event click 'Enviar', verify if the current fields is empty
        if event == '-Enviar-':
            if not snumero_maquina and snumero_maquina == '':
                sg.popup_error('Não deixe os campos com * vazio!')
                start_program()
            elif not sproduto and sproduto == '':
                sg.popup_error('Não deixe os campos com * vazio!')
                start_program()
            elif not scor and scor == '':
                sg.popup_error('Não deixe os campos com * vazio!')
                start_program()
            # quantidade
            elif not squantidade and squantidade == '':
                sg.popup_error('Não deixe os campos com * vazio!')
                start_program()
            # cliente
            elif not scliente and scliente == '':
                sg.popup_error('Não deixe os campos com * vazio!')
                start_program()
            else:
                # after cheks the filds, put the values in current function to write in sheet
                Function.sopro_title(str(snumero_maquina))
                Function.sopro_prod(str(sproduto))
                Function.sopro_cor(str(scor))
                Function.sopro_cli(str(scliente))

                try:
                    Function.sopro_qnt_total(int(squantidade))
                except:
                    sg.popup_error('Digite uma quantidade válida!')
                    start_program()

                Function.sopro_data()
                Function.sopro_OBs(str(sobs))
                Function.sstart()
                if values['-sImprimir-']:
                    Function.simprimir()
                    print('Arquivo enviado p/ impressão')
                else:
                    pass
                if values['-sSend_email-']:
                    for i in email_list:
                        try:
                            Function.send_email(i)
                            print(f'Email enviado p/ {i}')
                        except:
                            print('Falha ao enviar email')
                else:
                    pass
                print('==========================')
                sg.popup('Programação Feita!')

    play.close()


start_program()
