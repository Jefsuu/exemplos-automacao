
#importando as bibliotecas necessárias
#o tkinter é nativo do python
from tkinter import Frame, filedialog, messagebox
#mailmerge e pandas é necessário instalar
from mailmerge import MailMerge
from pandas import read_excel
from tkinter.ttk import Label, Entry, Button
#ttkthemes também é necessário instalar
from ttkthemes import ThemedTk
import os

#instanciando o objeto/janela do tkinter
ws = ThemedTk(theme="arc")
#defininado o tamanho
ws.geometry('400x300')
#criando um frame
frame_inputs = Frame(ws, width=380)
frame_inputs.grid(row=0, sticky="ew")
#definindo variaveis globais que são utilizadas nas funções
path_mp = None
path_ripd = None
path_folder = None

#label de texto
Label(frame_inputs, text='CNPJ do controlador:', anchor='w').grid(row=1, column=0, sticky='w', padx=5)
#label para inserção de informações
label_cpnj = Entry(frame_inputs, width=32)
#definindo o posicionamento do objeto anterior
label_cpnj.grid(row=1, column=1, sticky='ew')
Label(frame_inputs, text='Razão Social do controlador:').grid(row=2, column=0, sticky='w', padx=5)
label_rz_social = Entry(frame_inputs)
label_rz_social.grid(row=2, column=1, sticky='ew')
Label(frame_inputs, text='Nome do DPO:').grid(row=3, column=0, sticky='w', padx=5)
label_nome_encarregado = Entry(frame_inputs)
label_nome_encarregado.grid(row=3, column=1, sticky='ew')
Label(frame_inputs, text='E-mail do DPO:').grid(row=4, column=0, sticky='w', padx=5)
label_email_encarregado = Entry(frame_inputs)
label_email_encarregado.grid(row=4, column=1, sticky='ew')
Label(frame_inputs, text='Responsável Mapa de Adequação:').grid(row=5, column=0, sticky='w', padx=5)
label_responsavel = Entry(frame_inputs)
label_responsavel.grid(row=5, column=1, sticky='ew')

#criando as funções para pegar o caminho dos arquivos/pastas
def get_mp():
    global path_mp
    path_mp = filedialog.askopenfilename()
    return path_mp

def get_ripd():
    global path_ripd
    path_ripd = filedialog.askopenfilename()
    return path_ripd

def get_folder():
    global path_folder
    path_folder = filedialog.askdirectory()
    return path_folder

#definindo a função que pega informações de uma planilha e coloca essas informações em um arquivo do word que 
#possui um mapa de onde cada informação deve ser inserida
def fill_ripd(path_mp, path_ripd, path_folder):
    #abrindo a planilha desejada
    df = read_excel(path_mp, sheet_name='Risk Management', skiprows=5)
    #filtrando os valores de interesse
    df = df[df['RIPD'] == 'SIM']
    #como na planilha tem muitas celulas mescladas e o pandas não faz esse reconhecimento, foi usado o metodo
    #ffill para preencher as celulas vazias de acordo com o ultimo valor e encontrado antes das celulas vazias
    df['Processo'].ffill(inplace=True)
    #agrupando pela coluna desejada
    d = df.groupby('Processo')
    #obtendo os grupos criados
    grupos = [i for i in d.groups]
    #criando lista para registrar o nome dos arquivos criados
    arquivos_criados = []
    #criando um arquivo para cada grupo criado
    for n in grupos:
        #abrindo o documento mapeado com as tags
        document = MailMerge(path_ripd)
        #criando o nome do arquivo
        nome_atividade = n.split('-')[-1].strip()

        #extraindo informações da planilha para preencher o arquivo word#################
        area_responsavel = d.get_group(n)['Setor'].values[0]
        medidas_seguranca = '\n**'.join(d.get_group(n)['Medidas para mitigação'].dropna())
        riscos = d.get_group(n)[['Descrição do Risco', 'Unnamed: 8', 'Unnamed: 10', 'Medidas para mitigação']]
        riscos['id'] = [f'R{str(i).zfill(2)}' for i in range(1, len(riscos) + 1) ]
        riscos['id_2'] = riscos['id']
        riscos['produto'] = riscos['Unnamed: 8'] * riscos['Unnamed: 10']
        riscos = riscos.rename(columns={'Descrição do Risco':'risco','Unnamed: 8':'probabilidade',
    'Unnamed: 10':'impacto', 'Medidas para mitigação':'medidas_tab'}).astype('str').to_dict(orient='records')
        ####################################################################################

        #inserindo as informações no local mapeado no word
        #deve ser passado o valor da tag da mesma forma que se encontra no word e o valor que irá receber
        document.merge(cnpj_controlador=label_cpnj.get())
        document.merge(rz_social_controlador=label_rz_social.get())
        document.merge(nome_encarregado=label_nome_encarregado.get())
        document.merge(email_encarregado=label_email_encarregado.get())
        document.merge(atividade_de_tratamento=nome_atividade)
        document.merge(area_responsavel=area_responsavel)
        document.merge(responsavel_preenchimento=label_responsavel.get())
        document.merge(medidas=medidas_seguranca)
        document.merge_rows('id', riscos)
        document.merge_rows('id_2', riscos)

        #salvando o documento
        document.write(f'{path_folder}/RIPD - {nome_atividade}.docx')
        #adicionando o nome do documento na lista
        arquivos_criados.append(f'RIPD - {nome_atividade}.docx')

    #mensagem para sinalizar que os arquivos foram criados
    messagebox.showinfo(title='Sucesso', message=f'Os arquivos: {[i for i in arquivos_criados]} foram criados com sucesso', parent=ws)

#criando frame para os botões
frame_botoes = Frame(ws)
#posicionando frame
frame_botoes.grid(row=3, column=0, sticky='ns', pady=25)
#criando os botões que executa as funções que obtem o caminho dos arquivos/pastas
Button(frame_botoes, text='Selecione o mapa de adequação', command=get_mp).grid(row=1, sticky='ew')
Button(frame_botoes, text='Selecione o template de RIPD', command=get_ripd).grid(row=2, sticky='ew')
Button(frame_botoes, text='Selecione a pasta para salvar os arquivos', command=get_folder).grid(row=3, sticky='ew')

#butão que executa a função que preenche o word
Button(frame_botoes, text='Executar', command= lambda: fill_ripd(path_mp=path_mp, path_ripd=path_ripd, path_folder=path_folder)).grid(row=4, sticky='ew')

#executando a janela e os objetos instanciados
ws.mainloop()

