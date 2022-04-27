
from tkinter import Frame, filedialog, messagebox

from mailmerge import MailMerge

from pandas import read_excel
from tkinter.ttk import Label, Entry, Button
from ttkthemes import ThemedTk
import os

"""ws = ThemedTk(theme="arc")
ws.geometry('400x400')
frame_inputs = Frame(ws)
path_mp = None
path_ripd = None
path_folder = None
Label(text='Insira o CNPJ do controlador').pack()
label_cpnj = Entry(ws)
label_cpnj.pack()
Label(text='Insira a Razão Social do controlador').pack()
label_rz_social = Entry(ws)
label_rz_social.pack()
Label(text='Insira o Nome do DPO').pack()
label_nome_encarregado = Entry(ws)
label_nome_encarregado.pack()
Label(text='Insira o E-mail do controlador').pack()
label_email_encarregado = Entry(ws)
label_email_encarregado.pack()
Label(text='Insira o nome de quem preencheu o Mapa de Adequação').pack()
label_responsavel = Entry(ws)
label_responsavel.pack()"""

ws = ThemedTk(theme="arc")
ws.geometry('400x300')
frame_inputs = Frame(ws, width=380)
frame_inputs.grid(row=0, sticky="ew")
path_mp = None
path_ripd = None
path_folder = None
Label(frame_inputs, text='CNPJ do controlador:', anchor='w').grid(row=1, column=0, sticky='w', padx=5)
label_cpnj = Entry(frame_inputs, width=32)
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



def fill_ripd(path_mp, path_ripd, path_folder, cnpj, rz_social, nome, email, responsavel):
    df = read_excel(path_mp, sheet_name='Risk Management', skiprows=5)
    df = df[df['RIPD'] == 'SIM']
    df['Processo'].ffill(inplace=True)
    d = df.groupby('Processo')
    grupos = [i for i in d.groups]
    arquivos_criados = []
    for n in grupos:
        document = MailMerge(path_ripd)

        nome_atividade = n.split('-')[-1].strip()
        area_responsavel = d.get_group(n)['Setor'].values[0]

        medidas_seguranca = '\n**'.join(d.get_group(n)['Medidas para mitigação'].dropna())
        riscos = d.get_group(n)[['Descrição do Risco', 'Unnamed: 8', 'Unnamed: 10', 'Medidas para mitigação']]
        riscos['id'] = [f'R{str(i).zfill(2)}' for i in range(1, len(riscos) + 1) ]
        riscos['id_2'] = riscos['id']
        riscos['produto'] = riscos['Unnamed: 8'] * riscos['Unnamed: 10']
        riscos = riscos.rename(columns={'Descrição do Risco':'risco','Unnamed: 8':'probabilidade',
    'Unnamed: 10':'impacto', 'Medidas para mitigação':'medidas_tab'}).astype('str').to_dict(orient='records')

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

        document.write(f'{path_folder}/RIPD - {nome_atividade}.docx')
        arquivos_criados.append(f'RIPD - {nome_atividade}.docx')


    messagebox.showinfo(title='Sucesso', message=f'Os arquivos: {[i for i in arquivos_criados]} foram criados com sucesso', parent=ws)

cnpj_controlador_l = label_cpnj.get()
rz_social_controlador_l = label_rz_social.get()
nome_encarregado_l = label_nome_encarregado.get()
email_encarregado_l = label_email_encarregado.get()
responsavel_preenchimento_l = label_responsavel.get()

frame_botoes = Frame(ws)
frame_botoes.grid(row=3, column=0, sticky='ns', pady=25)

Button(frame_botoes, text='Selecione o mapa de adequação', command=get_mp).grid(row=1, sticky='ew')
Button(frame_botoes, text='Selecione o template de RIPD', command=get_ripd).grid(row=2, sticky='ew')
Button(frame_botoes, text='Selecione a pasta para salvar os arquivos', command=get_folder).grid(row=3, sticky='ew')

Button(frame_botoes, text='Executar', command= lambda: fill_ripd(path_mp=path_mp, path_ripd=path_ripd, path_folder=path_folder, cnpj=cnpj_controlador_l,
rz_social=rz_social_controlador_l, nome=nome_encarregado_l, email=email_encarregado_l, responsavel=responsavel_preenchimento_l)).grid(row=4, sticky='ew')


ws.mainloop()

