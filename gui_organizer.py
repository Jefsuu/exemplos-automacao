from tkinter import CENTER, Button, Frame, Label, Tk, filedialog
from tkinter.ttk import Style
import os, shutil

janela = Tk()
janela.geometry('200x100')

folder = None
def get_p():
    global folder
    folder = filedialog.askdirectory()
    return folder

dest = None
def get_d():
    global dest
    dest = filedialog.askdirectory()
    return dest

butonf = Button(master=janela,
text='Selecione a pasta dos arquivos',
fg='red', 
command=get_p,
bg='white')
butonf.place(relx=0.5, rely=0.5, anchor=CENTER)
butonf.pack()

butond = Button(master=janela,
text='Selecione a pasta de destino',
fg='red', 
command=get_d,
bg='white')
butond.place(relx=0.5, anchor=CENTER)
butond.pack()

janela.mainloop()
for i in os.listdir(folder):
    if len(i.split('-'))> 2:
        if os.path.exists(dest+'/'+i.split('-')[1].strip()) == False:
            os.mkdir(dest+'/'+i.split('-')[1].strip())
 

err_file= []
for i in os.listdir(folder):
    for x in os.listdir(dest):
        if len(i.split('-'))> 2:
            if i.split('-')[1].strip() == x:
                shutil.copy(folder+'/'+i,dest+'/'+x)
        else:
            err_file.append(i)

os.mkdir(dest+'/'+"A VERIFICAR")
for i in list(set(err_file)):
    shutil.copy(folder+'/'+i, dest+'/'+"A VERIFICAR")
p = dest+'/'+"A VERIFICAR"

print(f"{len(os.listdir(p))} - Arquivos precisam ser verificados")


