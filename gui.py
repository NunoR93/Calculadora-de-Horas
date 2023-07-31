import tkinter
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import Turnos as turnos

window = tkinter.Tk()
window.resizable(False,False)

file_path = StringVar()
file_savePath = StringVar()

conteudo = tkinter.Frame(window)
frame = tkinter.Frame(conteudo)
file_label = tkinter.Label(conteudo,text="Ficheiro")
file_saveLabel = tkinter.Label(conteudo,text="Destino")
file_entry = tkinter.Entry(conteudo, textvariable=file_path, state=DISABLED, bd=3)
file_save = tkinter.Entry(conteudo,textvariable=file_savePath,state=DISABLED,bd=3)

def select_file():
    file_Types = (('text files', '*.txt'),('All Files','*.*'))
    file_path  = filedialog.askopenfilename(
        title='Selecione o ficheiro',
        filetypes=file_Types
    )
    file_entry.delete(0,'end')
    file_entry.config(state=NORMAL)
    file_entry.insert(0,file_path)

def select_destination():
    file_savePath = filedialog.askdirectory(title="Choose directory")
    file_save.delete(0,'end')
    file_save.config(state=NORMAL)
    file_save.insert(0,file_savePath)

def execute_functions():
    if(len(file_savePath.get()) < 3):
        tkinter.messagebox.showinfo("Error","Escolha pasta de destino")
    else:
        turnos.createExcel((file_entry.get()),file_savePath.get())
        tkinter.messagebox.showinfo("Done","Ficheiro Criado")

button_search = tkinter.Button(conteudo,text="Procurar",command=select_file)
button_ok = tkinter.Button(conteudo,text="Calcular", command=lambda :execute_functions())
button_save = tkinter.Button(conteudo,text="Procurar",command=lambda :select_destination())

conteudo.grid(column=0,row=0)
frame.grid(column=0,row=0,columnspan=4,rowspan=3)
file_label.grid(column=0,row=0)
file_entry.grid(column=1,row=0,columnspan=2)
button_search.grid(column=3,row=0)
button_ok.grid(column=2,row=2)
button_save.grid(column=3,row=1)
file_save.grid(column=1,columnspan=2,row=1)
file_saveLabel.grid(column=0,row=1)
window.title("Calculadora Horas")
window.mainloop()