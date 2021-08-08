from tkinter import *
from tkinter import filedialog
import CalcHoras as ch
import writing_test as wt

root = Tk()

outFrameOne = Frame(root, height = "50px", width = "200px")
outFrameOne.pack(anchor = W)

frameOne = Frame(outFrameOne, height = "50px", width = "200px")
frameOne.pack( side = LEFT)

labelOne = Label(frameOne, text = "Ficheiro Horas")
labelOne.pack(anchor = W, side = TOP)

entryText = StringVar();

caminho = Entry(frameOne, width = 45, textvariable = entryText, state = DISABLED )
caminho.pack(side = LEFT)

def browseFiles():
    filename = filedialog.askopenfilename(initialdir = "C:/Users/Asus/Desktop",
                                                                        title = "Selecione o ficheiro das picagens",
                                                                        filetypes = (("Text files","*.txt*"),("all files","*.*")))
    caminho.delete(0, 'end')
    caminho.config(state = NORMAL)
    caminho.insert(0, filename)
    print(filename)

procurarBotao = Button(frameOne, text = "Procurar", command = browseFiles)
procurarBotao.pack(side = RIGHT)

outFrameTwo = Frame(root, height = "50px", width = "200px")
outFrameTwo.pack(anchor = W)

frameTwo = Frame(outFrameTwo, height = "50px", width = "200px", bd = 5)
frameTwo.pack(side = LEFT)

verificarBotao = Button(frameTwo, text = "Verificar Ficheiro", command = lambda: ch.verifyFile(ch.openFile(caminho.get())))
verificarBotao.pack(side = LEFT)
#verificarBotao.config(command = checkFile)

calcularBotao = Button(frameTwo, text = "Calcular Horas", command = lambda: wt.main(ch.openFile(caminho.get())))
calcularBotao.pack(side = RIGHT)

outFrameThree = Frame(root, height = "100px", width = "200px")
outFrameThree.pack(anchor = W)

frameThree = Frame(outFrameThree, height = "100px", width = "200px")
frameThree.pack(side = LEFT)

textcalc = StringVar()
messageBox = Message(frameThree, textvar = textcalc, relief = SUNKEN, width = 100)
messageBox.pack()

root.title("Calculadora Horários")
root.mainloop()
