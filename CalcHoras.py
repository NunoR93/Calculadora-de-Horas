

def countLines(file):
    #print("countLines")
    lines = file.readlines()
    file.seek(0)
    return lines
    
def verifyFile(file):
    #print("VerifyFile")
    try:
        nLines = len(countLines(file))-1
        if((nLines % 2) != 0):
            print("Ficheiro Inválido")
            print(str(nLines) + " Picagens\n")
        else:
            print("Ficheiro Válido")
        file.seek(0)
    except:
        print("Ficheiro Inválido")

def openFile(directory):
    print("openFile")
    try:
        file = open(directory, 'r')
    except:
        print("Ficheiro não encontrado")
    return file
    
def tirarTurnos(file):
    #print("tirarTurnos")
    try:
        file.seek(0)
        line = file.readline()
        turnos = []
        while line:
            tmp = line.replace("\n","").split("\t")
            turnos += [tmp[0],tmp[1].split(" ")]
            line = file.readline()
        file.seek(0)
        #print(turnos)
    except:
        print("Ficheiro Invalido")

def strToMin(s):
    #print("strToMin")
    try:
        time = s.split(":")
        time.pop()
        time[0] = int(time[0])
        if(time[0] < 10):
            time[0] += 24
        time[1] = int(time[1])
        #print(time)
        return time
    except:
        print("Formato invalido")

def hoursToMin(h):
    #print("hoursToMin")
    tmpH = h[0]*60
    tmpM = h[1]
    return tmpH+tmpM

def minToHours(minutes):
    h = int(minutes/60)
    m = minutes % 60
    s = "%02d:%02d" % (h,m)
    return s

#n nome, d dia, h hora
def writeToFile(n,d,h):
    file = open("Horas.txt","aw")

