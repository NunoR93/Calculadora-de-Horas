import xlsxwriter
import CalcHoras as ch
workbook = xlsxwriter.Workbook("Horas.xlsx")
#workbook.close()


#f = open("test.txt", "r")
def getNames(f):
    #print("Get_Names")
    tmp = []
    line = f.readline()
    while line:
        line = f.readline()
        tline = line.split('\t')
        if(tline[0] not in tmp):
            tmp+=[tline[0]]
        
    f.seek(0)
    return tmp

#Recebe ficheiro e nome de empregado e devolve lista com nome de empregado e turnos
def getTurn(file, n):
    #print("getTurn")
    line = file.readline()
    t = []
    j = 0
    name = ""
    
    #Tirar e formatar conteudo da linha
    while line:
        tmp = line.replace("\n","").split("\t")
        t += [[tmp[0],tmp[1].split(" ")]]
        line = file.readline()

    turnos = []

    #Percorrer lista de turnos e tirar apenas os que coincidem com o nome fornecido
    for i in t:
        if(n in i[0]):
            turnos += [i[1]]
            name = i[0]
            if(j == 0):
                print(i[0])
                worksheet = workbook.add_worksheet(i[0])
                worksheet.write(0,0,"Dia")
                worksheet.write(0,1,"Hora")
                j+=1

    #Transformar as horas que estão em formato String para uma lista [horas, minutos]
    for i in turnos:
        worksheet.write(j,0,i[0])
        worksheet.write(j,1,i[1])
        i[1] = ch.hoursToMin(ch.strToMin(i[1]))
        j+=1
    turnos += [name]
    #workbook.close()
    file.seek(0)
    return turnos

def calcH(turn):
    #print("CalcH")
    h = 0
    index = 0
    name = turn.pop()
    worksheet = workbook.get_worksheet_by_name(name)
    worksheet.write(0,4,"Total Horas")
    #print(name)
    while(index < len(turn)-1):
        #Se os dias forem o mesmo
        if(turn[index][0] == turn[index+1][0]):
            tmp = turn[index+1][1] - turn[index][1]
            h += tmp
            index+=2
        #Se os dias forem diferentes mas conta para o mesmo dia
        elif((turn[index][0] != turn[index+1][0]) and (turn[index+1][1] >= 24)):
            tmp = turn[index+1][1] - turn[index][1]
            h += tmp
            index+=2
        else:
            index+=1
    tmp = ch.minToHours(h)
    worksheet.write(1,4,tmp)

def main(file):
    names = getNames(file)
    names.pop()
    for i in names:
        turn = getTurn(file,i)
        calcH(turn)
    file.seek(0)
    file.close()
    workbook.close()

