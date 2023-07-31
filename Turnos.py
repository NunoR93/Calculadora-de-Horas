import datetime
from datetime import *

import xlsxwriter


# Função para verificar a falta de picagens
def checkPicagens(list):
    picagens_falta = []
    previous = list.pop(0)
    pic = 0
    while (len(list) > 1):
        try:
            pic = list.pop(0)
        except:
            picagens_falta += [pic]
        if (pic.date() != previous.date()):
            picagens_falta += [str(previous.date())]
            previous = pic
        else:
            previous = list.pop(0)
    # Este bloco serve para verificar se a ultima picagem esta correta
    # ou se ha alguma em falta devido ao ciclo for parar antes de completar a lista
    if (pic.date() != previous.date() and len(list) == 0):
        picagens_falta += [str(previous.date())]
    elif (len(list) == 1):
        pic = list.pop()
        if (pic.date() != previous.date()):
            picagens_falta += [str(previous.date())]
    return picagens_falta


def totalHoras(datas):
    total = timedelta(0)
    stringH = ""
    while (len(datas) > 1):
        previous = datas.pop(0)
        pic = datas.pop(0)
        total += pic - previous
    total = total.total_seconds()
    minutes, seconds = divmod(total, 60)
    hours, minutes = divmod(minutes, 60)
    stringH = "%02d:%02d:%02d" % (hours, minutes, seconds)
    return stringH


def createExcel(log,save):
    file = open(log, "r")
    next(file)
    empregado = ""
    picagens = []
    faltas = []
    index = -1
    path = save+"/Horas.xlsx"
    print(path)
    workbook = xlsxwriter.Workbook(path)

    # Percorre ficheiro e mete num dicionario com as keys nome e datas os valores do empregado e das picagens
    for i in file:
        tmp = i.replace("\n", "").split("\t")
        nome = tmp[0]
        pica = tmp[1]
        if (nome != empregado):
            empregado = nome
            picagem = {'nome': nome, 'datas': [datetime.strptime(pica, "%d-%m-%Y %H:%M:%S")]}
            picagens += [picagem]
            index += 1
        else:
            picagens[index]['datas'] += [datetime.strptime(pica, "%d-%m-%Y %H:%M:%S")]

    headformat = workbook.add_format({'align': 'center'})
    headformat.set_bg_color('gray')
    cellformat = workbook.add_format({'align': 'center'})
    cellformat.set_bg_color('white')

    for i in picagens:
        tmp = list(i['datas'])
        faltas += checkPicagens(tmp)
        worksheet = workbook.add_worksheet(i['nome'])
        worksheet.set_column(0, 2, 10)
        worksheet.set_footer(i['nome'])
        index = 2
        if len(faltas) == 0:
            total = totalHoras(list((i['datas'])))
            worksheet.write('A1', "Data", headformat)
            worksheet.write('B1', "Hora", headformat)
            worksheet.write('C1', "Total", headformat)
            worksheet.write('C2', total, cellformat)
            for j in i['datas']:
                dateformat = workbook.add_format(
                    {'num_format': 'dd-mm-yyyy', 'align': 'center', 'border': True, 'border_color': 'black'})
                hourformat = workbook.add_format(
                    {'num_format': 'hh:mm:ss', 'align': 'center', 'border': True, 'border_color': 'black'})
                if index % 2 != 0:
                    dateformat.set_bg_color('#b4b4b8')  # lightGray
                    hourformat.set_bg_color('#b4b4b8')
                else:
                    dateformat.set_bg_color('white')
                    hourformat.set_bg_color('white')
                worksheet.write_datetime('A' + str(index), j.date(), dateformat)
                worksheet.write_datetime('B' + str(index), j.time(), hourformat)
                index += 1
        else:
            for y in faltas:
                worksheet.write('A' + str(index), y)
                index += 1
            faltas = []
    workbook.close()
    file.close()
