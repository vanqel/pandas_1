import datetime
import os
import pandas as pd

def loadingproc(name, proc):
    print(name+" |"+"="*(int(proc))+">"+"-"*(100-int(proc))+"| "+str(proc)+" %")

timestart = datetime.datetime.now().strftime("%HH:%MM")
print("*", 'TIMESTART:',timestart , "*")
print("*", 'START ANALYZE?',timestart , "*")
timestart = datetime.datetime.now().strftime("%HH:%MM")
input()

print("XLSX OPENING")
mediaqd965 = pd.read_excel(r"Задание2.xlsx", header=0)
mediaqd965.set_index('Товар', inplace=True)

result = {}
uniq = {}
uniq_result = {}
print("XLSX OPEN")

c=0
for index, value in mediaqd965.iterrows():
    try:
        uniq[str(index)] = {'ID': value['ID'],
                            'Дата': value['Дата'],
                            'Город': value['Город']}
        uniq_result[str(index)] = {}
    except:
        pass
    c+=1
    loadingproc("PROGRAM STEP 1",(c / len(mediaqd965)) * 100)

c=0
for index, value in mediaqd965.iterrows():
    try:
        if uniq[index]['Дата'] == value['Дата'] and uniq[index]['Город'] == value['Город']:
            uniq_result[index].update({str(value['id']):{'Дата': value['Дата'],
                                                 'Город': value['Город'],}})
    except:
        pass
    c += 1
    loadingproc("PROGRAM STEP 2",(c / len(mediaqd965)) * 100)


c=0
for i,v in uniq_result.items():
    try:
        if len(v)>=2:
            for index,value in v.items():
                result[index] = {'Товар':i,
                                 'Дата':value['Дата'],
                                 'Город':value['Город']}
    except Exception as ex:
        print(ex)
    c += 1
    loadingproc("PROGRAM STEP 3", (c / len(mediaqd965)) * 100)

excel_wr = pd.ExcelWriter('Ответ.xlsx')
result_csv = pd.DataFrame.from_dict(result, orient='index')
result_csv.to_excel(excel_wr, sheet_name='Выборка', index_label='ID')
excel_wr.close()
timeend = datetime.datetime.now().strftime("%HH:%MM")


os.system('cls')
print("*"*16)
print("*", 'COMPLETE' , "*")
print("*", 'TIMESTART:',timestart, "*")
print("*", 'TIMEEND:',timeend, "*")
print("*", result_csv.head(5), "*")
print("*"*16)