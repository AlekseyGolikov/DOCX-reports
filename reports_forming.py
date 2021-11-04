import pandas as pd

file=pd.read_excel('sourses/sourse.xlsx')
# print(file.head())
data_frame = pd.DataFrame(file)
# print(data_frame)
# dict[data_frame[1][0]]=data_frame[1][1]
# print(data_frame['Основное средство'][0])

from docxtpl import DocxTemplate


observBuilding1 = 'осмотр несущих конструкций, свайного основания, стеновых покрытий, дверных и'
observBuilding2 = 'оконных проемов.'

observMacht = 'осмотр несущих конструкций, свайного основания.'
damageList = ('ДП2001490','ДП2001510','ДП2001406','ДП2001403','ДП2001401','АА8639','ДП2001547','ДП2001548','ДП2001549','ДП2001550')

for i in range(len(data_frame)):
    doc = DocxTemplate("sourses/template.docx")
    unit = str(data_frame['Основное средство'][i])
    # if 'ОЯЯНГКМ. 1. 1. ' in unit: 
    #     unit = unit.replace('ОЯЯНГКМ. 1. 1. ','')
    if '"' in unit:
        unit = unit.replace('"','')
    
    if 'Емкость дренажно-канализационная' in unit:
        unit = unit.replace('Емкость дренажно-канализационная','Блок-бокс над емкостью дренажно-канализационной')
    elif 'Сбор и транспорт газа' in unit:
        unit = unit.replace('Сбор и транспорт газа','БКУЭ кранового узла ТГП')
    
    if 'ОЯЯНГКМ. 1. 1. ' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 1. ','')
    elif 'ОЯЯНГКМ. 1. 1.' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 1.','')
    elif 'ОЯЯНГКМ. 1.1. ' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1.1. ','')
    elif 'ОЯЯНГКМ. 1. 2. ' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 2. ','')
    elif 'ОЯЯНГКМ. 1. 2.' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 2.','')
    elif 'ОЯЯНГКМ. 1. 8. ' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 8. ','')
    elif 'ОЯЯНГКМ. 1. 8.' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 8.','')
    elif 'ОЯЯНГКМ. 1. 11. ' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 11. ','')
    elif 'ОЯЯНГКМ. 1. 11.' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 11.','')
    elif 'ОЯЯНГКМ. 1. 12. ' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 12. ','')
    elif 'ОЯЯНГКМ. 1. 12.' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 12.','')
    elif 'ОЯЯНГКМ. 1. 17. ' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 17. ','')
    elif 'ОЯЯНГКМ. 1. 17.' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 17.','')
    elif 'ОЯЯНГКМ. 1. 18. ' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 18. ','')   
    elif 'ОЯЯНГКМ. 1. 18.' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 18.','')
    elif 'ОЯЯНГКМ. 1. 19. ' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 19. ','')
    elif 'ОЯЯНГКМ. 1. 19.' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 19.','')
    elif 'ОЯЯНГКМ. 1. 20. ' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 20. ','')
    elif 'ОЯЯНГКМ. 1. 20.' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 20.','')
    elif 'ОЯЯНГКМ. 1. 21. ' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 21. ','')
    elif 'ОЯЯНГКМ. 1. 21.' in unit:
        unit = unit.replace('ОЯЯНГКМ. 1. 21.','')
    elif 'ОНЯЯНГКМ.' in unit:
        unit = unit.replace('ОНЯЯНГКМ.','')
    elif 'ОЯЯНГКМ. ' in unit:
        unit = unit.replace('ОЯЯНГКМ. ','')
    elif 'ОЯЯНГКМ.' in unit:
        unit = unit.replace('ОЯЯНГКМ.','')
    elif ' Технологические сооружения.' in unit:
        unit = unit.replace(' Технологические сооружения.','')


    if ('БПО' in unit) or ('ДП' in unit):
        date = '8 ноября 2021г.'
    elif ('КЭ' in unit) or ('ПР' in unit):
        date = '9 ноября 2021г.'
    elif ('КОС' in unit) or ('ППС' in unit):
        date = '10 ноября 2021г.'
    elif ('КС' in unit) or ('ДКС' in unit):
        date = '11 ноября 2021г.'
    elif ('УДК' in unit) or ('УКПГ' in unit):
        date = '12 ноября 2021г.'
    elif ('УПН' in unit) or ('СОВ' in unit):
        date = '15 ноября 2021г.'
    elif ('ООВЭиОЭН' in unit) or ('№ 7' in unit):
        date = '18 ноября 2021г.'
    elif ('№11' in unit) or ('№ 11' in unit) or ('№3' in unit)  or ('№12' in unit)  or ('№ 12' in unit):
        date = '19 ноября 2021г.'
    elif ('№4' in unit) or ('№ 4' in unit)  or ('№15' in unit) or ('№ 15' in unit):
        date = '20 ноября 2021г.'
    elif ('№2' in unit) or ('№ 2' in unit)  or ('№5' in unit) or ('№ 5' in unit)  or ('№71' in unit) or ('№ 71' in unit):
        date = '20 ноября 2021г.'
    elif ('№7' in unit) or ('№ 7' in unit) :
        date = '20 ноября 2021г.'

    

    if 'ачта' in unit:
        # context = { 'unit' : data_frame['Основное средство'][i] , 'code': data_frame['Инвентарный номер'][i] , 'date' : date, 'observ1' : observMacht, 'observ2' : '', 'obj' : 'сооружения'}
        observ1 = observMacht
        observ2 = ''
        obj = 'сооружения'
    elif 'олниеотвод' in unit:
        # context = { 'unit' : data_frame['Основное средство'][i] , 'code': data_frame['Инвентарный номер'][i] , 'date' : date, 'observ1' : observMacht, 'observ2' : '', 'obj' : 'сооружения'}
        observ1 = observMacht
        observ2 = ''
        obj = 'сооружения' 
    else:
        observ1 = observBuilding1
        observ2 = observBuilding2
        obj = 'здания' 
    context = { 'unit' : unit , 'code': str(data_frame['Инвентарный номер'][i]) , 'date' : date, 'observ1' : observ1, 'observ2' : observ2, 'obj' : obj}
    doc.render(context)
    if str(data_frame['Инвентарный номер'][i]) not in damageList:
        path_to_file = 'reports/' + unit + '.docx'
    else:
        path_to_file = 'reports/damages/' + unit + '.docx'