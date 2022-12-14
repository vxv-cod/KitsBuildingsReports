import os
import re
import sys
from time import sleep
from time import time
import requests
import win32com.client
from requests_negotiate_sspi import HttpNegotiateAuth
# import threading

from rich import print

PT = 28.34646
HeightRule = 1
RowHeigh = 0.8 * PT

def decorTime(my_func):
    '''Обертка функции декоратором'''
    # *args — это сокращение от «arguments» (аргументы) в виде кортежа, 
    # **kwargs — сокращение от «keyword arguments» (именованные аргументы)
    def wrapper(*args):
        start_time = time()
        my_func(*args)
        print(f'"Выполнено: {round(time() - start_time, 2)} sec"')
    return wrapper

def importdata(sheet, StartRow, StartCol, EndRow, EndCol):
    '''Собираем данные из диапозона ячеек'''
    cel = sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol))
    # vals = cel.Formula
    vals = cel.Value
    if StartCol == EndCol:
        vals = [vals[i][x] for i in range(len(vals)) for x in range(len(vals[i]))]
    return vals


def exportdata(data, sheet, StartRow, StartCol, EndRow, EndCol):
    '''Отправляем данные в диапозон ячеек'''
    cel = sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol))
    # cel.Formula = data
    cel.Value = data
    return cel


def RangeCells(sheet, StartRow, StartCol, EndRow, EndCol):
    '''Выделяем диапозон ячеек'''
    cel = sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol))
    return cel


def resp(url):
    response = requests.get(url, auth=HttpNegotiateAuth())
    if response.status_code == 404:
        # print("'Страница не существует!'")
        response = 'HELP!!!'
        return response
    if response.status_code == 401:
        print("'Ошибка авторизации пользователя!'")
        # response = None
        return 
    # print(response)
    response = response.json()
    return response


def breakPage(Doc):
    '''Разрывам страницу'''
    '''Перемещаемся на последний параграф'''
    myRange = Doc.Paragraphs(Doc.Paragraphs.Count).Range
    '''Схлапываем выделение параграффа в его конец'''
    myRange.Collapse(0)
    '''Разрыв страницы'''
    myRange.InsertBreak()    
    '''Удаляем прпграфф перед разрывам страницы'''
    Doc.Paragraphs(Doc.Paragraphs.Count - 1).Range.Select()
    Selection = Doc.Application.Selection
    Selection.Collapse(1)     # '''Схлапываем выделение параграффа в его начало'''
    Selection.TypeBackspace()


def get_dictionary_item_Title(Id):
    '''Список Id словарей'''
    global dictionary_items
    for i in dictionary_items:
        if i['Id'] == Id:
            return (i['Title'])


def min_max(sheet, StartRow, StartCol, EndRow, EndCol):
    '''Объединение ячеек по вертикали с по объектам КО'''
    celMergeCol = importdata(sheet, StartRow, StartCol, EndRow, EndCol)

    ind_min_max_list = []
    xxx = []
    for i in range(len(celMergeCol)):
        if celMergeCol[i] != None:
            if len(xxx) > 1:
                ind_min_max_list.append([min(xxx), max(xxx)])
                xxx = []
        if celMergeCol[i] == None:
            pass
        xxx.append(i + StartRow)
        if i == len(celMergeCol) - 1:
            ind_min_max_list.append([min(xxx), max(xxx)])

    return ind_min_max_list


def ExcelFormat(wb, sheetName, data, StartRow, CopyStartRow):
    sheet = wb.Worksheets(sheetName)
    sheet.Activate()
    '''Индексы первой ячейки (строка, стоблец) для вставки данных'''
    StartCol = 1
    '''Индексы последней ячейки (строка, стоблец) для вставки данных'''
    EndRow = len(data) + StartRow - 1
    EndCol = len(data[0])
    '''Отправляем data в Excel'''
    cel = exportdata(data, sheet, StartRow, StartCol, EndRow, EndCol)
    '''Рисуем границы ячеек во всей таблице'''
    cel.Borders.Weight = 2
    sleep(0.3)
    '''Выделяем и копируем в буфер обмена всю таблицу из Excel'''
    tabEx = RangeCells(sheet, CopyStartRow, StartCol, EndRow, EndCol)
    tabEx.EntireRow.AutoFit()
    sleep(0.3)    
    tabEx.Copy()
    sleep(0.3)    


def WordFormat(NazvanieTab, Doc, Prim=None):
    '''Перемещаемся на последний параграф'''
    myRange = Doc.Paragraphs(Doc.Paragraphs.Count).Range
    '''Вставляем заголовок таблицы'''
    myRange.Text = NazvanieTab
    sleep(0.3)
    myRange.ParagraphFormat.Alignment = 2
    '''Перемещаемся на последний параграф'''
    sleep(0.3)
    myRange = Doc.Paragraphs(Doc.Paragraphs.Count).Range
    myRange.Font.Bold = True
    '''Вставляем скопированную таблицу в World'''
    myRange.PasteExcelTable(False, False, False)
    
    tabWord = Doc.Tables(Doc.Tables.Count)
    '''Автоподбор размера ячейки таблицы'''
    tabWord.AutoFitBehavior(2)    # по ширине активного окна
    tabWord.Rows.Height = 0.5 * PT
    '''Поля в ячейках таблицы'''
    tabWord.TopPadding = 0.05 * PT
    tabWord.BottomPadding = 0.05 * PT
    tabWord.LeftPadding = 0.1 * PT
    tabWord.RightPadding = 0.1 * PT
    sleep(0.3)

    tabWord.Rows.HeightRule = HeightRule   # указывает на способ изменения высоты
    tabWord.Rows.Height = RowHeigh         # RowHeight указывает на новую высоту строки в пунктах.

    '''Границы ячеек таблицы'''
    tabWord.Borders(-5).LineStyle = 1
    if Prim != None:
        '''Перемещаемся на последний параграф'''
        myRange = Doc.Paragraphs(Doc.Paragraphs.Count).Range
        '''Вставляем Примечание'''
        '''Добавить в конец параграф'''
        myRange.InsertAfter(Prim)
        myRange.ParagraphFormat.Alignment = 0
        myRange.Font.Bold = False
    '''Поля в ячейках таблицы'''
    tabWord.TopPadding = 5
    tabWord.BottomPadding = 5
    return tabWord

    
# def insertText(text, Bold, Align, Size):
#     myRange = Doc.Paragraphs(Doc.Paragraphs.Count).Range
#     myRange.Font.Bold = Bold
#     myRange.Font.Size = Size
#     myRange.InsertAfter(text)
#     myRange.ParagraphFormat.Alignment = Align

# insertText(text, True, 0, 16)
# insertText(NazvanieTab, True, 2, 11)



'''Площадные объекты'''
def dataTable_31(response):
    '''Таблица 3 - Идентификация зданий и сооружений площадочных и линейных объектов'''
    '''Выбираем нужные данные собираем их в список "data"'''
    global response_dop
    counterList = []
    counter = 5
    MergeCol5 = 5
    data = []
    temp = ''
    for i in response:
        MergeCol5 += 1
        # КО
        Description = i['KoItem']['Description']
        if Description != temp:
            data.append([Description] + [None] * 8)
            temp = Description
            counter += 1
            counterList.append(counter)

        xxx = [
            # № П/П
            i['GenplanNumber'],
            # ЗДАНИЕ/ СООРУЖЕНИЕ
            i['Title'],
            # НАЗНАЧЕНИЕ
            i['Appointment'],
            # ПРИНАДЛЕЖНОСТЬ К ОБЪЕКТАМ ТРАНСПОРТНОЙ ИНФРАСТРУКТУРЫ И К 
            'Да' if i['BelongingTransportFacilities'] == True else 'Нет',
            # ВОЗМОЖНОСТЬ ОПАСНЫХ ПРИРОДНЫХ ПРОЦЕССОВ И ЯВЛЕНИЙ И ТЕХНОГЕННЫХ ВОЗДЕЙСТВИЙ НА ТЕРРИТОРИИ
            # 'HELP!!!',
            response_dop['PossiblityDangerousNaturalProcesses'],
            # ПРИНАДЛЕЖНОСТЬ К ОПАСНЫМ ПРОИЗВОДСТВЕННЫМ ОБЪЕКТАМ
            'Да' if i['BelongingHazardousIndustries'] == True else 'Нет',
            # ПОЖАРНАЯ И ВЗРЫВОПОЖАРНАЯ ОПАСНОСТЬ (Категория здания/сооружения по взрывопожароопасности)
            get_dictionary_item_Title(i['CategoryBuildingFireId']),
            # НАЛИЧИЕ ПОМЕЩЕНИЙ С ПОСТОЯННЫМ ПРЕБЫВАНИЕМ ЛЮДЕЙ 
            'Да' if i['AvailabilityRoomsPeople'] == True else 'Нет',
            # УРОВЕНЬ ОТВЕТСТВЕННОСТИ            '''Неправильно определяется уровень ответственности'''
            get_dictionary_item_Title(i['CategoryLevelResponsibilityId'])
            
        ]
        data.append(xxx)
        counter += 1
    # print(data)
    # data = sorted(data, key = lambda i: i[0])
    return data, counterList, counter


'''Линейные объекты'''
def dataTable_32(response):
    '''Таблица 3 - Идентификация зданий и сооружений площадочных и линейных объектов'''
    '''Выбираем нужные данные собираем их в список "data"'''
    global response_dop

    data = [['Линейные объекты'] + [None] * 8]
    
    for i in response:
        xxx = [
            # № П/П
            i['SortIndex'],
            # ЗДАНИЕ/ СООРУЖЕНИЕ
            # i['Title'],
            f"""{i['Title']}\n{i['StartText']}\n{i['FinishText']}""",
            # НАЗНАЧЕНИЕ
            i['PurposeText'],
            # ПРИНАДЛЕЖНОСТЬ К ОБЪЕКТАМ ТРАНСПОРТНОЙ ИНФРАСТРУКТУРЫ И К 
            'Да' if i['IsTransportObject'] == True else 'Нет',
            # ВОЗМОЖНОСТЬ ОПАСНЫХ ПРИРОДНЫХ ПРОЦЕССОВ И ЯВЛЕНИЙ И ТЕХНОГЕННЫХ ВОЗДЕЙСТВИЙ НА ТЕРРИТОРИИ
            response_dop['PossiblityDangerousNaturalProcesses'],
            # ПРИНАДЛЕЖНОСТЬ К ОПАСНЫМ ПРОИЗВОДСТВЕННЫМ ОБЪЕКТАМ
            'Да' if i['IsDangerObject'] == True else 'Нет',
            # ПОЖАРНАЯ И ВЗРЫВОПОЖАРНАЯ ОПАСНОСТЬ (Категория здания/сооружения по взрывопожароопасности)
            get_dictionary_item_Title(i['CategoryBuildingFireId']),
            # НАЛИЧИЕ ПОМЕЩЕНИЙ С ПОСТОЯННЫМ ПРЕБЫВАНИЕМ ЛЮДЕЙ 
            'Да' if i['IsPeopleRoomExists'] == True else 'Нет',
            # УРОВЕНЬ ОТВЕТСТВЕННОСТИ            '''Неправильно определяется уровень ответственности'''
            get_dictionary_item_Title(i['CategoryLevelResponsibilityId'])
        ]
        data.append(xxx)

    # data = sorted(data, key = lambda i: int(i[0]))

    return data


def dataTable_3(response1, response2, wb, Doc):
    global tableName

    '''Индексы первой ячейки (строка, стоблец) для вставки данных'''
    StartRow = 6
    StartCol = 1

    if response1 != [] and response2 == []:
        data1, counterList1, counter = dataTable_31(response1)
        data = data1
        counterList = counterList1

    if response1 != [] and response2 != []:
        data1, counterList1, counter = dataTable_31(response1)
        data2 = dataTable_32(response2)
        counterList = counterList1 + [counter + 1]
        data = [i for g in [data1, data2] for i in g]

    if response1 == [] and response2 != []:
        data2 = dataTable_32(response2)
        data = data2
        counterList = [StartRow]

    '''------------------- Excel -------------------'''

    sheet = wb.Worksheets("Таблица (3)")
    sheet.Activate()

    '''Индексы последней ячейки (строка, стоблец) для вставки данных'''
    EndRow = len(data) + StartRow - 1
    EndCol = len(data[0])
    '''Отправляем data в Excel'''
    cel = exportdata(data, sheet, StartRow, StartCol, EndRow, EndCol)
    '''Рисуем границы ячеек во всей таблице'''
    cel.Borders.Weight = 2
    sleep(1)

    '''Объединение ячеек по вертикали с одинаковыми значениями (колонка 5)'''
    celMergeCol = importdata(sheet, StartRow, 5, EndRow, 5)
    celMergeCol.append(None)
    
    ind_min_max_list = []
    textCelMergeCol = []
    xxx = []
    text = ''
    for i in range(len(celMergeCol)):
        if celMergeCol[i] != None:
            xxx.append(i + StartRow)
            text = celMergeCol[i]

        if celMergeCol[i] == None:
            if xxx != [] and len(xxx) > 1:
                ind_min_max_list.append([min(xxx), max(xxx)])
                textCelMergeCol.append(text)
            xxx = []

    for i in ind_min_max_list:
        cel = RangeCells(sheet, i[0], 5, i[1], 5)
        sleep(0.1)
        cel.Merge()

    '''Объединение ячеек по всей строке таблице с название КО'''
    for i in counterList:
        cel = RangeCells(sheet, i, StartCol, i, EndCol)
        sleep(0.1)
        cel.Merge()
        sleep(0.3)
        cel.HorizontalAlignment = 1
        cel.Font.Bold = True
    
    '''Выделяем и копируем в буфер обмена всю таблицу из Excel'''
    tabEx = RangeCells(sheet, 4, StartCol, EndRow, EndCol)
    tabEx.Copy()
    sleep(1)

    NazvanieTab =   f'''Таблица {3}
                    Идентификация зданий и сооружений площадочных и линейных объектов
                    (Федеральный закон № 384 «Технический регламент о безопасности зданий и сооружений»)
                    '''
    tableName.append(' '.join([i.lstrip() for i in NazvanieTab.split('\n')[:3]]))
    
    '''------------------- Word -------------------'''

    Doc.Activate()
    '''Выбираем последний параграф'''
    myRange = Doc.Paragraphs(Doc.Paragraphs.Count).Range
    '''Установка интервала перед и после абзаца в тексте. Единица измерения интервала пт.'''
    myRange.ParagraphFormat.SpaceBefore = 0  #интервал перед
    myRange.ParagraphFormat.SpaceAfter = 0   #интервал после
    '''Вставляем заголовок таблицы'''
    myRange.Text = NazvanieTab
    myRange.ParagraphFormat.Alignment = 2
    myRange.Font.Bold = True
    '''Перемещаемся на последний параграф'''
    CountP = Doc.Paragraphs.Count
    myRange = Doc.Paragraphs(CountP).Range
    '''Вставляем скопированную таблицу в World'''
    myRange.PasteExcelTable(False, False, False)
    sleep(0.5)
    '''Подключаемся к таблице'''
    tabWord = Doc.Tables(Doc.Tables.Count)
    '''Выравниваем по вертикали все ячейки в таблице'''
    tabWord.Range.Cells.VerticalAlignment = 1
    '''Удаляем интервал после абзаца во всей таблице'''
    tabWord.Range.ParagraphFormat.SpaceBefore = 0  # интервал перед
    tabWord.Range.ParagraphFormat.SpaceAfter = 0  # интервал после
    '''Границы ячеек таблицы'''
    tabWord.Borders(-5).LineStyle = 1
    '''Повторять как заголовок на каждой странице'''
    tabWord.Cell(2, 1).Range.Rows.HeadingFormat = True
    '''Установка единицы измерения размера таблицы'''
    tabWord.PreferredWidthType = 3
    '''Автоподбор размера ячейки таблицы'''
    tabWord.AutoFitBehavior(0)      # фиксированный размер
    '''При работе с таблицами в сантиметрах'''
    PT = 28.34646  # количество "пт" в см
    '''Установка единицы измерения размера таблицы''' 
    tabWord.PreferredWidthType = 3    # CM
    '''Список размеров колонок'''
    # WidthList = [1.0, 5.0, 5.0, 3.0, 4.0, 2.7, 2.7, 2.7, 2.7]
    WidthList = [1.0, 5.0, 3.5, 3.5, 5.0, 2.7, 2.7, 2.7, 2.7]
    '''Задаем общую ширину таблицы для более точного определения'''
    tabWord.PreferredWidth = sum(WidthList) * PT
    '''Проходим по всем колонкам для установки размеров из списка'''
    for i in range(1, len(WidthList) + 1):
        col = tabWord.Cell(1, i).Range.Columns
        col.PreferredWidthType = 3
        col.PreferredWidth = WidthList[i-1] * PT
    '''Поля в ячейках таблицы'''
    tabWord.TopPadding = 0.05 * PT
    tabWord.BottomPadding = 0.05 * PT
    tabWord.LeftPadding = 0.1 * PT
    tabWord.RightPadding = 0.1 * PT

    tabWord.Rows.HeightRule = HeightRule   # указывает на способ изменения высоты
    tabWord.Rows.Height = RowHeigh         # RowHeight указывает на новую высоту строки в пунктах.

    l23 = tabWord.Cell(2, 1).Range.Rows
    l23.HeightRule = 2
    l23.Height = 0.5 * PT

    # l23.TopPadding = 0
    # l23.BottomPadding = 0
    '''Разрывам страницу'''
    breakPage(Doc)



def dataTable_4(response, wb, Doc, tableNomer):
    '''Таблица 4 - Топографическая съемка площадочных объектов'''
    global tableName

    counterList = []
    counter = 5
    MergeCol5 = 5
    data = []
    temp = ''
    for i in response:
        MergeCol5 += 1
        # КО
        Description = i['KoItem']['Description']
        if Description != temp:
            data.append(
                [
                Description,
                None,
                response_dop['CharacteristicTerritory'],
                'Сложная конфигурация',
                None,
                response_dop['ApproximateShootingArea'],
                'План 1:500\nПрофиль Мг 1:500,\nМв 1:100,\nМгео 1:100.',
                '0,5',
                response_dop['AdditionalOrSpecialRequirements']
                ]
                )
            temp = Description
            counter += 1
            counterList.append(counter)
        # Размеры
        if i['Diameter'] != None:
            try:
                size1 = str(round(float(i['Diameter']) / 1000 + 0.0, 3)).replace('.', ',')
                size2 = size1
            except:
                size1 = size2 = None
        else:
            try:
                size1 = str(round(float(i['Length']) / 1000 + 0.0, 3)).replace('.', ',')
            except:
                size1 = None
            try:
                size2 = str(round(float(i['Width']) / 1000 + 0.0, 3)).replace('.', ',')
            except:
                size2 = None
        
        xxx = [
            # № П/П
            i['GenplanNumber'],
            # ЗДАНИЕ/ СООРУЖЕНИЕ
            i['Title'],
            None,
            # Длина
            size1,
            # Ширина
            size2,
            None,
            None,
            None,
            None
        ]

        data.append(xxx)
        counter += 1

    '''------------------- Excel -------------------'''
    # wb = Excel.ActiveWorkbook
    sheet = wb.Worksheets("Таблица (4)")
    sheet.Activate()
    '''Индексы первой ячейки (строка, стоблец) для вставки данных'''
    StartRow = 6
    StartCol = 1
    '''Индексы последней ячейки (строка, стоблец) для вставки данных'''
    EndRow = len(data) + StartRow - 1
    EndCol = len(data[0])
    '''Отправляем data в Excel'''
    cel = exportdata(data, sheet, StartRow, StartCol, EndRow, EndCol)
    '''Рисуем границы ячеек во всей таблице'''
    cel.Borders.Weight = 2
    sleep(1)

    Col = 3
    ind_min_max_list = min_max(sheet, StartRow, Col, EndRow, Col)
    ColMergeList = [3, 6, 7, 8, 9]

    '''Объединение ячеек в колонках таблице'''
    for Col in ColMergeList:
        for index in ind_min_max_list:
            cel = RangeCells(sheet, index[0], Col, index[1], Col)
            sleep(0.1)
            cel.Merge()
            sleep(0.1)
    '''Объединение ячеек по всей строке таблице с название КО'''
    for row in counterList:
        cel = RangeCells(sheet, row, 1, row, 2)
        sleep(0.1)
        cel.Merge()
        cel.HorizontalAlignment = 1
        cel.Font.Bold = True
        cel = RangeCells(sheet, row, 4, row, 5)
        sleep(0.1)
        cel.Merge()
        cel.HorizontalAlignment = 1
    

    '''Выделяем и копируем в буфер обмена всю таблицу из Excel'''
    tabEx = RangeCells(sheet, 3, StartCol, EndRow, EndCol)
    tabEx.Copy()
    sleep(0.5)

    tableNomer += 1
    NazvanieTab =   f'''Таблица {tableNomer}
                    Топографическая съемка площадочных объектов
                    '''
    tableName.append(' '.join([i.lstrip() for i in NazvanieTab.split('\n')[:2]]))
    Prim = '''Примечание: Площадь съемки указывается с округлением до 0,1 га.'''
    tabWord = WordFormat(NazvanieTab, Doc, Prim)
    '''Высота строки с номерами колонок'''
    tabWord.Cell(3, 1).Height = 0.5 * PT
    '''Повторять как заголовок на каждой странице'''
    tabWord.Cell(3, 1).Range.Rows.HeadingFormat = True

    l23 = tabWord.Cell(3, 4).Range.Rows
    l23.HeightRule = 2
    l23.Height = 0.5 * PT

    '''Разрывам страницу'''
    breakPage(Doc)
    return tableNomer


def dataTable_5(response, wb, Doc, tableNomer):
    '''Таблица 5 - Топографическая съемка линейных объектов'''
    global tableName
    
    data = []
    for i in response:
        xxx = [
            # № по ГП
            i['SortIndex'],
            # НАИМЕНОВАНИЕ ТРАССЫ, ЕЁ НАЧАЛЬНЫЕ И КОНЕЧНЫЕ ПУНКТЫ 
            # i['Title'],
            f"""{i['Title']}\n{i['StartText']}\n{i['FinishText']}""",
            # ПРЕДВАРИТЕЛЬНАЯ ПРОТЯЖЕННОСТЬ ТРАССЫ, КМ
            i['Length'],
            # ШИРИНА ПОЛОСЫ СЪЕМКИ, М
            i['TakeOffWidthText'],
            # МАСШТАБ СЪЕМКИ
            i['TakeOffScaleText'],
            # СЕЧЕНИЕ РЕЛЬЕФА, М
            i['ReliefSize'],
            # МАСШТАБ ПРОДОЛЬНОГО ПРОФИЛЯ
            i['ProfileLongScaleText'],
            # ДОПОЛНИТЕЛЬНЫЕ ИЛИ ОСОБЫЕ ТРЕБОВАНИЯ
            i['AddRequiresText']
        ]
        data.append(xxx)
    ExcelFormat(wb, "Таблица (5)", data, StartRow=5, CopyStartRow=3)
    tableNomer += 1
    NazvanieTab =   f'''Таблица {tableNomer}
                    Топографическая съемка линейных объектов
                    '''
    tableName.append(' '.join([i.lstrip() for i in NazvanieTab.split('\n')[:2]]))
    Prim = '''Примечание: Протяженность указывается с округлением до 0,1 км.'''
    tabWord = WordFormat(NazvanieTab, Doc, Prim)
    '''Высота строки с номерами колонок'''
    tabWord.Cell(3, 1).Height = 0.5 * PT
    '''Повторять как заголовок на каждой странице'''
    tabWord.Cell(3, 1).Range.Rows.HeadingFormat = True
    '''Корректируем высоту строк (ячеек)'''
    l23 = tabWord.Cell(2, 1).Range.Rows
    l23.HeightRule = 2
    l23.Height = 0.5 * PT    
    '''Разрывам страницу'''
    breakPage(Doc)
    return tableNomer


def dataTable_6(response, wb, Doc, tableNomer):
    '''Таблица 5 - Топографическая съемка линейных объектов'''
    global tableName
    data = []
    for i in response:
        xxx = [    
            # № по ГП
            i['SortIndex'],
            # НАИМЕНОВАНИЕ ТРАССЫ 
            # i['Title'],
            f"""{i['Title']}\n{i['StartText']}\n{i['FinishText']}""",
            # ПРОТЯЖЕННОСТЬ ТРАССЫ, КМ
            i['Length'],
            # Столбец 4
            i['BasementDetailsText'],
            # Диаметр
            i['DiameterText'],
            # ДАВЛЕНИЕ
            i['Pressure'],
            # МАТЕРИАЛЬНОЕ ИСПОЛНЕНИЕ
            i['MaterialText'],
            # ОСОБЫЕ УСЛОВИЯ СТРОИТЕЛЬСТВ
            i['SpecialConditionsText']
        ]
        data.append(xxx)
    # wb = Excel.ActiveWorkbook
    sheet = wb.Worksheets("Таблица (6)")
    sheet.Activate()
    '''Индексы первой ячейки (строка, стоблец) для вставки данных'''
    StartRow = 6
    StartCol = 1
    '''Индексы последней ячейки (строка, стоблец) для вставки данных'''
    EndRow = len(data) + StartRow - 1
    EndCol = len(data[0])
    '''Отправляем data в Excel'''
    cel = exportdata(data, sheet, StartRow, StartCol, EndRow, EndCol)
    '''Рисуем границы ячеек во всей таблице'''
    cel.Borders.Weight = 2
    sleep(1)
    '''Выделяем и копируем в буфер обмена всю таблицу из Excel'''
    tabEx = RangeCells(sheet, 3, StartCol, EndRow, EndCol)
    tabEx.Copy()
    sleep(0.5)

    tableNomer += 1
    NazvanieTab =   f'''Таблица {tableNomer}
                    Техническая характеристика линейных объектов для инженерно-геологических изысканий
                    '''
    tableName.append(' '.join([i.lstrip() for i in NazvanieTab.split('\n')[:2]]))
    Prim = '''Примечание: Протяженность указывается с округлением до 0,1 км.'''
    tabWord = WordFormat(NazvanieTab, Doc, Prim)
    '''Высота строки с номерами колонок'''
    tabWord.Cell(3, 1).Height = 0.5 * PT
    '''Повторять как заголовок на каждой странице'''
    tabWord.Cell(3, 1).Range.Rows.HeadingFormat = True
    '''Установка единицы измерения размера таблицы''' 
    tabWord.PreferredWidthType = 3
    '''Список размеров колонок'''
    WidthList = [0.8, 5.5, 3.0, 7.5, 2.5, 2.5, 2.5, 4.5]
    '''Задаем общую ширину таблицы для более точного определения'''
    tabWord.PreferredWidth = sum(WidthList) * PT
    '''Проходим по всем колонкам для установки размеров из списка'''
    for i in range(1, len(WidthList) + 1):
        col = tabWord.Cell(3, i).Range.Columns
        col.PreferredWidthType = 3
        col.PreferredWidth = WidthList[i-1] * PT
    '''Корректируем высоту строк (ячеек)'''
    l23 = tabWord.Cell(3, 3).Range.Rows
    l23.HeightRule = 2
    l23.Height = 0.5 * PT    
    '''Разрывам страницу'''
    breakPage(Doc)
    return tableNomer

def dataTable_7(response, wb, Doc, tableNomer):
    global tableName
    counterList = []
    counter = 6
    data = []
    temp = ''
    for i in response:
        # КО
        Description = i['KoItem']['Description']

        if Description != temp:
            data.append([Description] + [None] * 17)
            temp = Description
            counter += 1
            counterList.append(counter)
        
        col_3 = ''

        if i['Length'] == None and i['Width'] == None and i['Diameter'] != None:
            col_3 = str(round(float(i['Diameter']) / 1000, 3))

        if i['Length'] != None and i['Width'] == None and i['Diameter'] != None:
            col_3 = f"{round(float(i['Length']) / 1000, 3)} x {round(float(i['Diameter']) / 1000, 3)}"

        if i['Length'] == None and i['Width'] != None and i['Diameter'] != None:
            col_3 = f"{round(float(i['Width']) / 1000, 3)} x {round(float(i['Diameter']) / 1000, 3)}"

        if i['Length'] != None and i['Width'] != None and i['Diameter'] == None:
            col_3 = f"{round(float(i['Length']) / 1000, 3)} x {round(float(i['Width']) / 1000, 3)}"

        xxx = [
            # № ЭКСПЛИКАЦИИ ПО СХЕМЕ ГЕНПЛАНА
            i['GenplanNumber'],
            # НАИМЕНОВАНИЕ СООРУЖЕНИЙ
            i['Title'],
            # КОСНТРУКТИВНЫЕ ОСОБЕННОСТИ 
            get_dictionary_item_Title(i['CategoryLandLevelId']),
            # РАЗМЕР В ПЛАНЕ, М
            col_3.replace('.', ','),
            # ОБЩАЯ ВЫСОТА, М
            i['Height'],
            # КОЛИЧЕСТВО ЭТАЖЕЙ
            i['NumberFloorsBuilding'],
            # ОРИЕНТИРОВОЧНАЯ МАССА, Т
            i['Mass'],
            # ТИП (ПЛИТА, ЛЕНТОЧНЫЙ, СВАЙНЫЙ И ДР.)
            get_dictionary_item_Title(i['CategoryFoundationId']),
            # ПРЕДПОЛОГАЕМАЯ ГЛУБИНА ЗАЛОЖЕНИЯ, М
            i['DepthLaying'],
            # СЕЧЕНИЕ СВАЙ, ММ
            i['CrossSectionPiles'],
            # НА ОДНУ СВАЮ (КУСТ СВАЙ), КН (ТС)
            i['LoadPerPile'],
            # НА 1 ПОГОННЫЙ МЕТР ДЛИНЫ ЛЕНТОЧНОГО ФУНДАМЕНТА, КН/М2 (ТС/М2)
            i['LoadPerMeterRibbonFoundation'],
            # ПРЕДПОЛОГАЕМАЯ НА ГРУНТЫ, КН/М2 (ТС/М2)
            i['LoadSoil'],
            # ГЛУБИНА, М
            i['DepthBasement'],
            # НАЗНАЧЕНИЕ
            i['AppointmentBasement'],
            # ДИНАМИЧЕСКИХ НАГРУЗОК
            'Да' if i['AvailabilityDynamicLoads'] == True else 'Нет',
            # МОКРЫХ ТЕХНОЛОГИЧЕСКИХ ПРОЦЕССОВ
            'Да' if i['AvailabilityWetProcesses'] == True else 'Нет',
            # ДОПУСТИМЫЕ ВЕЛИЧИНЫ ДЕФОРМАЦИИ ОСОВАНИЯ, СМ
            '15'
        ]
        data.append(xxx)
        counter += 1
    # wb = Excel.ActiveWorkbook
    sheet = wb.Worksheets("Таблица (7)")
    sheet.Activate()
    '''Индексы первой ячейки (строка, стоблец) для вставки данных'''
    StartRow = 7
    StartCol = 1
    '''Индексы последней ячейки (строка, стоблец) для вставки данных'''
    EndRow = len(data) + StartRow - 1
    EndCol = len(data[0])
    '''Отправляем data в Excel'''
    cel = exportdata(data, sheet, StartRow, StartCol, EndRow, EndCol)
    '''Рисуем границы ячеек во всей таблице'''
    cel.Borders.Weight = 2
    sleep(1)

    '''Объединение ячеек по всей строке таблице с название КО'''
    for row in counterList:
        cel = RangeCells(sheet, row, StartCol, row, EndCol)
        sleep(0.1)
        cel.Merge()
        cel.Font.Bold = True
    '''Выделяем и копируем в буфер обмена всю таблицу из Excel'''
    tabEx = RangeCells(sheet, 3, StartCol, EndRow, EndCol)
    tabEx.EntireRow.AutoFit()
    sleep(0.5)
    tabEx.Copy()
    sleep(0.5)

    tableNomer += 1
    NazvanieTab =   f'''Таблица {tableNomer}
                    Техническая характеристика площадочных объектов для инженерно-геологических изысканий
                    '''
    tableName.append(' '.join([i.lstrip() for i in NazvanieTab.split('\n')[:2]]))
    tabWord = WordFormat(NazvanieTab, Doc)
    '''Корректируем высоту строк (ячеек)'''
    tabWord.Cell(1, 11).Height = 0.6 * PT
    tabWord.Cell(2, 11).Height = 0.6 * PT
    tabWord.Cell(3, 11).Height = 3.5 * PT
    '''Высота строки с номерами колонок'''
    tabWord.Cell(4, 11).Height = 0.5 * PT
    '''Повторять как заголовок на каждой странице'''
    tabWord.Cell(4, 11).Range.Rows.HeadingFormat = True
    tabWord.AutoFitBehavior(1)    # размер таблицы по содержимому
    tabWord.AutoFitBehavior(2)    # по ширине активного окна
    '''Корректируем ширину колонки'''
    col = tabWord.Cell(3, 12).Range.Columns
    col.PreferredWidthType = 3
    col.PreferredWidth = 1.7 * PT
    '''Корректируем высоту строк (ячеек)'''
    l23 = tabWord.Cell(4, 12).Range.Rows
    l23.HeightRule = 2
    l23.Height = 0.5 * PT    
    '''Разрывам страницу'''
    breakPage(Doc)
    return tableNomer


def dataTable_8(response1, response2, wb, Doc, tableNomer):
    global tableName
    counter = 0
    data = []

    tempList = []
    if response1 != []:
        for i in response1:
            Description = i['KoItem']['Description']
            if Description not in tempList:
                counter += 1
                tempList.append(Description)
                xxx = [
                    counter,
                    # ИСТОЧНИК ВОЗДЕЙСТВИЯ
                    i['KoItem']['Description'],
                    # РАСПОЛОЖЕНИЕ И ОБЪЕМЫ ИЗЬЯТИЯ ПРИРОДНЫХ РЕСУРСОВ (ЗЕМЕЛЬНЫХ, ВОДНЫХ, ЛЕСНЫХ И Т.Д.)
                    'Земельные участки в пределах отвода на период строительства и эксплуатацию',
                    # ШИРИНА ЗОНЫ ВОЗДЕЙСТВИЯ, М
                    response_dop['KitBuildImpactZoneWidth'],
                    # ГЛУБИНА ВОЗДЕЙСТВИЯ, М
                    response_dop['KitBuildImpactZoneDeep'],
                    # СОСТАВ ЗАГРЯЗНЯЮЩИХ ВЕЩЕСТВ ИЛИ ВИД ВОЗДЕЙСТВИЯ
                    response_dop['KitBuildCompositionPollutants'],
                    # ИНТЕНСИВНОСТЬ И ДЛИТЕЛЬНОСТЬ ВОЗДЕЙСТВИЯ
                    response_dop['KitBuildIntensityDurationExposure']
                ]
                data.append(xxx)
    if response2 != []:
        for i in response2:
            counter += 1
            xxx = [
                counter,
                # ИСТОЧНИК ВОЗДЕЙСТВИЯ
                i['Title'],
                # РАСПОЛОЖЕНИЕ И ОБЪЕМЫ ИЗЬЯТИЯ ПРИРОДНЫХ РЕСУРСОВ (ЗЕМЕЛЬНЫХ, ВОДНЫХ, ЛЕСНЫХ И Т.Д.)
                'Земельные участки в пределах отвода на период строительства и эксплуатацию',
                # ШИРИНА ЗОНЫ ВОЗДЕЙСТВИЯ, М
                response_dop['KitLineImpactZoneWidth'],
                # ГЛУБИНА ВОЗДЕЙСТВИЯ, М
                response_dop['KitLineImpactZoneDeep'],
                # СОСТАВ ЗАГРЯЗНЯЮЩИХ ВЕЩЕСТВ ИЛИ ВИД ВОЗДЕЙСТВИЯ
                response_dop['KitLineCompositionPollutants'],
                # ИНТЕНСИВНОСТЬ И ДЛИТЕЛЬНОСТЬ ВОЗДЕЙСТВИЯ
                response_dop['KitLineIntensityDurationExposure']
            ]
            data.append(xxx)
    ExcelFormat(wb, "Таблица (8)", data, StartRow=5, CopyStartRow=3)

    tableNomer += 1
    NazvanieTab =   f'''Таблица {tableNomer}
                    Характеристика существующих и проектируемых источников воздействия
                    '''
    tableName.append(' '.join([i.lstrip() for i in NazvanieTab.split('\n')[:2]]))
    tabWord = WordFormat(NazvanieTab, Doc)    
    '''Повторять как заголовок на каждой странице'''
    tabWord.Cell(2, 1).Range.Rows.HeadingFormat = True
    '''Корректируем высоту строк (ячеек)'''
    l23 = tabWord.Cell(2, 1).Range.Rows
    l23.HeightRule = 2
    l23.Height = 0.5 * PT
    return tableNomer


def dataTable_1(Excel, wb, Doc, NameObject):
    global tableName
    wb = Excel.ActiveWorkbook
    sheet = wb.Worksheets("Таблица (1)")
    sheet.Activate()
    data = []
    for i in range(len(tableName)):
        data.append([i + 1, tableName[i], None])

    '''Индексы первой ячейки (строка, стоблец) для вставки данных'''
    StartRow = 6
    StartCol = 1
    '''Индексы последней ячейки (строка, стоблец) для вставки данных'''
    EndRow = len(data) + StartRow - 1
    EndCol = len(data[0])
    '''Отправляем data в Excel'''
    cel = exportdata(data, sheet, StartRow, StartCol, EndRow, EndCol)
    '''Рисуем границы ячеек во всей таблице'''
    cel.Borders.Weight = 2
    sleep(1)

    '''Выделяем и копируем в буфер обмена всю таблицу из Excel'''
    tabEx = RangeCells(sheet, 4, StartCol, EndRow, EndCol)
    tabEx.EntireRow.AutoFit()
    sleep(0.5)
    tabEx.Copy()
    sleep(0.5)

    myRange = Doc.Paragraphs(6).Range
    myRange.PasteExcelTable(False, False, False)
    tabWord = Doc.Tables(1)
    tabWord.AutoFitBehavior(2)    # по ширине активного окна

    '''Корректируем высоту строк в таблице без объединенных ячеек'''
    tabWord.Rows.Height = 1.0 * PT
    tabWord.Rows(1).Height = 1.5 * PT
    tabWord.Rows(2).Height = 0.5 * PT

    '''Назначить значение пользовательскому свойсту'''
    UserProp = Doc.CustomDocumentProperties('NameObject')
    UserProp.Value = NameObject
    '''Обновляем поля свойств в основной области с текстом'''
    Doc.Fields.Update()


# def thread(my_func):
#     '''Обертка функции в потопк (декоратор)'''
#     def wrapper():
#         global thr
#         # threading.Thread(target=my_func, daemon=True).start()
#         thr = threading.Thread(target=my_func, daemon=True).start()
#     return wrapper


@decorTime
def GO(Id, TypeReport, NameFaileDoc):

    global dictionary_items, response_dop
    global tableName

    """Создаем COM объект Excel"""
    Excel = win32com.client.Dispatch("Excel.Application")
    Excel.Visible = 0
    # Excel.Visible = 1
    Excel.DisplayAlerts = 0
    
    """Создаем COM объект Word"""
    # Word = win32com.client.Dispatch("Word.Application")
    Word = win32com.client.gencache.EnsureDispatch("Word.Application")
    Word.Visible = 0
    # Word.Visible = 1
    Word.DisplayAlerts = 0

    '''Коды словарей'''    
    dictionary_items = resp(r"http://tnnc-pir-app/test-kits-buildings-api/dictionaries/dictionary-items")
    '''Дополнительные данные для формирования ТЗИИ'''
    response_dop = resp(f"http://tnnc-pir-app/test-kits-buildings-api/kits/kit-collection/{Id}")

    '''Наименование объекта'''
    NameObject = response_dop['Project']['Description']

    wb = Excel.Workbooks.Open(os.getcwd() + "\\ShablonTZII.xltx")
    Doc = Word.Documents.Add(os.getcwd() + "\\ShablonTZII.dotx")

    response1 = resp(f"http://tnnc-pir-app/test-kits-buildings-api/kits/kit-build-items/{Id}")
    response1 = [i for i in response1 if i['IsUsedTechnicalSpecificationEngineeringSurvey'] == True]
    response1 = sorted(response1, key = lambda i: int(i['GenplanNumber']))
    # response1.sort(key = lambda i: int(i['GenplanNumber']))

    response2 = resp(f'http://tnnc-pir-app/test-kits-buildings-api/kits/kit-line-items/{Id}')
    
    tableName = [
        'Таблица 1 Перечень Приложений к ТЗ на ИИ',
        f'Таблица 2 Лист согласования к ТЗ на выполнение ИИ по объекту \"{NameObject}\"'
    ]

    tableNomer = 3
    dataTable_3(response1, response2, wb, Doc)
    if response1 != []:
        tableNomer = dataTable_4(response1, wb, Doc, tableNomer)
    if response2 != []:
        tableNomer = dataTable_5(response2, wb, Doc, tableNomer)
        tableNomer = dataTable_6(response2, wb, Doc, tableNomer)
    if response1 != []:
        tableNomer = dataTable_7(response1, wb, Doc, tableNomer)
    tableNomer = dataTable_8(response1, response2, wb, Doc, tableNomer)
    dataTable_1(Excel, wb, Doc, NameObject)

    '''Установка интервала перед и после абзаца в тексте. Единица измерения интервала пт.'''
    myRange = Doc.Content
    myRange.ParagraphFormat.SpaceBefore = 0  #интервал перед
    myRange.ParagraphFormat.SpaceAfter = 0   #интервал после
    myRange.Font.Name = "Times New Roman"
    
    '''Закрываем экземпляр Excel'''
    Excel.Quit()
    '''Сохранить как'''
    Doc.SaveAs(FileName = os.getcwd() + f"\\reports\{NameFaileDoc}")  
    Word.Quit()
    '''====================================================================='''     


# os.system(r'call C:\vxvproj\tnnc-Kits_buildings_reports\Kits_buildings_reports.py 2271 TZnaII TZII.docx')


if __name__ == "__main__":
    Id, TypeReport, NameFaileDoc = '2286', 'TZnaII', 'TZII.docx'
    sys.exit(GO(Id, TypeReport, NameFaileDoc))

    
    

