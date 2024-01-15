from excell_funcs import get_col_data,change_col_data,get_vpr
import datetime


def test_1():
    """
    Функция test_1 сравнивает данные из двух входных файлов и обновляет столбец в одном из файлов на
    основе результата сравнения.
    """
    adding_result = []
    thirth_input = get_col_data("input_folder/3. Перечень техники","Номер актива")
    forth_input = get_col_data("input_folder/4. Данные по списаниям","Номер актива")
    for fourth_data in forth_input:
        th_list = []
        for thirth_data in thirth_input:
           th_list.append(thirth_data)
        if fourth_data in th_list:
            adding_result.append("Да")
        elif not fourth_data in th_list:
            adding_result.append("Нет")
    change_col_data("input_folder/4. Данные по списаниям","Актив в периметре",adding_result)
    


def test_2():
    """
    Функция `test_2` извлекает данные из определенного столбца в файле, проверяет, содержит ли каждый из
    данных строку "ТО", добавляет "Да" или "Нет" в список на основе проверки, а затем обновляет другой
    столбец в том же файле. файл с добавленным списком.
    """
    adding_result = []
    forth_input = get_col_data("input_folder/4. Данные по списаниям","Опер с активами")
    for fourth_data in forth_input:
        # print(fourth_data)
        if "ТО" in str(fourth_data) :
            adding_result.append("Да")
        elif not "ТО" in fourth_data :
            adding_result.append("Нет")
    change_col_data("input_folder/4. Данные по списаниям","Плановые работы",adding_result)
    

def test_3():
    """
    Функция test_3 извлекает данные столбца из файла, выполняет сегментацию на основе определенного
    столбца и обновляет данные столбца с учетом результатов сегментации.
    """
    col_data = get_col_data("input_folder/1. Перечень номенклатур","Код позиции")   
    vpr = get_vpr("input_folder/2. Сегментация","Код позиции","Сегмент",col_data)
    change_col_data("input_folder/1. Перечень номенклатур","Сегмент",vpr)


def test_4():
    """
    Функция test_4 вычисляет среднее значение определенного столбца набора данных за последние 6 месяцев
    и сохраняет результат в новом столбце.
    """
    a1 = get_col_data("input_folder/6. Данные по выработке техники","Фактический показатель (Текущий месяц – 5)")
    a2 = get_col_data("input_folder/6. Данные по выработке техники","Фактический показатель (Текущий месяц – 4)")
    a3 = get_col_data("input_folder/6. Данные по выработке техники","Фактический показатель (Текущий месяц – 3)")
    a4 = get_col_data("input_folder/6. Данные по выработке техники","Фактический показатель (Текущий месяц – 2)")
    a5 = get_col_data("input_folder/6. Данные по выработке техники","Фактический показатель (Текущий месяц – 1)")
    a6 = get_col_data("input_folder/6. Данные по выработке техники","Фактический показатель (Текущий месяц)")
    res =[]
    i = 0
    for l in a1:
        data = a1[i]+a2[i]+a3[i]+a4[i]+a5[i]+a6[i]
        res.append(int(data/6))
        i= i+1
    change_col_data("input_folder/6. Данные по выработке техники","Средняя фактическая наработка за последние 6 месяцев",res)

def test_5():
    """
    Функция test_5 извлекает данные из определенных столбцов файла, вычисляет среднее значение этих
    столбцов, а затем обновляет другой столбец в том же файле, используя вычисленное среднее значение.
    """
    a1 = get_col_data("input_folder/6. Данные по выработке техники","Плановый показатель (Текущий месяц + 1)")
    a2 = get_col_data("input_folder/6. Данные по выработке техники","Плановый показатель (Текущий месяц + 2)")
    a3 = get_col_data("input_folder/6. Данные по выработке техники","Плановый показатель (Текущий месяц + 3)")
    a4 = get_col_data("input_folder/6. Данные по выработке техники","Плановый показатель (Текущий месяц + 4)")
    a5 = get_col_data("input_folder/6. Данные по выработке техники","Плановый показатель (Текущий месяц + 5)")
    a6 = get_col_data("input_folder/6. Данные по выработке техники","Плановый показатель (Текущий месяц + 6)")
    res =[]
    
    i = 0
    for l in a1:
        data = a1[i]+a2[i]+a3[i]+a4[i]+a5[i]+a6[i]
        res.append(int(data/6))
        i= i+1
    change_col_data("input_folder/6. Данные по выработке техники","Средняя плановая наработка на следующие 6 месяцев",res)

def test_6():
    """
    Функция «test_6» рассчитывает коэффициент К1 на основе среднего фактического и планового
    использования оборудования за период 6 месяцев.
    """
    res = []
    first = get_col_data("input_folder/6. Данные по выработке техники","Средняя фактическая наработка за последние 6 месяцев")
    second = get_col_data("input_folder/6. Данные по выработке техники","Средняя плановая наработка на следующие 6 месяцев")
    g = 0
    for i in first:
        if i == 0:
           res.append(1) 
        else:
            res.append(second[g]/i)
        g = g+1
        
    change_col_data("input_folder/6. Данные по выработке техники","Коэффициент К1",res)

def test_7():
    """
    Функция test_7 извлекает данные из двух столбцов файла, объединяет значения в одну строку, а затем
    обновляет другой столбец в том же файле объединенными значениями.
    """
    res =[]
    g = 0
    data1 = get_col_data("input_folder/5. Данные по закупкам","Дата отгрузки") 
    data2 = get_col_data("input_folder/5. Данные по закупкам","Дата подписания") 
    for i in data1:
        res.append((datetime.datetime.strptime(data1[g],"%d.%m.%Y") - data2[g]).days)
        g = g+1
    change_col_data("input_folder/5. Данные по закупкам","Фактический срок поставки",res)
