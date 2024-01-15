from excell_funcs import get_col_data,change_col_data,get_vpr,new_file,filters
    

def firsr():
    new_file("calculate/5 Расчет параметров планирования.xlsx") 
    fdict = {"Сегмент":2}
    filtred_data = filters("input_folder/1. Перечень номенклатур",fdict,"Код позиции")
    change_col_data("calculate/5 Расчет параметров планирования","Код позиции",filtred_data)
    vpr = get_vpr("","Код позиции","",filtred_data)
    change_col_data("calculate/5 Расчет параметров планирования","",vpr)
    vpr1 = get_vpr("input_folder/1. Перечень номенклатур","Код позиции","Срок поставки",filtred_data)
    change_col_data("calculate/5 Расчет параметров планирования","Срок поставки",vpr1)
    vpr2 = get_vpr("input_folder/7. Коэффициенты риска по срокам поставки K_LT","Код позиции","Коэффициент К_LT",filtred_data)
    change_col_data("calculate/5 Расчет параметров планирования","Коэффициент К_LT",vpr2)
    vpr3 = get_vpr("calculate_folder/3 Расчет стандартного отклонения срока поставки","Код позиции" ,"Станд. откл. факт. срока поставки за последние 12 мес.",filtred_data)
    change_col_data("calculate_folder/5 Расчет параметров планирования","Стандартное отклонение срока поставки",vpr3)
    vpr4 = get_vpr("calculate_folder/4 Анализ потребления","Код позиции","Среднемесячный спрос за последние 12 месяцев",filtred_data) 
    change_col_data("calculate/5 Расчет параметров планирования","Среднемесячный спрос за последние 12 месяцев",vpr4)
    vpr5 = get_vpr("calculate_folder/4 Анализ потребления","Код позиции","Стандартное отклонение месячного спроса за последние 12 месяцев",filtred_data)
    change_col_data("calculate/5 Расчет параметров планирования","Стандартное отклонение месячного спроса за последние 12 месяцев",vpr5)
    vpr6 = get_vpr("calculate_folder/4 Анализ потребления","Код позиции","Минимальный месячный спрос за последние 12 месяцев",filtred_data)
    change_col_data("calculate/5 Расчет параметров планирования","Минимальный месячный спрос за последние 12 месяцев",vpr6)
    
