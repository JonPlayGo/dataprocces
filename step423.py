from excell_funcs import change_col_data, new_file, filters, get_col_data, get_vpr
import datetime
import numpy as np

numbers = [1, 2, 3, 4, 5]

std = np.std(numbers)

print(std)


def first():
    new_file("calculate_folder/3 Расчет стандартного отклонения срока поставки.xlsx")
    fdict = {
        "Код классификатора 1 уровень": ["007.07", "007.08"]
    }

    change_col_data("calculate_folder/3 Расчет стандартного отклонения срока поставки", "Код позиции",
                    filters("input_folder/1. Перечень номенклатур", fdict, "Код позиции"))
    change_col_data("calculate_folder/3 Расчет стандартного отклонения срока поставки", "Код классификатора 1 уровень",
                    filters("input_folder/1. Перечень номенклатур", fdict, "Код классификатора 1 уровень"))
    change_col_data("calculate_folder/3 Расчет стандартного отклонения срока поставки", "Код классификатора 2 уровень",
                    filters("input_folder/1. Перечень номенклатур", fdict, "Код классификатора 2 уровень"))
    change_col_data("calculate_folder/3 Расчет стандартного отклонения срока поставки", "Код классификатора 3 уровень",
                    filters("input_folder/1. Перечень номенклатур", fdict, "Код классификатора 3 уровень"))


def second():
    pos_codes = get_col_data(
        "calculate_folder/3 Расчет стандартного отклонения срока поставки",
        "Код позиции"
    )
    pod_date = get_vpr(
        "input_folder/5. Данные по закупкам",
        "Позиция",
        "Дата подписания",
        pos_codes
    )
    otg_date = get_vpr(
        "input_folder/5. Данные по закупкам",
        "Позиция",
        "Дата отгрузки",
        pos_codes
    )
    res = {}
    pos_numb = 0
    for pos_code in pos_codes:
        if not pos_code in res:
            res[pos_code] = 0
        if pos_numb == len(otg_date):
            continue
        if not pod_date[pos_numb] is None or otg_date[pos_numb] is None:
            if datetime.datetime.strptime(otg_date[pos_numb], "%d.%m.%Y") >= pod_date[pos_numb]:
                res[pos_code] = res[pos_code] + 1

        pos_numb = pos_numb + 1
    result = []
    for i in res:
        result.append(res[i])
    change_col_data("calculate_folder/3 Расчет стандартного отклонения срока поставки",
                    "Количество поставок с известными датами",
                    result)


def thirth():
    pos_codes = get_col_data(
        "calculate_folder/3 Расчет стандартного отклонения срока поставки",
        "Код позиции"
    )
    pod_date = get_vpr(
        "input_folder/5. Данные по закупкам",
        "Позиция",
        "Дата подписания",
        pos_codes
    )
    otg_date = get_vpr(
        "input_folder/5. Данные по закупкам",
        "Позиция",
        "Дата отгрузки",
        pos_codes
    )
    res = {}
    pos_numb = 0
    for pos_code in pos_codes:
        if not pos_code in res:
            res[pos_code] = 0
        if pos_numb == len(otg_date):
            continue
        if not pod_date[pos_numb] is None or otg_date[pos_numb] is None:
            if datetime.datetime.strptime(otg_date[pos_numb], "%d.%m.%Y") >= pod_date[pos_numb]:
                res[pos_code] = res[pos_code] + 1

        pos_numb = pos_numb + 1
    result = []
    for i in res:
        if res[i] == 0:
            result.append("н/д")
        else:
            fact_data = get_vpr("input_folder/5. Данные по закупкам",
                                "Позиция",
                                "Фактический срок поставки",
                                [i])
            result.append(np.std(fact_data))
    change_col_data("calculate_folder/3 Расчет стандартного отклонения срока поставки",
                    "Станд. откл. факт. срока поставки за последние 12 мес.",
                    result)


