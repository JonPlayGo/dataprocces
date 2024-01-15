from excell_funcs import get_col_data, change_col_data, get_vpr, new_file, filters,get_one_vpr
import datetime


def one():
    new_file('calculate_folder/4 Анализ потребления.xlsx')
    filt = {
        "Плановые работы": "Да",
        "Актив в периметре": "Да"
    }
    pos_codes = filters("input_folder/4. Данные по списаниям", filt, "Код позиции", )
    pos_codes = list(dict.fromkeys(pos_codes))
    change_col_data("calculate_folder/4 Анализ потребления", "Код позиции", pos_codes)
    today = datetime.datetime.today()
    res2 = {
        -12: [],
        -11: [],
        -10: [],
        -9: [],
        -8: [],
        -7: [],
        -6: [],
        -5: [],
        -4: [],
        -3: [],
        -2: [],
        -1: []
    }

    trans_dates = get_vpr("input_folder/4. Данные по списаниям", "Код позиции", "Дата транзакции", pos_codes)
    deltas = {}
    numibs = 0
    for pos_code in pos_codes:
        if not pos_code in deltas:
            (trans_dates[numibs].year - today.year) * 12 + (trans_dates[numibs].month - today.month
    # for col in res2:
    #
    #         if (trans_dates[numibs].year - today.year) * 12 + (trans_dates[numibs].month - today.month) == col:
    #             amount = get_vpr("input_folder/4. Данные по списаниям","Код позиции","Количество",[pos_code])
    #             res2[col].append(amount)
    #         else:
    #             res2[col].append(0)
    #         numibs = numibs + 1

    res2 = {
        "Количество списаний (Текущий месяц – 12)": res2[-12],
        "Количество списаний (Текущий месяц – 11)": res2[-11],
        "Количество списаний (Текущий месяц – 10)": res2[-10],
        "Количество списаний (Текущий месяц – 9)": res2[-9],
        "Количество списаний (Текущий месяц – 8)": res2[-8],
        "Количество списаний (Текущий месяц – 7)": res2[-7],
        "Количество списаний (Текущий месяц – 6)": res2[-6],
        "Количество списаний (Текущий месяц – 5)": res2[-5],
        "Количество списаний (Текущий месяц – 4)": res2[-4],
        "Количество списаний (Текущий месяц – 3)": res2[-3],
        "Количество списаний (Текущий месяц – 2)": res2[-2],
        "Количество списаний (Текущий месяц – 1)": res2[-1]
    }
    print(res2)
    for r2 in res2:
        change_col_data("calculate_folder/4 Анализ потребления", r2, res2[r2])
