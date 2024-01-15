from excell_funcs import get_col_data, change_col_data, get_vpr, new_file, filters
import datetime


def two1():
    """
    The function `two1` updates the data in the "2 Коэффициент спроса.xlsx" file based on data from the
    "4. Данные по списаниям" file.
    """
    new_file("calculate_folder/2 Коэффициент спроса.xlsx")
    current_date = datetime.datetime.today()
    future_date = current_date + datetime.timedelta(days=365)
    month = future_date.month
    day = 1
    rounded_date = datetime.datetime(current_date.year, month, day)
    fdict = {
        "Дата транзакции": rounded_date,
        "Плановые работы": "Нет",
        "Актив в периметре": "Да"
    }
    test_list = []
    for i in filters("input_folder/4. Данные по списаниям", fdict, "Код позиции"):
        if not i in test_list:
            test_list.append(i)
    change_col_data("calculate_folder/2 Коэффициент спроса", "Код позиции", test_list)
    test_list2 = []
    get_list2 = get_vpr("input_folder/4. Данные по списаниям", "Код позиции", "Группа актива", test_list)
    for i in range(0, len(test_list)):
        test_list2.append(get_list2[i])

    change_col_data("calculate_folder/2 Коэффициент спроса", "Группа актива", test_list2)


def two2():
    pos_codes = get_col_data("calculate_folder/2 Коэффициент спроса", "Код позиции")
    res = {}
    for pos_code in pos_codes:
        pos_res = get_vpr("input_folder/4. Данные по списаниям", "Код позиции", "Количество", [pos_code])
        res[pos_code] = pos_res
    result = []
    for pos_name in res:
        res2 = 0
        for adding_amount in res[pos_name]:
            res2 = res2 + adding_amount
        result.append(res2)
    change_col_data("calculate_folder/2 Коэффициент спроса", "Количество", result)


def two3():
    """
    Функция «two3» извлекает данные из двух разных файлов, выполняет вычисления с использованием
    полученных данных и обновляет определенный столбец в другом файле вычисленными значениями.
    """
    with_data = get_col_data("calculate_folder/2 Коэффициент спроса", "Группа актива")
    vpr = get_vpr("calculate_folder/1 Коэффициент спроса на группу", "Группа техники", "Коэффициент К1_группа",
                  with_data)
    change_col_data("calculate_folder/2 Коэффициент спроса", "Коэффициент К1_группа", vpr)


def two4():
    """
    Функция «two4» вычисляет произведение двух столбцов в определенном файле и сохраняет результат в
    другом столбце того же файла.
    """
    a = get_col_data("calculate_folder/2 Коэффициент спроса", "Количество")
    b = get_col_data("calculate_folder/2 Коэффициент спроса", "Коэффициент К1_группа")
    res = []
    for i in range(0, len(a)):
        res.append(a[i] * b[i])
    change_col_data("calculate_folder/2 Коэффициент спроса", "Расчетное поле для веса", res)


def two5():
    pos_codes = get_col_data("calculate_folder/2 Коэффициент спроса", "Код позиции")
    pos_codes = list(dict.fromkeys(pos_codes))
    res = []
    for pos_code in pos_codes:
        amounts = get_vpr("calculate_folder/2 Коэффициент спроса", "Код позиции", "Количество", [pos_code])
        sum_amount = 0
        for amount in amounts:
            sum_amount = sum_amount + amount
        if sum_amount == 0:
            res.append(1)
        else:
            calculate_cells = get_vpr("calculate_folder/2 Коэффициент спроса", "Код позиции", "Расчетное поле для веса",
                                      [pos_code])
            res.append(calculate_cells[0] / sum_amount)

    print(res)
    change_col_data("calculate_folder/2 Коэффициент спроса", "Коэффициент К_спрос", res)
