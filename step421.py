from excell_funcs import get_col_data, change_col_data, get_vpr, new_file


def first():
    new_file("calculate_folder/1 Коэффициент спроса на группу.xlsx")
    actives_group = get_col_data("input_folder/6. Данные по выработке техники", "Группа актива")
    res = []
    for active in actives_group:
        if not active in res:
            res.append(active)
    change_col_data("calculate_folder/1 Коэффициент спроса на группу", "Группа техники", res)


cols = ['Фактический показатель (Текущий месяц – 5)',
        "Фактический показатель (Текущий месяц – 4)",
        "Фактический показатель (Текущий месяц – 3)",
        "Фактический показатель (Текущий месяц – 2)",
        "Фактический показатель (Текущий месяц – 1)",
        "Фактический показатель (Текущий месяц)",
        "Плановый показатель (Текущий месяц + 1)",
        "Плановый показатель (Текущий месяц + 2)",
        "Плановый показатель (Текущий месяц + 3)",
        "Плановый показатель (Текущий месяц + 4)",
        "Плановый показатель (Текущий месяц + 5)",
        "Плановый показатель (Текущий месяц + 6)"
        ]


def second():
    for col in cols:
        all_sum = {}
        no_sort = get_col_data("input_folder/6. Данные по выработке техники", col)
        his_vpr = get_vpr("input_folder/6. Данные по выработке техники", col, "Группа актива", no_sort)
        for hvpr in his_vpr:
            for nsrt in no_sort:
                if not hvpr in all_sum:
                    all_sum[hvpr] = nsrt
                else:
                    all_sum[hvpr] += nsrt
        res = []
        for name, numb in all_sum.items():
            res.append(numb)
        change_col_data("calculate_folder/1 Коэффициент спроса на группу", f"Фактические и плановые показатели ({col})",
                        res)


def thirth():
    res = []
    a6 = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                      "Фактические и плановые показатели (Фактический показатель (Текущий месяц – 5))")
    a5 = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                      "Фактические и плановые показатели (Фактический показатель (Текущий месяц – 4))")
    a4 = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                      "Фактические и плановые показатели (Фактический показатель (Текущий месяц – 3))")
    a3 = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                      "Фактические и плановые показатели (Фактический показатель (Текущий месяц – 2))")
    a2 = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                      "Фактические и плановые показатели (Фактический показатель (Текущий месяц – 1))")
    a1 = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                      "Фактические и плановые показатели (Фактический показатель (Текущий месяц))")
    for i in range(0, len(a1)):
        data = a1[i] + a2[i] + a3[i] + a4[i] + a5[i] + a6[i]
        res.append(int(data / 6))
    change_col_data("calculate_folder/1 Коэффициент спроса на группу",
                    "Средняя фактическая наработка по группе за последние 6 месяцев", res)


def fourth():
    res = []
    a6 = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                      "Фактические и плановые показатели (Плановый показатель (Текущий месяц + 1))")
    a5 = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                      "Фактические и плановые показатели (Плановый показатель (Текущий месяц + 2))")
    a4 = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                      "Фактические и плановые показатели (Плановый показатель (Текущий месяц + 3))")
    a3 = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                      "Фактические и плановые показатели (Плановый показатель (Текущий месяц + 4))")
    a2 = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                      "Фактические и плановые показатели (Плановый показатель (Текущий месяц + 5))")
    a1 = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                      "Фактические и плановые показатели (Плановый показатель (Текущий месяц + 6))")
    for i in range(0, len(a1)):
        data = a1[i] + a2[i] + a3[i] + a4[i] + a5[i] + a6[i]
        res.append(int(data / 6))
    change_col_data("calculate_folder/1 Коэффициент спроса на группу",
                    "Средняя плановая наработка по группе на следующие 6 месяцев", res)


def fivth():
    res = []
    first = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                         "Средняя фактическая наработка по группе за последние 6 месяцев")
    second = get_col_data("calculate_folder/1 Коэффициент спроса на группу",
                          "Средняя плановая наработка по группе на следующие 6 месяцев")
    g = 0
    for i in first:
        if i == 0:
            res.append(1)
        else:
            res.append(second[g] / first[g])
        g = g + 1
    change_col_data("calculate_folder/1 Коэффициент спроса на группу", "Коэффициент К1_группа", res)
