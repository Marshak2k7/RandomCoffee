import pandas as pd  # для чтения, обработки и выгрузки данных

import random  # для выбора случайного кандидата
import datetime as dt  # для вывода даты

from collections import defaultdict  # для создания словаря с работниками

df = pd.read_excel('RandomCoffee.xlsx', sheet_name=None, header=None, index_col=None)  # чтение excel файла

# формирование одной колонки и удаление пустых значений
workers = pd.Series(df['Полный список участвующих'].values.ravel()).dropna().apply(lambda x: x.strip())

workers = workers.tolist()  # формирование списка из колонки

n = len(workers)  # количество работников

m = n // 2  # количество возможных пар

weeks = int((n * (n - 1) / 2) // (m + 2) - (n // 10 + 2))  # количество недель


# генерация пар
def generate_pair(workers, all_people):
    workers_set = set(workers)  # множество из работников
    a = workers[0]  # первый работник

    candidates = workers_set.difference(all_people[a]).difference({a})  # кандидаты для первого работника
    b = random.choice(list(candidates))  # случайный выбор пары

    workers.remove(a)  # удаление работника из общего списка
    workers.remove(b)  # удаление собеседника из общего списка
    return [a, b], workers


# генерация троек
def generate_triple(workers, all_people):
    workers_shuffled = random.sample(workers, len(workers))  # перемешивание работников
    group, workers = generate_pair(workers_shuffled, all_people)  # вызов функции geberate_pair и запись ее результатов
    a, b = group  # распаковывание пары

    workers_set = set(workers)  # создание множества из работников
    candidates_for_third = workers_set.difference(all_people[a]).difference(all_people[b]).difference(
        {a, b})  # определение возможных кандидатов для тройки
    c = random.choice(list(candidates_for_third))  # случайный выбор третьего сотрудника

    workers.remove(c)  # удаление сотрудника из общего списока
    return [a, b, c], workers


def generate_groups(workers_sorted, all_people):
    workers = list(workers_sorted.keys())  # отсортированные работники
    groups = list()

    # при нечетном количестве работников вызываем generate_triple
    if len(workers) % 2 == 1:
        group, workers = generate_triple(workers, all_people)
        groups.append(group)

    while True:
        # когда список работников опустеет, остановим цикл
        if len(workers) == 0:
            break
        # при четном количестве работников вызываем generate_pair
        group, workers = generate_pair(workers, all_people)
        groups.append(group)
    return groups


def get_all_people(other_sheets):
    all_people = defaultdict(set)  # словарь, где в качестве ключей будут участники RC
    for sheet in other_sheets:
        for i, row in sheet[1].iterrows():
            people = row.dropna().apply(lambda x: x.strip()).values
            if len(people) == 2:
                b, c = people
                all_people[b].add(c)
                all_people[c].add(b)
            else:
                b, c, d = people
                all_people[b] = all_people[b].union({c, d})
                all_people[c] = all_people[c].union({b, d})
                all_people[d] = all_people[d].union({b, c})
    for i in [x for x in workers if x not in all_people]:
        all_people[i] = set()

    return all_people


def get_final_groups(other_sheets):
    all_people = get_all_people(other_sheets)

    workers_sorted = {k: v for k, v in sorted(all_people.items(), key=lambda item: len(item[1]), reverse=True) if
                      k in workers}
    repeats = 0
    while repeats < 100:
        try:
            groups = generate_groups(workers_sorted, all_people)
            break
        except:
            repeats += 1
    if repeats == 100:
        return get_final_groups(other_sheets[1:])
    else:
        df = pd.DataFrame(groups)
        return df


if len(df) == 1:

    def make_pairs(workers):
        groups = []  # список случайных пар
        workers = workers.copy()  # копирование списка работников
        random.shuffle(workers)  # перемешевание списка работников

        # формирование списка случайных пар
        for i in range(0, len(workers), 2):
            groups.append(workers[i:i + 2])

            # отработка нечетного количества участников
        if len(workers) % 2 == 1:
            groups[-2].append(*groups[-1])
            groups.pop()  # удаление последнего элемента
        return groups


    df = pd.DataFrame(make_pairs(workers))  # формирование Pandas таблицы

elif 1 < len(df) <= weeks:
    other_sheets = list(df.items())[1:]
    df = get_final_groups(other_sheets)  # формирование Pandas таблицы

elif len(df) - 1 == weeks:
    other_sheets = list(df.items())[2:]
    df = get_final_groups(other_sheets)

else:
    other_sheets = list(df.items())[-weeks:]
    df = get_final_groups(other_sheets)

today_date = dt.date.today().strftime('%d.%m.%Y')  # текущая дата

# запись в excel
with pd.ExcelWriter('RandomCoffee.xlsx', mode='a') as writer:
    df.to_excel(writer, today_date, index=False, header=False)
