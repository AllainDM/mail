from datetime import datetime, timedelta
import os

import xlrd3
import xlwt

import filter
import for_api

# Версия для xlrd
# workbook = xlrd.open_workbook(f'files/{filename}')
# worksheet = workbook.sheet_by_index(0)

# Версия для openpyxl
# workbook_all = openpyxl.load_workbook(f'files/{filename}')
# sheet_all = workbook_all.active
# for row in sheet_all.iter_rows(values_only=True):

def econtracts(date, to, list_filenames):
    # list_filenames это список полученный с почты(только что обработанный)
    # Получим список всех файлов с папки
    list_files = os.listdir(f"files")
    # Обработка даты для поиска в названии файла
    # Тестовый варинт за минус 1 день
    date_now = datetime.now()
    date_ago = date_now - timedelta(3)  # здесь мы выставляем минус день
    date = date_ago.strftime("%Y-%m-%d")
    print(date)

    list_all_files = []
    print(f"Обрабатываем список файлов: {list_files}")
    if type(list_files) == list:
        print("Подтверждено что список файлов действительно список.")
        for f in list_files:
            # print(f[0:10])
            # print(f[15])
            if f[0:10] == date and f[15] == 'e':
                print(f"Найден файл с необходимым названием: {f}.")
                readed_file = read_exel(f, to)
                list_all_files.append(readed_file)
    print(f"list_all_files {list_all_files}")

    save_to_exel(readed_file, to)


# Чтение файла ексель
def read_exel(filename, to):
    list_to = []
    list_hz = [["" for _ in range(10)],
               ["", "Не прошедние по фильтрам:", "", "", "", "", "", "", "", ""]]
    # Версия для xlrd3
    wb = xlrd3.open_workbook(f'files/{filename}')
    sheet = wb.sheet_by_index(0)
    # Старт со второй строчки
    for row in range(1, sheet.nrows):
        list_one = []

        # 1 Бренд, получим с помощью API
        try: list_one.append(for_api.search_brand(int(sheet.cell_value(row, 1))))
        except ValueError: list_one.append(" ")

        # 2 Дата
        list_one.append(sheet.cell_value(row, 0))

        # 3 Лицевой счет
        try: list_one.append(int(sheet.cell_value(row, 1)))
        except ValueError: list_one.append(sheet.cell_value(row, 1))

        # 4 Номер заявки
        try: list_one.append(int(sheet.cell_value(row, 2)))
        except ValueError: list_one.append(sheet.cell_value(row, 2))

        # 5 Улица
        street = filter.filter_street(sheet.cell_value(row, 3))
        list_one.append(street)

        # 6 Дом
        try: list_one.append(int(sheet.cell_value(row, 4)))
        except ValueError: list_one.append(sheet.cell_value(row, 4))

        # 7 Квартира
        try: list_one.append(int(sheet.cell_value(row, 5)))
        except ValueError: list_one.append(sheet.cell_value(row, 5))

        # 8 Мастер
        master = sheet.cell_value(row, 6)
        if master in filter.filter_master_no_to:
            continue
        else:
            list_one.append(sheet.cell_value(row, 6))

        # Проверим столбец оборудования до выставления типа договора
        router = filter.filter_router(sheet.cell_value(row, 8))

        # 9 Тип договора
        if router == "Услуга":
            list_one.append("Услуга")
        else:
            list_one.append(sheet.cell_value(row, 7))

        # 10 Оборудование
        if router == "Услуга":
            list_one.append("")
        else:
            list_one.append(router)

        # Итог
        if to == "north" and master in filter.filter_master_north:
            list_to.append(list_one)
            # else:
            #     list_hz.append(list_one)
        elif to == "south" and master in filter.filter_master_south:
            list_to.append(list_one)
            # else:
            #     list_hz.append(list_one)
        elif to == "west" and master in filter.filter_master_west:
            list_to.append(list_one)
            # else:
            #     list_hz.append(list_one)
        elif to == "east" and master in filter.filter_master_east:
            list_to.append(list_one)
        elif (master not in filter.filter_master_north and
              master not in filter.filter_master_south and
              master not in filter.filter_master_west and
              master not in filter.filter_master_east):
            list_hz.append(list_one)

    # Сложим список мастеров выбранного ТО и список неопределенных мастеров
    # return list_to
    return list_to + list_hz




def save_to_exel(list_to_exel, to):
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Электронные акты")

    last_n = 0  # Последняя строка, для вывода оповещения

    ws.write(last_n, 1, "Электронные договора.")

    for n, v in enumerate(list_to_exel):
        ws.write(n+2, 0, v[0])  # Бренд
        ws.write(n+2, 1, v[1])  # Дата
        ws.write(n+2, 2, v[2])  # Лицевой счет
        ws.write(n+2, 3, v[3])  # Номер заявки
        ws.write(n+2, 4, v[4])  # Улица
        ws.write(n+2, 5, v[5])  # Дом
        ws.write(n+2, 6, v[6])  # Квартира
        ws.write(n+2, 7, v[7])  # Мастер
        ws.write(n+2, 9, v[8])  # Тип договора
        ws.write(n+2, 10, v[9])  # Оборудование
        last_n = n

    ws.write(last_n + 4, 1, "Больше электронных договоров нет.")


    date_now = datetime.now()
    date_now_year = date_now.strftime("%d.%m.%Y %H:%M")

    wb.save(f'result/{to}_{date_now_year}.xlsx')
