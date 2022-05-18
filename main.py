# импортируем функцию listdir из модуля ОС для просмотра содержимого папки
from datetime import datetime
from os import listdir

# импортируем модуль pandas, под псевдонимом pd, для работы с таблицей
import pandas as pd

from settings import CONFIG


# объявляем функцию которая будет находить все excel файлы в нашей папке проекта
def find_excel_files() -> tuple[str]:
    # объявляем список для найденных excel файлов
    excel_files = []
    # перебираем имена файлов в текущей папке
    for file in listdir('.'):
        # проверяем расширение файла
        if file.endswith('.xlsx'):
            # если расширение xlsx, то добавляем в список
            excel_files.append(file)
    # возвращаем кортеж найденных имен excel файлов
    return tuple(excel_files)


# функция по считыванию найденных файлов в Dataframe
def read_excel(excel_files: tuple[str]) -> tuple[pd.DataFrame]:
    # заводим список под dataframes
    frames = []
    # перебираем найденные excel файлы
    for excel in excel_files:
        try:
            # пытаемся открыть необходимую книгу в файле и добавить в список
            frames.append(pd.read_excel(excel, sheet_name=CONFIG['sheetname']))
        except:
            pass
    # возвращаем кортеж таблиц
    return tuple(frames)


# осуществляем выборку записей у которых есть дата в колонке 'ГУ "НА"'
def filter_excel(frames: tuple[pd.DataFrame]) -> pd.DataFrame:
    # заводим словарь-шаблон для будущего dataframe с результатом
    result = {
        CONFIG['registration_number']: [],
        CONFIG['date']: [],
        CONFIG['documents_count']: [],
        CONFIG['package']: [],
    }
    # заводим счетчик количества документов
    doc = 0
    # перебираем таблицы
    for frame in frames:
        try:
            # перебираем записи из таблицы
            for row in frame.loc:
                # проверяем наличие даты в колонке ГУ "НА"
                if row[CONFIG['archive']] is not pd.NaT:
                    # проверяет наличие даты в колонке ИМНС и проверяем разницу лет между текущим годом и годом в этой колонке на >= 4
                    if row[CONFIG['tax']] is not pd.NaT and 4 <= (datetime.now().year - row[CONFIG['tax']].year):
                        # print(row)
                        # подсчитываем количесво документы
                        doc += row[CONFIG['documents_count']]
                        # вносим данные в словарь
                        result[CONFIG['registration_number']].append(row[CONFIG['registration_number']])
                        result[CONFIG['date']].append(datetime.strftime(row[CONFIG['date']].date(), '%d.%m.%Y'))
                        result[CONFIG['documents_count']].append(row[CONFIG['documents_count']])
                        result[CONFIG['package']].append(row[CONFIG['package']])
                    # проверяем что в колонке 'Дата прекращения' стоит дата и разница лет между текущим годом и из этой колонки >= 11 лет
                    elif row[CONFIG['date']] is not pd.NaT and (datetime.now().year - row[CONFIG['date']].year) >= 11:
                        # print(row)
                        # подсчитываем количесво документы
                        doc += row[CONFIG['documents_count']]
                        # вносим данные в словарь
                        result[CONFIG['registration_number']].append(row[CONFIG['registration_number']])
                        result[CONFIG['date']].append(datetime.strftime(row[CONFIG['date']].date(), '%d.%m.%Y'))
                        result[CONFIG['documents_count']].append(row[CONFIG['documents_count']])
                        result[CONFIG['package']].append(row[CONFIG['package']])
        except:
            pass
    #     дописываем последнюю строку с общим количеством дел
    result[CONFIG['documents_count']].append(doc)
    result[CONFIG['date']].append('')
    result[CONFIG['registration_number']].append('')
    result[CONFIG['package']].append('')
    # словарь преобразуем в DataFrame
    result = pd.DataFrame(result)
    # возвращаем полученный dataframe
    return result


def write_result(frame: pd.DataFrame, file_name: str) -> None:
    frame.to_excel(f'{file_name}.xlsx', sheet_name='result')


def number_for_a_given_year(year: int, frames: tuple[pd.DataFrame]) -> pd.DataFrame:
    result = {
        CONFIG['registration_number']: [],
        CONFIG['date']: [],
        CONFIG['documents_count']: [],
        CONFIG['package']: []
    }
    doc = 0
    for frame in frames:
        try:
            for row in frame.loc:
                if row[CONFIG['date']] is not pd.NaT and row[CONFIG['date']].year == year:
                    print(row)
                    result[CONFIG['registration_number']].append(row[CONFIG['registration_number']])
                    result[CONFIG['date']].append(datetime.strftime(row[CONFIG['date']].date(), '%d.%m.%Y'))
                    result[CONFIG['documents_count']].append(row[CONFIG['documents_count']])
                    result[CONFIG['package']].append(row[CONFIG['package']])
                    doc += row[CONFIG['documents_count']]
        except:
            pass

    result[CONFIG['registration_number']].append('')
    result[CONFIG['date']].append('')
    result[CONFIG['documents_count']].append(doc)
    result[CONFIG['package']].append('')
    return pd.DataFrame(result)


def main():
    choice = input('Введите:\n\t1 - Сведения о выбывших\n\t2 - Выборка по году\n')
    while not choice.isdigit() and choice not in ['1', '2']:
        print('Неверный выбор, попробуйте еще раз!')
        choice = input('Введите:\n\t1 - Сведения о выбывших\n\t2 - Выборка по году\n')
    excel_files = find_excel_files()
    frames = read_excel(excel_files)
    if choice == '1':
        res = filter_excel(frames)
        write_result(res, 'result')
    elif choice == '2':
        year = input('Введите год прекращения: ')
        while not year.isdigit():
            print('Неверный выбор, попробуйте еще раз!')
            year = input('Введите год прекращения: ')
        year = int(year)
        res = number_for_a_given_year(year, frames)
        write_result(res, 'year')


if __name__ == '__main__':
    main()
