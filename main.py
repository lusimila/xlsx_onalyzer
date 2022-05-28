# импортируем функцию listdir из модуля ОС
from os import listdir  # для того, чтобы посмотреть, какие эксель файлы есть в папке
from datetime import datetime  # импортируем из библиотеки datetime модуль datetime для работы с датой

# импортируем модуль pandas, под псевдонимом pd, для работы с таблицей
import pandas as pd

# импортируем CONFIG из файла settings. CONFIG является у нас словарем. Он формируется при считывании
# файла config.yaml и превращается в словарь
from settings import CONFIG


# объявляем функцию find_excel_files, которая будет находить все excel файлы в нашей
# папке проекта и возвращать кортеж найденных имен файлов
def find_excel_files() -> tuple[str]:  # функция будет возвращать кортеж строк
    # объявляем список для найденных имен excel файлов
    excel_files = []
    # перебираем имена файлов в текущей папке
    for file in listdir('.'):
        # проверяем расширение каждого файла

        if file.endswith('.xlsx'):
            # если расширение xlsx есть, то мы добавляем эти имена файлов в список
            excel_files.append(file)
    return tuple(excel_files)  # превращаем полученный список в кортеж и возвращаем его


# DataFrame - таблица в Pandas
# функция  read_excel считывает все найденные excel файлы и превращает их в таблицы (dataframe)
# выбирает нужный лист 'Книга учета дел', указанный в config.yaml, возвращает кортеж таблиц (dataframe)
def read_excel(excel_files: tuple[str]) -> tuple[pd.DataFrame]:
    # заводим список под таблицы dataframes
    frames = []
    # перебираем найденные excel файлы
    for excel in excel_files:
        try:
            # пытаемся открыть необходимую книгу в файле и добавить в список
            frames.append(pd.read_excel(excel, sheet_name=CONFIG['sheetname']))
        # если листа не окажется произойдет ошибка, except уловит эту ошибку для предотвращения падения программы
        except Exception:
            pass
    # далее возвращаем кортеж таблиц (dataframe)
    return tuple(frames)  # превращает список таблиц (dataframe) в кортеж таблиц (dataframe) и возвращает его


# В функции filter_excel осуществляем: фильтрацию таблиц, перебор  записей в таблице через цикл for,
# проверку наличия дат  в колонке 'ГУ "НА"
# Далее  проводятся:  фильтрация, перебор записей на наличие дат в колонке "ИМНС", сверка дат,
# подсчет количества документов и внесение результирующих сведений в словарь.
# Для создания полноценного dataframe полученный словарь дописываем пустыми строками (т.к. один из столбцов в таблице
# длинее других), далее словарь преобразуем в Dataframe,  и оператор return возвращает dataframe после выполнения работы
# функции.
def filter_excel(frames: tuple[pd.DataFrame]) -> pd.DataFrame:
    # заводим словарь-шаблон для будущего dataframe с результатом
    result = {
        CONFIG['registration_number']: [],
        CONFIG['documents_count']: [],
        CONFIG['date']: [],
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
                if row[CONFIG['archive']] is not pd.NaT:  # NaT - дата без даты из Pandas,
                    # Pandas воспринимает пустую ячейку в колонке дат как NaT, а не None
                    # проверяем наличие даты в колонке ИМНС и проверяем разницу лет между текущим годом
                    # и годом в этой колонке на >= 4
                    if row[CONFIG['tax']] is not pd.NaT and 4 <= (datetime.now().year - row[CONFIG['tax']].year):
                        # 1. проверяет наличие даты в колонке ИМНС
                        # 2. Если она есть, то проверяет разницу между текущим годом и годом в ячейке
                        # print(row)
                        # подсчитываем количество документов
                        doc += row[CONFIG['documents_count']]
                        # вносим данные в словарь
                        result[CONFIG['registration_number']].append(row[CONFIG['registration_number']])
                        result[CONFIG['date']].append(datetime.strftime(row[CONFIG['date']].date(), '%d.%m.%Y'))
                        result[CONFIG['documents_count']].append(row[CONFIG['documents_count']])
                        result[CONFIG['package']].append(row[CONFIG['package']])
                    # проверяем что в колонке 'Дата прекращения' стоит дата и разница лет между текущим годом
                    # и из этой колонки >= 11 лет
                    # 1. проверяет наличие даты в колонке Дата прекращения
                    # 2. Если она есть, то проверяет разницу между текущим годом и годом в ячейке
                    elif row[CONFIG['date']] is not pd.NaT and (datetime.now().year - row[CONFIG['date']].year) >= 11:
                        # print(row)
                        # подсчитываем количество документов
                        doc += row[CONFIG['documents_count']]
                        # вносим данные в словарь
                        result[CONFIG['registration_number']].append(row[CONFIG['registration_number']])
                        result[CONFIG['date']].append(datetime.strftime(row[CONFIG['date']].date(), '%d.%m.%Y'))
                        result[CONFIG['documents_count']].append(row[CONFIG['documents_count']])
                        result[CONFIG['package']].append(row[CONFIG['package']])
        except Exception:
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


#  функция number_for_a_given_year осуществляет выборку данных по заданному году, перебирает все найденные листы,
#  перебирает построчно и отбирает среди них строки с записями по интересующему году в колонке "Дата прекращения",
#  формирует словарь, на основании которого создает dataframe,
#  который функция write_result записывает уже в виде excel.
def filter_year(year: int, frames: tuple[pd.DataFrame]) -> pd.DataFrame:
    result = {
        CONFIG['registration_number']: [],
        CONFIG['documents_count']: [],
        CONFIG['date']: [],
        CONFIG['package']: []
    }
    doc = 0
    for frame in frames:
        try:
            for row in frame.loc:  # перебираем построчно таблицы
                # если в колонке "Дата прекращения" есть дата и она соответствует искомой, то добавляем в словарь
                if row[CONFIG['date']] is not pd.NaT and row[CONFIG['date']].year == year:
                    # print(row)
                    # заполнения словаря
                    result[CONFIG['registration_number']].append(row[CONFIG['registration_number']])
                    result[CONFIG['date']].append(datetime.strftime(row[CONFIG['date']].date(), '%d.%m.%Y'))
                    result[CONFIG['documents_count']].append(row[CONFIG['documents_count']])
                    result[CONFIG['package']].append(row[CONFIG['package']])
                    # подсчитываем количество документов
                    doc += row[CONFIG['documents_count']]
        except Exception:
            pass
    # заполняем последнюю строку с общим количеством документов
    result[CONFIG['registration_number']].append('')
    result[CONFIG['date']].append('')
    result[CONFIG['documents_count']].append(doc)
    result[CONFIG['package']].append('')
    # превращаем наш словарь в таблицу (dataframe) и возвращаем его
    return pd.DataFrame(result)


# функция write_result записывает полученный dataframe уже в форме excel
def write_result(frame: pd.DataFrame, file_name: str) -> None:
    frame.to_excel(f'{file_name}.xlsx', sheet_name='result')


# главная исполняемая функция
# 1. спрашивает у пользователя действие, которое он хочет произвести
# 2. запускает цепочку вызовов функций для выполнения поставленной задачи
def main():
    # опрос пользователя
    choice = input('Введите:\n\t1 - Сведения о выбывших\n\t2 - Выборка по году\nВвод: ')
    # проверка ввода пользователя на корректность
    while not choice.isdigit() and choice not in ['1', '2']:
        # если ввод не верный, предлагаем пользователю повторить ввод
        print('Неверный выбор, попробуйте еще раз!')
        choice = input('Введите:\n\t1 - Сведения о выбывших\n\t2 - Выборка по году\nВвод: ')
    # вызываем функцию поиска всех excel файлов
    excel_files = find_excel_files()
    # вызываем функцию преобразовывания файла в таблицу (dataframe)
    frames = read_excel(excel_files)
    # проверяем ввод пользователя
    if choice == '1':
        # если "1", то запускаем функцию фильтрации по колонкам ИМНС и "ГУ НА"
        res = filter_excel(frames)
        # вызываем функцию записи полученной таблицы в excel
        write_result(res, 'выбывшие')
    # проверка ввода
    elif choice == '2':
        # спрашиваем у пользователя год
        year = input('Введите год прекращения: ')
        # если год некорректный, то предлагаем повторный ввод
        while not year.isdigit():
            print('Неверный выбор, попробуйте еще раз!')
            year = input('Введите год прекращения: ')
        # меняем тип у года с строки на число
        year = int(year)
        # вызываем функцию фильтрации таблицы по году на оснвовании колонки "Дата прекращения"
        res = filter_year(year, frames)
        # записываем получившуюся таблицу (dataframe) в excel
        write_result(res, f'за_{year}')


if __name__ == '__main__':
    main()
