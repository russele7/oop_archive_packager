from random import randint
from faker import Faker
import xlsxwriter
import csv
from tqdm import tqdm
import numpy as np

N_COLS_MIN = 10
N_ROWS_MIN = 5000
N_ROWS_MAX = 10000
COLS_REPORT_1 = f'Введите число столбцов не меньше {N_COLS_MIN}:'
COLS_REPORT_2 = (
    f'Число столбцов должно быть >= {N_COLS_MIN}. Повторите попытку')
ROWS_REPORT_1 = ('Введите число строк '
                 f'в диапазоне от {N_ROWS_MIN} до {N_ROWS_MAX}:')
ROWS_REPORT_2 = (
    'Число строк должно быть в диапазоне от '
    f'{N_ROWS_MIN} до {N_ROWS_MAX}. Повторите попытку')
REPORT_3 = 'Вы ввели не число. Повторите попытку'


def data_size_from_user(
        min_value=0, max_value=np.inf,
        report_1='', report_2='', report_3=''):
    while True:
        try:
            result = int(
                input(report_1))
            if min_value <= result <= max_value:
                break
            else:
                print(report_2)
        except ValueError:
            print(report_3)
    return result


class ArchivePackager:
    ''''''
    def __init__(self) -> None:
        pass


class DataGenerator:
    '''DataGenerator'''
    def __init__(self, n_cols, n_rows) -> None:
        self.n_cols = n_cols
        self.n_rows = n_rows
    
    def generate_data(self):
        # data generation
        fake = Faker('ru_RU')

        # generiruem kaskadom
        file_csv = open('csv_data.csv', 'w', encoding='utf-8')
        excel_wb = xlsxwriter.Workbook('excel_data.xlsx')
        excel_ws = excel_wb.add_worksheet()

        for i in tqdm(range(self.n_rows)):
            row = [fake.job() for j in range(self.n_cols)]
            # save to csv
            write = csv.writer(file_csv)
            write.writerow(row)

            # save to excel
            for j, item in enumerate(row):
                excel_ws.write(i, j, item)
        excel_wb.close()


if __name__ == '__main__':
    print('Приложение "Генератор данных запущено"')
    n_cols = data_size_from_user(
        N_COLS_MIN, np.inf,
        COLS_REPORT_1, COLS_REPORT_2, REPORT_3)

    while True:
        n_rows_mark = input('Число строк задаст пользователь(y/n)?')
        if n_rows_mark.lower() not in ('y', 'n'):
            print('Недопустимый ответ пользователя. Повторите попытку')
        else:
            break

    if n_rows_mark.lower() == 'y':
        n_rows = data_size_from_user(
            N_ROWS_MIN, N_ROWS_MAX,
            ROWS_REPORT_1, ROWS_REPORT_2, REPORT_3)
    else:
        n_rows = randint(N_ROWS_MIN, N_ROWS_MAX)

    generator = DataGenerator(n_cols, n_rows)
    generator.generate_data()
