import pandas as pd
import glob
import csv


def split_string(string: str):
    index = 0
    while index < len(string):
        if string[index].isdigit():
            break
        index += 1
    return string[0:index], int(string[index:]) - 2


def scan_folder(path=''):
    filter = '20*_*.xlsx'
    return glob.glob(path + filter)


def save_to_csv(data, path):
    with open(path, 'a', newline='') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerows(data)


def main():
    svod = pd.read_excel('svod.xlsx', header=1, usecols='B:I')
    svod = svod.to_dict()
    for file in scan_folder():
        data = []
        print('Open file: ', file)
        year = file[0:4]
        territoria = file.split('_')[1]
        for element in zip(svod['Показатель'].values(), svod['Лист'].values(),
                           svod['Раздел'].values(), svod['Ячейка'].values(), svod['Тип поселения'].values()):
            if element[4].upper() == territoria.upper():
                column, row = split_string(element[3])
                df = pd.read_excel(file, usecols=column, sheet_name=element[1])
                value = df.iat[row, 0]
                result = [element[0], element[2], element[-1]] + [year, value]
                data.append(result)
        save_to_csv(data, 'result.csv')


if __name__ == '__main__':
    main()
