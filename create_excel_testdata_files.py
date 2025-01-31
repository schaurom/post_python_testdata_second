import pandas
import os
import time
import random
from faker import Faker
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook


def create_workbook(workbook_name, sheet_names):
    print(time.asctime(), 'create_workbook Start')
    # write_only to not get a default 'Sheet' in the file
    workbook = Workbook(write_only=True)
    # iterate through sheet names
    for sheet_name in sheet_names:
        # create new sheet with given name
        sheet = workbook.create_sheet(sheet_name)
        # create testdata for this sheet and fill it into a dataframe
        data = pandas.DataFrame(create_testdata())
        # print(data)
        # fill dataframe into the sheet
        for row in dataframe_to_rows(data, index=False):
            sheet.append(row)

    # save the workbook with given name
    workbook.save(workbook_name)
    print(time.asctime(), 'create_workbook Ende')

def create_testdata():
    print(time.asctime(), 'create_testdata Start')
    fake = Faker('de_AT')
    data_list = []
    for num in range(200):
        data_list.append({
            'Name': fake.last_name(),
            'Vorname': fake.first_name(),
            'Zufall': zufall(),
            'Titel': fake.prefix(),
            'Telefon': fake.phone_number(),
            'Strasse': fake.street_address(),
            'Postleitzahl': fake.postcode(),
            'Stadt': fake.city(),
            'Bank': fake.iban(),
            'Eintritt': fake.date_between().strftime('%d.%m.%Y')
        })
    print(time.asctime(), 'create_testdata Ende')
    return data_list


def zufall():
    randWert = random.choice(['true', 'false'])
    #print(randWert)
    return randWert


if __name__ == '__main__':
    zeitanfang = time.process_time()
    print(time.asctime())
    path_dir = r'C:\Users\schau\Desktop\Testdaten'
    create_workbook(os.sep.join([path_dir, 'Adressen.xlsx']), ['A'])
    # create_workbook(os.sep.join([path_dir, '1.xlsx']), ['A', 'B', 'C'])
    zeitende = time.process_time()
    print(time.asctime())
    print('Durchlaufdauer: ', (zeitende - zeitanfang), 's')
