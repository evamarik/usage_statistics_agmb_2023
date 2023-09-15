import pandas as pd
import numpy as np
import openpyxl

'''Excel file columns (updated file)
    'ISBN': column B -> in python 1
    'Titel': column C -> in python 2
    'Ausleihzähler gesamt': column J -> in python 9
    'Ausleihzähler lf. Jahr': column K -> in python 10
    'Ausleihzähler Vorjahr': column L -> in python 11
    'Ausleihzähler Vorvorjahr': column M -> in python 12
    'Ausleihzähler aller Jahre vor Vorvorjahr': column N -> in python 13
    '''


# file_name should be a string containing only the original (excel) files name (must be exactly the same), without .xlsx
def prepare_excel_file(file_name):
    # read Excel file
    input_file_name = file_name + '.xlsx'
    input_file = pd.read_excel(input_file_name)

    length = input_file.shape[0]
    length_list = np.arange(length)

    # check if every book has an ISBN (not really necessary)
    for i in length_list:
        if(input_file.iloc[i, 16] == None):
            print('Error: keine ISBN vorhanden', i)

    # create new table with filtered information and combined rows (combining books with same ISBN)
    input_file_rearranged = input_file.groupby(['ISBN'], sort=False).agg({
        'Signatur': 'first',
        'ISBN': 'first',
        'Titel': 'first',
        'Titelzusatz': 'first',
        'Beilagen': 'first',
        'Verfasser / Urheber': 'first',
        'Verlag': 'first',
        'Ort': 'first',
        'Jahr': 'first',
        'Ausleihzähler gesamt': 'sum',
        'Ausleihzähler lf. Jahr': 'sum',
        'Ausleihzähler Vorjahr': 'sum',
        'Ausleihzähler Vorvorjahr': 'sum'
    })

    # add a column which contains the numbers from all years before 'the year before last year'
    column_name = 'Ausleihzähler aller Jahre vor Vorvorjahr'

    input_file_rearranged[column_name] = input_file_rearranged["Ausleihzähler gesamt"] - input_file_rearranged["Ausleihzähler lf. Jahr"] \
                                         - input_file_rearranged["Ausleihzähler Vorjahr"] - input_file_rearranged["Ausleihzähler Vorvorjahr"]

    # save new table as Excel file
    output_file_name = file_name + '_updated.xlsx'
    input_file_rearranged.to_excel(output_file_name, index=False)