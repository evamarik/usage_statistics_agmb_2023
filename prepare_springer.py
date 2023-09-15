import pandas as pd
import openpyxl
from openpyxl.reader.excel import load_workbook
import xlsxwriter
import csv
import os


# input: workbook_name without .xlsx
# output: returns name of new tsv file with .tsv
def create_ISBN_list(workbook_name, worksheet_name):
    path_name = workbook_name + '.xlsx'
    wb = load_workbook(path_name)
    ws = wb[worksheet_name]

    isbn_list = []

    # iterate over all the rows in the sheet
    for row in ws:
        # use the row only if it has not been filtered out (i.e., it's not hidden)
        if ws.row_dimensions[row[0].row].hidden == False:
            isbn = row[0].value
            if str(isbn).startswith('9'):
                isbn_list.append(isbn)

    # create a new workbook
    isbn_wb_name = 'isbn_owned_by_TBM.tsv'

    with open(isbn_wb_name, 'w', newline='') as file:
        writer = csv.writer(file, delimiter='\t')
        column_name = ['ISBN']
        writer.writerow(column_name)
        for isbn in isbn_list:
            writer.writerow([isbn])

    return isbn_wb_name


# input: report_name without .tsv
# output: returns name of new tsv file with .tsv
def prepare_report_file(report_name):
    input_path_name = report_name + '.tsv'
    report_file = pd.read_csv(input_path_name, sep='\t', skiprows=13)
    report_file_unique = report_file[report_file.Metric_Type == 'Unique_Title_Requests']
    report_file_unique.reset_index(drop=True)

    output_path_name = report_name + '_updated.tsv'
    report_file_unique.to_csv(output_path_name, sep='\t', index=True)

    return output_path_name


# filter prepared report file -> only keep books the TBM owns
# workbook_name without .xlsx and report_name without .tsv
def merge_with_report(workbook_name, worksheet_name, report_name, year):
    # create ISBN tsv for the merge if it does not already exist
    if os.path.exists('isbn_owned_by_TBM.tsv'):
        isbn_wb_name = 'isbn_owned_by_TBM.tsv'
    else:
        isbn_wb_name = create_ISBN_list(workbook_name, worksheet_name)

    isbn_list = pd.read_csv(isbn_wb_name, sep='\t')

    # update report file
    report_path_name = prepare_report_file(report_name)
    report_file = pd.read_csv(report_path_name, sep='\t')

    # merge isbn_list and report_file (report_file is the prepared version)
    TBM_report_file = pd.merge(left=isbn_list, right=report_file, how='inner', left_on=['ISBN'], right_on=['ISBN'])

    # save result as new Excel file
    output_file_name = str(year) + '_Springer_Nature_TBM.xlsx'
    TBM_report_file.to_excel(output_file_name, index=False)


wb = '...'  # name of Excel file with data to all the Springer e-books owned by your branch library
ws = '...'  # name of the worksheet
report = ''  # name of a year report

merge_with_report(wb, ws, report, '...')  # add the year of the report inside ''
