import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl as op

input_name = 'Bestandsliste_172-dr3Dj_updated.xlsx'  # File that is created by 'prepare_excel_file'
work_sheet = 'Sheet1'

# read Excel file
Bestandsliste_172 = pd.read_excel(input_name, sheet_name=work_sheet)

top_amount = 20

# sort Bestandsliste_172 and save the sorted lists and their tops separately:
# overall:
Bestandsliste_172_overall = Bestandsliste_172.sort_values(by='Ausleihzähler gesamt', ascending=False)
overall_top = Bestandsliste_172_overall.head(top_amount)

# current year:
Bestandsliste_172_current = Bestandsliste_172.sort_values(by='Ausleihzähler lf. Jahr', ascending=False)
current_top = Bestandsliste_172_current.head(top_amount)

# last year:
Bestandsliste_172_last = Bestandsliste_172.sort_values(by='Ausleihzähler Vorjahr', ascending=False)
last_top = Bestandsliste_172_last.head(top_amount)
last_top_10 = Bestandsliste_172_last.head(10)

# year before last:
Bestandsliste_172_before_last = Bestandsliste_172.sort_values(by='Ausleihzähler Vorvorjahr', ascending=False)
before_last_top = Bestandsliste_172_before_last.head(top_amount)

# all years before 'the year before last year' (in our case 2020):
Bestandsliste_172_rest = Bestandsliste_172.sort_values(by='Ausleihzähler aller Jahre vor Vorvorjahr', ascending=False)
rest_top = Bestandsliste_172_rest.head(top_amount)



'''Some parameters used in the following functions:
    column_name: name of one of the columns that contains usage numbers
    top: one of the top_lists that were created above 
    input_file: should be a DataFrame containing information from the whole list, e.g.: Bestandsliste_172
    bar_names: list that contains four strings, one for each bar '''

def add_labels_vertical(x, y, text, size):
    for i in range(len(x)):
        plt.text(i, y[i], text[i], fontsize=size, fontweight='bold', horizontalalignment='center')


def add_labels_horizontal(usage, size):
    for i, v in enumerate(usage):
        plt.text(v + 1.5, i, str(v), va='center', color='black', fontweight='bold', fontsize=size)


def plot_top_of_year(top, column_name, plot_title):
    usage = []
    usage.extend(int(values) for values in top[column_name].values)  # additionally convert values to int

    title = []
    title.extend(top['Titel'].values)

    y_position = np.arange(len(title))

    plt.figure(figsize=(30, 20))
    plt.barh(y_position, usage, color='#ad007c')

    # x, y- axes:
    plt.xlim(0, 500)
    plt.xticks(fontsize=50)

    # write usage number next to bar
    add_labels_horizontal(usage, 45)

    plt.yticks(y_position, title, fontsize=50)

    # save plot
    plot_name = 'top_' + str(len(top)) + '_' + plot_title.replace(" ", "_") + '.svg'
    plt.savefig(plot_name, bbox_inches='tight')

    plt.tight_layout()
    plt.show()


def plot_top_10_2022_shorttitle(top):
    column_name = 'Ausleihzähler Vorjahr'

    title_file = pd.read_excel('Bestandsliste_172-dr3Dj_updated_Kurznamen.xlsx')  # list with manually added shortnames
    top_10_title = title_file.head(10)

    usage = []
    usage.extend(int(values) for values in top[column_name].values)  # convert values to int

    title = []
    title.extend(top_10_title['Kurzname'].values)

    y_position = np.arange(len(title))

    plt.figure(figsize=(30, 20))
    plt.barh(y_position, usage, color='#ad007c')

    # x, y- axes:
    plt.xlim(0, 500)
    plt.xticks(fontsize=50)

    # write usage number next to bar
    add_labels_horizontal(usage, 45)

    plt.yticks(y_position, title, fontsize=50)

    # save plot
    plot_name = 'top_10_2022_Kurztitel.svg'
    plt.savefig(plot_name, bbox_inches='tight')

    plt.tight_layout()
    plt.show()


def plot_top_over_years(input_file, top_current, top_last, top_before_last, top_rest, bar_names):
    # sum of lend books per year
    current_sum = 0
    last_sum = 0
    before_last_sum = 0
    rest_sum = 0

    length = input_file.shape[0]
    i = 0

    while i < (length-1):
        current_sum += input_file.iloc[i, 10]
        last_sum += input_file.iloc[i, 11]
        before_last_sum += input_file.iloc[i, 12]
        rest_sum += input_file.iloc[i, 13]
        i += 1

    # sum of tops per year
    current_top_sum = 0
    last_top_sum = 0
    before_last_top_sum = 0
    rest_top_sum = 0

    j = 0
    while j < top_amount-1:
        current_top_sum += top_current.iloc[j, 10]
        last_top_sum += top_last.iloc[j, 11]
        before_last_top_sum += top_before_last.iloc[j, 12]
        rest_top_sum += top_rest.iloc[j, 13]
        j += 1

    # put the values in lists
    sum_list = [rest_sum, before_last_sum, last_sum, current_sum]
    top_sum_list = [rest_top_sum, before_last_top_sum, last_top_sum, current_top_sum]


    # plot the tops at the top of the bar (in order to do so plot the sums and then sum minus tops):
    UP_list = []  # the sums-minus-tops list

    k = 0
    while k < len(bar_names):
       x = sum_list[k] - top_sum_list[k]
       UP_list.append(x)
       k += 1

    x_position = np.arange(len(bar_names))

    # plot
    w = 0.65
    top_label = 'Top ' + str(top_amount)
    plt.rcParams['hatch.color'] = 'white'
    plt.rcParams['hatch.linewidth'] = 1.5
    plt.bar(bar_names, sum_list, color='#ad007c', hatch='///', width=w, label=top_label)
    plt.bar(bar_names, UP_list, color='#ad007c', width=w, label='Lend overall')
    plt.xticks(fontsize=10)
    plt.yticks(fontsize=10)

    plt.legend()

    # save plot
    plot_name = 'relation_top_overall.svg'
    plt.savefig(plot_name)

    plt.tight_layout()
    plt.show()


def get_preclinic (input_file):
    # the following tuple contains pressmark-prefixes for pre-clinic books
    pressmark_tuple = ('172/W', '172/XB 1400', '172/XB 2200', '172/XC 2800 S389', '172/XC 2801', '172/XC 30', '172/XC 34',
                       '172/XC 38', '172/XD 1300', '172/XD 1500', '172/XF 3400', '172/YH 1300', '172/YH 1400 K42')

    current_sum = 0
    last_sum = 0
    before_last_sum = 0
    rest_sum = 0

    length = input_file.shape[0]
    i = 0
    while i < (length-1):
        pressmark = input_file.iloc[i, 0]
        if pressmark.startswith(pressmark_tuple):
            current_sum += input_file.iloc[i, 10]
            last_sum += input_file.iloc[i, 11]
            before_last_sum += input_file.iloc[i, 12]
            rest_sum += input_file.iloc[i, 13]
        i += 1

    preclinic_over_years = {
        'current year': current_sum,
        'last year': last_sum,
        'year before last year': before_last_sum,
        'rest': rest_sum
    }

    return preclinic_over_years


def plot_preclinic_over_years(input_file, bar_names):
    preclinic_dict = get_preclinic(input_file)

    current_preclinic = preclinic_dict['current year']
    last_preclinic = preclinic_dict['last year']
    before_last_preclinic = preclinic_dict['year before last year']
    rest_preclinic = preclinic_dict['rest']

    # sum of lend books per year
    current_sum = 0
    last_sum = 0
    before_last_sum = 0
    rest_sum = 0

    length = input_file.shape[0]
    i = 0

    while i < (length - 1):
        current_sum += input_file.iloc[i, 10]
        last_sum += input_file.iloc[i, 11]
        before_last_sum += input_file.iloc[i, 12]
        rest_sum += input_file.iloc[i, 13]
        i += 1

    # put the values in lists
    sum_list = [rest_sum, before_last_sum, last_sum, current_sum]
    preclinic_list = [rest_preclinic, before_last_preclinic, last_preclinic, current_preclinic]


    # calculate percentage: amount of pre-clinic books in books lend overall
    preclinic_percentage = []
    j = 0
    while j < len(bar_names):
        x = float(preclinic_list[j] / sum_list[j]) * 100.00
        preclinic_percentage.append(str(round(x)) + '%')
        j += 1

    # plot preclinic numbers at the top of the bar (in order to do so plot the sums and then sum minus pre-clinic):
    UP_list = []

    k = 0
    while k < len(bar_names):
        x = sum_list[k] - preclinic_list[k]
        UP_list.append(x)
        k += 1

    x_position = np.arange(len(bar_names))

    # plot
    w = 0.65
    preclinic_label = 'Vorklinik'
    plt.rcParams['hatch.color'] = 'white'
    plt.rcParams['hatch.linewidth'] = 1.5
    plt.bar(bar_names, sum_list, color='#ad007c', hatch='//', width=w, label=preclinic_label)
    plt.bar(bar_names, UP_list, color='#ad007c', width=w, label='Insgesamt entliehen')
    plt.xticks(fontsize=15)
    plt.yticks(fontsize=15)

    # add percentage
    add_labels_vertical(bar_names, sum_list, preclinic_percentage, 15)


    plt.legend(fontsize=13, loc='upper left')

    # save plot
    plot_name = 'relation_preclinic_overall.svg'
    plt.savefig(plot_name)

    plt.tight_layout()
    plt.show()


# use the functions:
bars = ['2020', '2021', '2022', '2023']
plot_top_over_years(Bestandsliste_172, current_top, last_top, before_last_top, rest_top, bars)
plot_preclinic_over_years(Bestandsliste_172, bars)

# plot the top 10s for 2022
plot_top_of_year(last_top_10, 'Ausleihzähler Vorjahr', '2022')
plot_top_10_2022_shorttitle(last_top_10)
