import itertools
import math
import seaborn as sns
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt



''' Parameters used in the following functions:
    file_name: without '.xlsx'
    year: must be a string with four digits
    publisher: must be one of those 3: 'Thieme', 'Springer', 'Elsevier' '''

# the get-functions remove books that were lend zero times
def get_springer_usage(file_name):
    file_path = file_name + '.xlsx'
    report_file = pd.read_excel(file_path)

    usage_list = []

    length = report_file.shape[0]
    i = 0
    while i < length:
        usage = report_file.iloc[i, 13]  # 13 is 'Reporting_Period_Total' column (zero based)
        if usage != 0:
            usage_list.append(int(usage))
        i += 1

    return usage_list


def get_thieme_usage(file_name):
    file_path = file_name + '.xlsx'
    report_file = pd.read_excel(file_path)

    usage_list = []

    length = report_file.shape[0]
    i = 0
    while i < length:
        usage = report_file.iloc[i, 1]  # 1 is 'Reporting_Period_Total' column (zero based)
        if usage != 0:
            usage_list.append(int(usage))
        i += 1

    return usage_list


def get_thime_usage_2020():
    file_path = 'Thieme_2020_eRef-Lehrbücher.xlsx'
    report_file = pd.read_excel(file_path, sheet_name='Zusammenfassung')

    usage_list = []

    length = 13
    i = 0
    while i < length:
        usage = int(report_file.iloc[i, 1])  # 1 is 'Sum_Unique_Title_Requests_2020' column (zero based)
        if usage != 0:
            usage_list.append(int(usage))
        i += 1

    return usage_list


def get_elsevier_usage(file_name, year):
    file_path = file_name + '.xlsx'
    report_file = pd.read_excel(file_path, skiprows=9)

    column_name = 'Total ' + year
    column_index = report_file.columns.get_loc(column_name)

    usage_list = []

    length = report_file.shape[0] - 1  # 'remove' last row, since it contains information we don't need for this code
    i = 0
    while i < length:
        usage = report_file.iloc[i, column_index]
        if math.isnan(usage) or usage == 0:
            pass
        else:
            usage_list.append(int(usage))
        i += 1

    return usage_list


def publishers_per_year(year):
    thieme_list = []
    springer_list = []
    elsevier_list = []

    # get Thieme usage
    if year == '2020':
        thieme_list = get_thime_usage_2020()
    else:
        thieme_name = 'Thieme_TBM_Liste_' + year + '_updated'   # based on how we named our files: please change when using this code!!
        thieme_list = get_thieme_usage(thieme_name)

    # get Springer usage
    springer_name = year + '_Springer_Nature_TBM'  # based on the names given by merge_with_report from prepare_springer.py
    springer_list = get_springer_usage(springer_name)

    # get Elsevier usage
    elsevier_name = '...'  # add file_name without '.xlsx' when using this code
    elsevier_list = get_elsevier_usage(elsevier_name, year)

    # create dataframe out of lists in order to plot them with seaborn
    thieme_len = len(thieme_list)
    springer_len = len(springer_list)
    elsevier_len = len(elsevier_list)

    # the following function creates a table with two columns 'Usage' and 'Publisher'. The first column lists all the values saved in the 3 lists, while
    # the second column contains the name of the publisher to which the data belongs
    # \n is for plotting
    publisher_usage = pd.DataFrame(
        {'Usage': thieme_list + springer_list + elsevier_list, 'Publisher': ['Thieme eRef'] * thieme_len +
                                                                            ['Springer\nMedizin-Pakete'] * springer_len +
                                                                            ['Elsevier\nClinicalKey Student'] * elsevier_len})

    # plot
    sns.set_theme(style="ticks")
    f, ax = plt.subplots(figsize=(15, 12))  # initialize the figure
    ax.set_yscale("log")  # logarithmic scale
    sns.boxplot(x='Publisher', y='Usage', data=publisher_usage, whis=[0, 100], width=.6,
                palette=['white'])  # plot UIR with vertical boxes
    sns.stripplot(x='Publisher', y='Usage', data=publisher_usage, color='#ad007c', size=7,
                  linewidth=0)  # add in points to show each observation

    ax.set_title(year, fontsize=55)
    ax.set_xlabel('')
    ax.set_ylabel('E-Book Ausleihe (logarithmisch)', fontsize=40)
    ax.tick_params(axis='x', labelsize=30)
    ax.tick_params(axis='y', labelsize=30)
    sns.despine(trim=True, left=True)

    plot_name = 'E-book_' + year + '_boxplots.svg'
    plt.savefig(plot_name)

    plt.tight_layout()
    plt.show()


def publisher_over_years(publisher):
    list_2020 = []
    list_2021 = []
    list_2022 = []

    # title will be used as the plots title
    if publisher == 'Thieme':
        list_2020 = get_thime_usage_2020()
        # based on how we named our files: please change when using this code!!
        list_2021 = get_thieme_usage('Thieme_TBM_Liste_2021_updated')
        list_2022 = get_thieme_usage('Thieme_TBM_Liste_2022_updated')
        title = publisher + ' eRef'
    elif publisher == 'Springer':
        # based on the names given by merge_with_report from prepare_springer.py
        list_2020 = get_springer_usage('2020_Springer_Nature_TBM')
        list_2021 = get_springer_usage('2021_Springer_Nature_TBM')
        list_2022 = get_springer_usage('2022_Springer_Nature_TBM')
        title = publisher + ' Medizin-Pakete'
    elif publisher == 'Elsevier':
        # add file_name without '.xlsx' and a year when using this code, e.g. get_elsevier_usage('Elsevier_Report', '2021')
        list_2020 = get_elsevier_usage('...', '...')
        list_2021 = get_elsevier_usage('...', '...')
        list_2022 = get_elsevier_usage('...', '...')
        title = publisher + ' ClinicalKey Student'
    else:
        print('Something went wrong, please check the publisher. \nYou entered: ' + publisher +
              '\nRemember, you can only enter one of these: Thieme, Springer, Elsevier')
        return

    # create dataframe out of lists in order to plot them with seaborn
    len_2020 = len(list_2020)
    len_2021 = len(list_2021)
    len_2022 = len(list_2022)

    # the following function creates a table with two columns 'Usage' and 'Year'. The first column lists all the values saved in the 3 lists, while
    # the second column contains the year which the data belongs to
    publisher_usage = pd.DataFrame(
        {'Usage': list_2020 + list_2021 + list_2022, 'Year': ['2020'] * len_2020 + ['2021'] * len_2021 + ['2022'] * len_2022})

    analysis = publisher_usage.groupby("Year").describe()
    print(publisher, '\n', analysis)

    # plot
    sns.set_theme(style="ticks")
    f, ax = plt.subplots(figsize=(15, 12))  # initialize the figure
    ax.set_yscale("log")  # logarithmic scale
    sns.boxplot(x='Year', y='Usage', data=publisher_usage, whis=[0, 100], width=.6,
                palette=['white'])  # plot UIR with vertical boxes
    sns.stripplot(x='Year', y='Usage', data=publisher_usage, color='#ad007c', size=9,
                  linewidth=0)  # add in points to show each observation

    if publisher == 'Elsevier':
        y_label = 'Chapter Requests (BR2, R4)'  # because Elsevier used Counter4
    else:
        y_label = 'Unique Title Requests'

    ax.set_title(title, fontsize=55)
    ax.set_xlabel('')
    ax.set_ylabel(y_label, fontsize=40)
    ax.tick_params(axis='x', labelsize=40)
    ax.tick_params(axis='y', labelsize=30)
    sns.despine(trim=True, left=True)

    plot_name = publisher + '_usage_over_years' + '_boxplots.svg'
    plt.savefig(plot_name)

    plt.tight_layout()
    plt.show()


# use the functions:
publishers_per_year('2020')
publishers_per_year('2021')
publishers_per_year('2022')

publisher_over_years('Thieme')
publisher_over_years('Springer')
publisher_over_years('Elsevier')