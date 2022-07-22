import csv
import openpyxl

def ranges_calculator(weeks_and_months_csv, whs_anual_metrics_xlsx, worksheet_whs_anual_metrics_xlsx, index_openpyxl_padapter):
    """This function get ranges indexes from a input .csv file and columns indexes from a output .xlsx"""

    # Load weeks, months, quarters and year from Weeks and months.csv, make a list of each with it index
    with open(weeks_and_months_csv) as f:
        reader = csv.reader(f, delimiter = ';')
        header_row = next(reader)

        list_weeks = []
        month_list = []
        quarter_list = []
        year = f'{header_row[3]}'
        # print(year)
        for row in reader:
            if (row != []):
                list_weeks.append(row[0])
                month_list.append(row[1])
                quarter_list.append(row[2])
            else:
                continue

    # Initializes a list to store ranges from Weeks and months.csv, that happens to be the same of the Site metrics by week worksheet
    headers_index_csv = []

    # Creates a dictionary to each month, and assignee it first and last index from Weeks and months.csv
    # Then append it to a general list of headers indexes
    # The '+ 4' is to adapt to openpyxl module 
    list_months_first_week_index = []
    list_months_last_week_index = []
    list_dicitonaries_months_indexes = []

    for month in month_list:
        dictionaries_months_first_week_index = {
            'header': month, 
            'first_index': month_list.index(month)
            }
        if dictionaries_months_first_week_index not in list_months_first_week_index:
            list_months_first_week_index.append(dictionaries_months_first_week_index)
        else:
            continue

    month_list.reverse()

    for month in month_list:
        dictionaries_months_last_week_index = {
            'header': month, 
            'last_index': len(month_list) - month_list.index(month)
            }
        if dictionaries_months_last_week_index not in list_months_last_week_index:
            list_months_last_week_index.append(dictionaries_months_last_week_index)
        else:
            continue

    for dictionaries_months_first_week_index in list_months_first_week_index:
        for dictionaries_months_last_week_index in list_months_last_week_index:
            if dictionaries_months_first_week_index['header'] == dictionaries_months_last_week_index['header']:
                month_dictionary_first_and_last_indexes = {
                    'header': f"{dictionaries_months_first_week_index['header'].title()}/{year}", 
                    'first_index': (int(dictionaries_months_first_week_index['first_index']) + index_openpyxl_padapter), 
                    'last_index': (int(dictionaries_months_last_week_index['last_index']) + index_openpyxl_padapter)
                    }
                list_dicitonaries_months_indexes.append(month_dictionary_first_and_last_indexes)
            else:
                continue

    month_list.reverse()

    for item in list_dicitonaries_months_indexes:
        headers_index_csv.append(item)

    # Creates a dictionary to each quarter, and assignee it first and last index from Weeks and months.csv
    # Then append it to a general list of headers indexes
    # The '+ 4' is to adapt to openpyxl module 
    list_quarters_first_week_index = []
    list_quarters_last_week_index = []
    list_dicitonaries_quarters_indexes = []

    for quarter in quarter_list:
        dictionaries_quarters_first_week_index = {
            'header': quarter, 
            'first_index': quarter_list.index(quarter)
            }
        if dictionaries_quarters_first_week_index not in list_quarters_first_week_index:
            list_quarters_first_week_index.append(dictionaries_quarters_first_week_index)
        else:
            continue

    quarter_list.reverse()

    for quarter in quarter_list:
        dictionaries_quarters_last_week_index = {
            'header': quarter, 
            'last_index': len(quarter_list) - quarter_list.index(quarter)
            }
        if dictionaries_quarters_last_week_index not in list_quarters_last_week_index:
            list_quarters_last_week_index.append(dictionaries_quarters_last_week_index)
        else:
            continue

    for dictionaries_quarters_first_week_index in list_quarters_first_week_index:
        for dictionaries_quarters_last_week_index in list_quarters_last_week_index:
            if dictionaries_quarters_first_week_index['header'] == dictionaries_quarters_last_week_index['header']:
                quarter_dictionary_first_and_last_indexes = {
                    'header': f"{dictionaries_quarters_first_week_index['header'].title()}/{year}", 
                    'first_index': (int(dictionaries_quarters_first_week_index['first_index']) + index_openpyxl_padapter), 
                    'last_index': (int(dictionaries_quarters_last_week_index['last_index']) + index_openpyxl_padapter)
                    }
                list_dicitonaries_quarters_indexes.append(quarter_dictionary_first_and_last_indexes)
            else:
                continue

    quarter_list.reverse()

    for item in list_dicitonaries_quarters_indexes:
        headers_index_csv.append(item)

    # Creates a dictionary to year, and assignee it first and last index from Weeks and months.csv
    # Then append it to a general list of headers indexes
    # The '+ 4' is to adapt to openpyxl module 
    dictionary_year_indexes = {
        'header': f"Year{year}", 
        'first_index': (int(list_weeks.index(list_weeks[0])) + index_openpyxl_padapter), 
        'last_index': (int(len(month_list)) + index_openpyxl_padapter)
        }
    headers_index_csv.append(dictionary_year_indexes)

    # Creates a list of dictionaries to each header in Site metrics by year worksheet, and assignee it column index

    # Start openpyxl library to open existing .xlsx file
    workbook = openpyxl.load_workbook(whs_anual_metrics_xlsx)

    # Works on metrics by year worksheet
    worksheet_whs_anual_metrics = workbook[worksheet_whs_anual_metrics_xlsx]

    # Get metrics by year worksheet headers to set first and last index from week's
    headers_list = []
    for col in worksheet_whs_anual_metrics['1']:
        headers_list.append(col.value)

    headers_columns = []
    for header in headers_list:
        dict_first_header_index = {
            'header': header, 
            'col': headers_list.index(header) + 1
            }
        if dict_first_header_index['header'] != None:
            headers_columns.append(dict_first_header_index)
        else:
            continue

    # Bring data togehter from previous methods to a list:
    # The range of each period get from Weeks and months.csv, that's the same of the Site metrics by week worksheet
    # The colum of each period present in Site metrics by year worksheet
    list_sum_data_dictionary = []

    for header_index_csv in headers_index_csv:
        for header_column in headers_columns:
            if header_index_csv['header'] == header_column['header']:
                sum_data_dicitonary = {
                    'header': header_column['header'], 
                    'first_index': header_index_csv['first_index'], 
                    'last_index': header_index_csv['last_index'], 
                    'data_column': header_column['col']
                    }
                list_sum_data_dictionary.append(sum_data_dicitonary)
            else:
                continue

    return list_sum_data_dictionary