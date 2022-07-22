import csv
import xlsxwriter

def whs_anual_metrics_excel_generator(incidents_count_and_rates_csv, weeks_and_months_csv, sites_and_categories_csv):
    """This function generates a empty report using .csv files as parameters"""

    # Load incidents from Incidents count and rates.csv
    with open(incidents_count_and_rates_csv) as f:
        reader = csv.reader(f, delimiter = ';')
        header_row = next(reader)

        list_incidents = []
        for row in reader:
            if row != []:
                list_incidents.append(f"{row[0]}")
            else:
                continue

    # Load weeks from Weeks and months.csv
    with open(weeks_and_months_csv) as f:
        reader = csv.reader(f, delimiter = ';')
        weeks_and_months_header_row = next(reader)

        list_quarters = []
        list_months = []
        list_weeks = []
        for row in reader:
            if row != []:
                list_weeks.append(f"{row[0]}")
                if row[1] not in list_months:
                    list_months.append(row[1])
                else:
                    continue
                if row[2] not in list_quarters:
                    list_quarters.append(row[2])
                else:
                    continue
            else:
                continue
        year = f"{weeks_and_months_header_row[3]}"

    # Load sites from Sites and categories.csv
    with open(sites_and_categories_csv) as f:
        reader = csv.reader(f, delimiter = ';')
        sites_and_categorie_header_row = next(reader)

        # Creates a dictionary for each site and a list of categories
        list_sites_dictionaries_wbr = []
        list_site_categories_wbr = []
        for row in reader:
            if row != []:
                site_dict_wbr = {
                    f'{sites_and_categorie_header_row[0]}':f'{row[0]}', 
                    f'{sites_and_categorie_header_row[1]}':f'{row[1]}'
                    }
                list_sites_dictionaries_wbr.append(site_dict_wbr)

                if row[0] not in list_site_categories_wbr:
                    list_site_categories_wbr.append(row[0])
                else:
                    continue
            else:
                continue
    
    # Start xlsxwriter library and export all previous data generated on this file to a excel
    # Writes a empty file
    workbook = xlsxwriter.Workbook('Output/WHS Anual Metrics.xlsx')

    worksheet_site_last_metrics = workbook.add_worksheet('Last site metrics')
    worksheet_category_last_metrics = workbook.add_worksheet('Last category metrics')
    worksheet_site_metrics_by_week = workbook.add_worksheet('Site metrics by week')
    worksheet_site_metrics_by_month_and_year = workbook.add_worksheet('Site metrics by year')
    worksheet_category_metrics_by_week = workbook.add_worksheet('Category metrics by week')
    worksheet_category_metrics_by_month_and_year = workbook.add_worksheet('Category metrics by year')

    # Writes on both worksheets rows time intervals
    row = 0
    col = 3
    for week in list_weeks:
        worksheet_site_metrics_by_week.write(row, col, week)
        col += 1

    row = 0
    col = 3
    for month in list_months:
        worksheet_site_metrics_by_month_and_year.write(row, col, f'{month.title()}/{year}')
        col += 1
    col += 1
    for quarter in list_quarters:
        worksheet_site_metrics_by_month_and_year.write(row, col, f'{quarter.upper()}/{year}')
        col += 1
    col += 1
    worksheet_site_metrics_by_month_and_year.write(row, col, f'Year{year}')

    # Writes sites and its incidents
    row = 1
    col = 0
    for site_dictionary in list_sites_dictionaries_wbr:
        for incident in list_incidents:
            worksheet_site_metrics_by_week.write(row, col, site_dictionary[f'{sites_and_categorie_header_row[0]}'].title())
            worksheet_site_metrics_by_week.write(row, col + 1, site_dictionary[f'{sites_and_categorie_header_row[1]}'].upper())
            worksheet_site_metrics_by_week.write(row, col + 2, incident)
            worksheet_site_metrics_by_month_and_year.write(row, col, site_dictionary[f'{sites_and_categorie_header_row[0]}'].title())
            worksheet_site_metrics_by_month_and_year.write(row, col + 1, site_dictionary[f'{sites_and_categorie_header_row[1]}'].upper())
            worksheet_site_metrics_by_month_and_year.write(row, col + 2, incident)
            worksheet_site_last_metrics.write(row, col, site_dictionary[f'{sites_and_categorie_header_row[0]}'].title())
            worksheet_site_last_metrics.write(row, col + 1, site_dictionary[f'{sites_and_categorie_header_row[1]}'].upper())
            worksheet_site_last_metrics.write(row, col + 2, incident)            
            row += 1

    # Writes on both worksheets rows time intervals
    row = 0
    col = 2
    for week in list_weeks:
        worksheet_category_metrics_by_week.write(row, col, week)
        col += 1

    row = 0
    col = 2
    for month in list_months:
        worksheet_category_metrics_by_month_and_year.write(row, col, f'{month.title()}/{year}')
        col += 1
    col += 1
    for quarter in list_quarters:
        worksheet_category_metrics_by_month_and_year.write(row, col, f'{quarter.upper()}/{year}')
        col += 1
    col += 1
    worksheet_category_metrics_by_month_and_year.write(row, col, f'Year{year}')

    # Writes categories and its incidents
    row = 1
    col = 0
    for site_category in list_site_categories_wbr:
        for incident in list_incidents:
            worksheet_category_metrics_by_week.write(row, col, site_category.title())
            worksheet_category_metrics_by_week.write(row, col + 1, incident)
            worksheet_category_metrics_by_month_and_year.write(row, col, site_category.title())
            worksheet_category_metrics_by_month_and_year.write(row, col + 1, incident)
            worksheet_category_last_metrics.write(row, col, site_category.title())
            worksheet_category_last_metrics.write(row, col + 1, incident)
            row += 1

    workbook.close()