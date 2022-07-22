import os
from modules.whs_report_site_and_category_filler_by_week import site_and_category_filler_by_week
from modules.whs_report_site_and_category_calculator_by_month_to_year import site_and_category_calculator_by_month_to_year
from modules.empty_whs_report_exporter import whs_anual_metrics_excel_generator
from modules.whs_report_last_data_selector import last_data_selector
from modules.whs_report_pdf_exporter import pdf_exporter

weeks_and_months_csv = 'Input/Parameters/Weeks and months.csv'
sites_and_categories_csv = 'Input/Parameters/Sites and categories.csv'
incidents_count_and_rates_csv = 'Input/Parameters/Incidents count and rates.csv'
quip_directory = 'Input/Quip/Current Year'
whs_anual_metrics_xlsx = 'Output/WHS Anual Metrics.xlsx'
whs_weekly_metrics_report = 'Output/WHS Weekly Metrics Report.xlsx'

try:
	whs_anual_metrics_excel_generator(incidents_count_and_rates_csv, weeks_and_months_csv, sites_and_categories_csv)
except FileNotFoundError:
    print("Couldn't find nescessary files")
else:
    pass

try:
    list_files = os.listdir(quip_directory)
    for quip_file in list_files:
        file_name = f"{quip_directory}/{quip_file}"
        site_and_category_filler_by_week(file_name, sites_and_categories_csv, whs_anual_metrics_xlsx, incidents_count_and_rates_csv)
except FileNotFoundError:
    print("Couldn't find nescessary files")
else:
    pass

try:
	site_and_category_calculator_by_month_to_year(whs_anual_metrics_xlsx, weeks_and_months_csv, incidents_count_and_rates_csv)
except FileNotFoundError:
    print("Couldn't find nescessary files")
else:
    pass

try:
	last_data_selector(whs_anual_metrics_xlsx, weeks_and_months_csv, quip_directory)
except FileNotFoundError:
    print("Couldn't find nescessary files")
else:
    pass

try:
	pdf_exporter(whs_weekly_metrics_report)
except FileNotFoundError:
    print("Couldn't find nescessary files")
else:
    pass