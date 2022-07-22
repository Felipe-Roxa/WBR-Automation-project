import csv

def get_csv_data_site(csv_file):
    """This function returns a list of data from a .csv file"""

    with open(csv_file) as f:
        reader = csv.reader(f, delimiter = ';')
        header_row = next(reader)

        list_site_data_title = []
        for row in reader:
            if row != []:
                list_site_data_title.append(f"{row[0]}")
            else:
                continue

    list_site_data = []
    for site_data in list_site_data_title:
        site_data = site_data.lower().replace(" ", "_")
        list_site_data.append(site_data)
    
    return list_site_data