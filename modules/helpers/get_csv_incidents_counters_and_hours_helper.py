import csv

def get_csv_incidents_counters_and_hours(csv_file):
    """This function returns a list of incidents counters and hours from a .csv file"""

    with open(csv_file) as f:
        reader = csv.reader(f, delimiter = ';')
        header_row = next(reader)

        list_incidents_counters_and_hours_title = []
        for row in reader:
            if (row != []) and (row[1] != 'rate'):
                list_incidents_counters_and_hours_title.append(f"{row[0]}")
            else:
                continue

    list_incidents_counters_and_hours = []
    for incident in list_incidents_counters_and_hours_title:
        incident = incident.lower().replace(" ", "_")
        list_incidents_counters_and_hours.append(incident)

    return list_incidents_counters_and_hours