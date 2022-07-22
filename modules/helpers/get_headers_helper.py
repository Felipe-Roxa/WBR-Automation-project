def get_headers(worksheet_and_line):
    """This function returns a list of headers from a given worsheet line of openpyxl module"""

    header_list = []
    for item in worksheet_and_line:
        if item.value != None:
            header_list.append(item.value)
        else:
            continue
    return header_list