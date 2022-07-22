def get_headers_with_none(worksheet_and_line):
    """This function returns a list of headers (including none values) from a given worsheet line of openpyxl module"""

    header_list = []
    for item in worksheet_and_line:
        header_list.append(item.value)

    return header_list