def get_index(a_list, a_string):
    """This function return a index of a string in a given list"""
    item_index = 0

    for item in a_list:
        if a_string.lower() == item.lower():
            item_index = a_list.index(item)
        else:
            continue

    return item_index