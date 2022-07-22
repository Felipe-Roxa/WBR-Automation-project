def get_index_with_none(a_list, a_string):
    """This function return a index of a string in a given list that contains none values"""
    item_index = 0

    for item in a_list:
        if item != None:
            if a_string.lower() == item.lower():
                item_index = a_list.index(item)
            else:
                continue
        else:
            continue

    return item_index
