import pandas as pd
import numpy as np
import os
import time


def read_dataframe(fileName) -> pd.DataFrame:
    """Function - read_dataframe:
        -> return - DataFrame object with values from a file"""

    if os.path.getsize(fileName) == 0:
        print("File is EMPTY!!!")
    else:
        return pd.read_excel(fileName)


def conv_file_to_list(dataFrame) -> list:
    """Function - Converts dataframe to a list
    convert nested 2-dimensional array to 1d array
    sort alphabetical order and convert to a list"""

    newList = []
    items = dataFrame.to_numpy()
    formatted = np.reshape(items, (1, -1)).tolist()
    modifiedList = np.sort(formatted).tolist()

    for sublist in modifiedList:
        for item in sublist:
            newList.append(item)

    return newList


def compare_lists(list1, list2) -> list:
    match_list = []
    for item in list1:
        for value in list2:
            if item == value:
                match_list.append(value)
    return match_list


def write_to_excel(matched_list) -> None:
    if os.path.isfile('matched_data.xlsx'):
        print('File exists')
    else:
        dataframe = pd.DataFrame(matched_list)
        dataframe.to_excel('matched_data.xlsx', header=False, index=False)


def main():
    start = time.process_time()

    xcl_1 = read_dataframe('applist.xlsx')
    final_list1 = conv_file_to_list(xcl_1)

    xcl_2 = read_dataframe('applist2.xlsx')
    final_list2 = conv_file_to_list(xcl_2)

    match_found = compare_lists(final_list1, final_list2)
    write_to_excel(match_found)

    print(time.process_time() - start)


if __name__ == '__main__':
    main()
