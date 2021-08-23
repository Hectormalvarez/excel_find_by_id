#!/usr/bin/env python3
"""
Author : hadev <hadev@localhost>
Date   : 2021-08-21
Purpose: Get Prices From Manufacturer
"""

import argparse
import os
import sys
from openpyxl.reader.excel import load_workbook


def get_args():
    """Get command-line arguments"""

    parser = argparse.ArgumentParser(
        description='Count Words',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)

    parser.add_argument('file',
                        help='A readable file',
                        nargs='*', )

    args = parser.parse_args()

    return args


# --------------------------------------------------
def get_workbook_to_search():
    files = dict(enumerate([f for f in os.listdir(os.curdir) if os.path.isfile(f) and f[-4:] == "xlsx"]))

    # list worksheets in same directory
    for key, file in files.items():
        print(str(key) + " " + file)

    # have user enter number key to workbook name than sheet to compare from
    print("Please enter number next to workbook")
    workbook_number = int(input("Number left of workbook name: "))  # ask user for number
    print("------------------------------------\n")
    print("Loading: " + files[workbook_number])  # prints workbook name of selected number
    print("------------------------------------\n")
    workbook_object = load_workbook(files[workbook_number])  # loads workbook into variable
    sheet_dict = dict(enumerate(workbook_object.sheetnames))  # loads sheet names into a dict
    for key, file in sheet_dict.items():  # prints out index and sheet for user to select
        print(str(key) + " " + file)
    print("Please Enter Sheet we will work from")
    print("------------------------------------\n")
    sheet_number = int(input("number left of sheetname: "))
    print("------------------------------------\n")

    return workbook_object, workbook_object[sheet_dict[sheet_number]], files[workbook_number]


# --------------------------------------------------
def find_column(sheet_object, column_name):
    # ask if identifying column is known
    identifying_column = input(f"{column_name} Column?(leave blank if unknown): ")
    print("------------------------------------\n")
    if identifying_column == "":  # if nothing is entered
        looking_for_source_identifier = True  # we will proceed to look for identifying column
        column_to_look_in = 1  # start with row 1
    else:
        looking_for_source_identifier = False  # if something was entered we are NOT looking
        identifying_column = int(identifying_column)  # turn into integer
    # Looking for the identifying column if not known
    while looking_for_source_identifier:
        for cell in sheet_object[column_to_look_in]:  # Going to list each rows value and row
            if cell.value is not None:
                print(str(cell.column) + ": " + str(cell.value))
            else:
                continue
            column_to_look_in += 1

        # ask if identifying value is listed
        is_looking_for_source_identifier = input("Enter number of Identifier(leave blank if not listed): ")
        print("------------------------------------\n")

        # if identifying column is entered it is assigned, if blank looking continues
        if is_looking_for_source_identifier != "":
            identifying_column = int(is_looking_for_source_identifier)
            looking_for_source_identifier = False

    return identifying_column


# --------------------------------------------------
def main():
    """Make a jazz noise here"""

    args = get_args()

    print("Lets Find Source Worksheet")
    print("------------------------------------")
    source_workbook_object, source_sheet_object, source_workbook_name = get_workbook_to_search()
    source_identifying_column_number = find_column(source_sheet_object, "Identifying")
    source_list_price_column_number = find_column(source_sheet_object, "List Price")
    sso = source_sheet_object
    sicn = source_identifying_column_number
    slpcn = source_list_price_column_number

    print("Lets Find Manufacturer Worksheet")
    print("------------------------------------")
    manufacturer_workbook_object, manufacturer_price_sheet_object, manufacturer_workbook_name = get_workbook_to_search()
    manufacturer_identifying_column_number = find_column(manufacturer_price_sheet_object, "Identifying")
    manufacturer_list_price_column_number = find_column(manufacturer_price_sheet_object, "List Price")
    mpo = manufacturer_price_sheet_object
    mlpcn = manufacturer_list_price_column_number
    micn = manufacturer_identifying_column_number

    # start iterations
    for source_identifying_column in sso.iter_cols(max_col=sicn,  # Generates column for iterating
                                                   min_col=sicn):
        for source_cell in source_identifying_column:  # for each cell in source identifying column
            if source_cell.value is not None:  #
                for manufacturer_column in mpo.iter_cols(max_col=micn, min_col=micn):
                    for manufacturer_cell in manufacturer_column:
                        if source_cell.value == manufacturer_cell.value:
                            mcr = manufacturer_cell.row  # row of matching manufacturer cell
                            scr = source_cell.row  # row of matching source cell
                            manufacturer_list_price = mpo.cell(row=mcr, column=mlpcn).value
                            print(str(source_cell.value) + " price: " + str(mpo.cell(row=mcr, column=mlpcn).value))
                            sso.cell(row=scr, column=slpcn).value = manufacturer_list_price
                        else:
                            continue
            else:
                continue

        print("Done Updating!")
        source_workbook_object.save(source_workbook_name)


# --------------------------------------------------
if __name__ == '__main__':
    main()
