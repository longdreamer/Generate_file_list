#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# Author: Kun.Li
# Date: 2017/11/16

import os
import xlrd
import shutil
from pypinyin import lazy_pinyin

# if you want to debug it with pycharm, open os.chdir('..')
os.chdir('..')

proj_path = os.getcwd()
print(proj_path)
self_name = os.path.basename(__file__)
drug_lib_filename = proj_path + '/Resource/drug_lib/default_drug_lib.xlsx'

drug_items_sheet = "drug_items"
drug_items_sheet_title_index = "drug index"
drug_items_sheet_title_english = "drug name/english"
drug_items_sheet_title_chinese = "drug name/simplified_chinese"
drug_items_sheet_title_portuguese = "drug name/portuguese"
drug_items_sheet_title_spanish = "drug name/spanish"
drug_items_sheet_title_is_display = "is display"
drug_items_sheet_title_tips = "tips"
drug_items_sheet_title_dose_lower = "dose lower limitation"
drug_items_sheet_title_dose_upper = "dose upper limitation"
drug_items_sheet_title_concentration_lower = "concentration lower limitation"
drug_items_sheet_title_concentration_uppwer = "concentration upper limitation"
drug_items_sheet_title_color = "drug color"

drug_category_sheet = "drug_categroy"
drug_category_sheet_title_index = "drug category index"
drug_category_sheet_title_english = "drug category name/english"
drug_category_sheet_title_chinese = "drug category name/simplified_chinese"
drug_category_sheet_title_portuguese = "drug category name/portuguese"
drug_category_sheet_title_spanish = "drug category name/spanish"
drug_category_sheet_title_count = "drug item count"
drug_category_sheet_title_items_index = "drug item index"
drug_category_sheet_title_color = "drug color"

version_sheet = "version"
version_major_row = 1
version_minor_row = 2
version_revision_row = 3
version_major_col = 1
version_minor_col = 1
version_revision_col = 1

lang_filename = proj_path + '/Resource/lang.xlsx'
lang_sheet = 'string_of_all_lang'
lang_sheet_title_english = "english"
lang_sheet_title_string_id = "string_id"

# print((lang_filename))
# print(drug_lib_filename)
# print(drug_lib_category_filename)

# print error function
def exit_p(str):
    print("Error: "+str)
    exit(-1)

def getExcleSheet(filename, sheet_name):
    if os.path.exists(filename):
       r_workbook = xlrd.open_workbook(filename)
       r_sh = r_workbook.sheet_by_name(sheet_name)
       return r_sh
    else:
        exit_p("%s does not exist" % filename)


def getExcleColsList(filename, sheet_name,col_index):
    col_list = []
    sheet = getExcleSheet(filename, sheet_name)
    for row in range(1, sheet.nrows):
        value = sheet.cell_value(row, col_index)
        col_list.append(value)
    return  col_list

def getTitleColIndex(filename, sheet_name, title):
    i_col = -1
    sheet = getExcleSheet(filename, sheet_name)
    for col in range(0, sheet.ncols):
        col_value = sheet.cell_value(0, col)
        if title == col_value:
            i_col = col
            break

    if i_col == -1 :
        exit_p("%s does not exist" % title)

    return i_col

def printStringIdListAndDrugNameListDiff(drug_name_in_lang_list, drug_name_list):

    if len(drug_name_list) >= len(drug_name_in_lang_list):
        print("drug name list length is great than drug name in lang.xlsx list length\n")
        for i in range(0, len(drug_name_list)):
            if i < len(drug_name_in_lang_list):
                print(drug_name_list[i] + ' <--> ' + drug_name_in_lang_list[i])
            else:
                print(drug_name_list[i] + ' <--> ')
    else:
        print("drug name in lang.xlsx list length is great than drug name list length\n")
        for i in range(0, len(drug_name_in_lang_list)):
            if i < len(drug_name_list):
                print(drug_name_list[i] + ' <--> ' + drug_name_in_lang_list[i])
            else:
                print(' <--> ' + drug_name_in_lang_list[i] )

def getDrugLibStringIdList(list):
    col_list = []
    drug_name_in_lang_list = []
    
    # get language col index
    language_index_col = getTitleColIndex(lang_filename, lang_sheet, lang_sheet_title_english)
    string_id_col = getTitleColIndex(lang_filename, lang_sheet, lang_sheet_title_string_id)
    # open lang.xlsx
    sheet = getExcleSheet(lang_filename, lang_sheet)
    # find the same drug name and store the drug name string id
    for row in range(1, sheet.nrows):
        lang_cell_value = sheet.cell_value(row, language_index_col)
        if lang_cell_value in list:
            str_id = 'E_STR_' + sheet.cell_value(row, string_id_col)
            col_list.append(str_id)
            drug_name_in_lang_list.append(lang_cell_value)
    # check length of the drug name list is equal to the length of the drug name in lang.xlsx list
    if len(drug_name_in_lang_list) != len(list):
        printStringIdListAndDrugNameListDiff(drug_name_in_lang_list, list)
        exit_p("drug name in lang.xlsx list length is not equal to drug name list length")

    return  col_list

def getDrugLibColList(sheet_name,col_name):
    col = getTitleColIndex(drug_lib_filename, sheet_name, col_name)
    string_list = getExcleColsList(drug_lib_filename, sheet_name, col)
    # print(string_list)
    return string_list

def changeIsDisplayFlagList(sheet_name):
    string_list = []
    drug_is_display_list = getDrugLibColList(sheet_name, drug_items_sheet_title_is_display)
    # print(drug_is_display_list)
    for value in drug_is_display_list:
        if value.upper() == 'YES' or value.upper() == 'Y':
            string_list.append('DRUG_FLAG_DISPLAY')
        else:
            string_list.append('DRUG_FLAG_NOT_DISPLAY')

    return string_list

def getTipsIdList(sheet_name):
    string_list = []
    drug_tips_list = getDrugLibColList(sheet_name, drug_items_sheet_title_tips)
    # print(drug_tips_list)
    for value in drug_tips_list:
        if value != '':
            print("tips is not empty")
        string_list.append('E_STR_NULL')
    return string_list

def getDrugColorList(sheet_name):
    color_list = ['blue','green','red','cyan','magenta','yellow','lightblue','lightgreen','lightred','lightcyan','lightmagenta','lightyellow',
                  'darkblue','darkgreen','darkred','darkcyan','darkmagenta','darkyellow','white','lightgray','gray','darkgray',
                  'black','brown','orange','transparent']
    string_list = []
    drug_color_list = getDrugLibColList(sheet_name, drug_items_sheet_title_color)
    # print(drug_color_list)
    for value in drug_color_list:
        if value in color_list:
            string_list.append('GUI_' + value.upper())
        else:
            print(color_list)
            exit_p("%s color is not support, please use above color" % value)
    return string_list

def getDrugItemIndex(sheet, index_col, name_col, item):
    drug_index = -1
    # find the same drug name and store the drug name string id
    for row in range(1, sheet.nrows):
        cell_value = sheet.cell_value(row, name_col)
        if cell_value == item:
            # get the drug index
            drug_index = sheet.cell_value(row, index_col)
            return drug_index

    if item == 'null':
        drug_index = 0

    if drug_index == -1:
        exit_p("%s does not exist" % item)

    return drug_index

def getDrugItemsIndex(list):
    # get language col index
    drug_name_col = getTitleColIndex(drug_lib_filename, drug_items_sheet, drug_items_sheet_title_english)
    drug_index_col = getTitleColIndex(drug_lib_filename, drug_items_sheet, drug_items_sheet_title_index)
    # open default_drug_lib.xlsx
    sheet = getExcleSheet(drug_lib_filename, drug_items_sheet)

    drug_items_index_list = []
    for value in list:
        # print(value)
        drug_item_list = value.split(',\n')
        # print(drug_item_list)
        for item in drug_item_list:
            # get drug index from drug_items sheet
            drug_item_index = getDrugItemIndex(sheet, drug_index_col, drug_name_col, item)
            index = (int)(drug_item_index)
            # print(index)
            drug_items_index_list.append(index)

    # print(drug_items_index_list)
    return drug_items_index_list

def getDrugItemsNameByIndex(list,language):
    drug_items_name_title = "drug name/" + language
    # get language col index
    drug_name_col = getTitleColIndex(drug_lib_filename, drug_items_sheet, drug_items_name_title)
    drug_index_col = getTitleColIndex(drug_lib_filename, drug_items_sheet, drug_items_sheet_title_index)
    # open default_drug_lib.xlsx
    sheet = getExcleSheet(drug_lib_filename, drug_items_sheet)

    drug_items_index_list = []

    for item in list:
        # get drug index from drug_items sheet
        drug_item_index = getDrugItemIndex(sheet, drug_index_col, drug_name_col, item)
        index = (int)(drug_item_index)
        # print(index)
        drug_items_index_list.append(index)

    # print(drug_items_index_list)
    return drug_items_index_list

def getDrugItemIndexStr(drug_item_cnt, drug_items_index_list, language):
    string = ""
    count = (int)(drug_item_cnt)
    # Commonly Used Drug
    if count == 0:
        count = 1

    for i in range(0, count):
        # always pop first item
        value = drug_items_index_list.pop(0)
        # print(value)
        # check last line
        if i == (count-1):
            string += (str)(value)
        else:
            string += (str)(value) + ','

    return string

def getDrugItemSortedIndexStr(drug_item_cnt, drug_items_index_list, language):
    string = ""
    count = (int)(drug_item_cnt)
    # Commonly Used Drug
    if count == 0:
        count = 1

    drug_names_in_categroy_list = []
    # open default_drug_lib.xlsx
    sheet = getExcleSheet(drug_lib_filename, drug_items_sheet)
    drug_items_name_title = "drug name/" + language
    drug_name_col = getTitleColIndex(drug_lib_filename, drug_items_sheet, drug_items_name_title)

    for i in range(0, count):
        # always pop first item
        value = drug_items_index_list.pop(0)
        drug_name = sheet.cell_value(value+1, drug_name_col)
        drug_names_in_categroy_list.append(drug_name)
        # print(value)
        # check last line
        if i == (count-1):
            string += (str)(value)
        else:
            string += (str)(value) + ','

    # Sort by language.
    if language == "simplified_chinese":
        drug_names_in_categroy_list.sort(key=lambda char: lazy_pinyin(char)[0][0])
        # print(drug_names_in_categroy_list)
        drug_names_sorted_in_categroy_list = drug_names_in_categroy_list
    else:
        drug_names_sorted_in_categroy_list = sorted(drug_names_in_categroy_list, key=str.lower)

    print(language + ":")
    print(drug_names_sorted_in_categroy_list)

    # Get sorted index list
    drug_items_sorted_index_list = getDrugItemsNameByIndex(drug_names_sorted_in_categroy_list, language)
    # print(drug_items_sorted_index_list)

    # change index to string
    sorted_string = ""
    for i in range(0, count):
        # always pop first item
        value = drug_items_sorted_index_list.pop(0)
        # print(value)
        # check last line
        if i == (count - 1):
            sorted_string += (str)(value)
        else:
            sorted_string += (str)(value) + ','

    # print("string:" + string)
    # print("sorted_string:" + sorted_string + '\n')

    return sorted_string

def createDrugLibItem():
    contens = ('static const TS_DEFAULT_DRUG_ITEM default_drug_items[] = \n{\n')

    # get drug index list in drug_lib.xlsx
    drug_index_list = getDrugLibColList(drug_items_sheet, drug_items_sheet_title_index)
    # get drug name list in drug_lib.xlsx and get to drug string id list in lang.xlsx
    drug_name_list = getDrugLibColList(drug_items_sheet, drug_items_sheet_title_english)
    drug_string_id_list = getDrugLibStringIdList(drug_name_list)
    # print(drug_string_id_list)
    drug_is_display_flag_list = changeIsDisplayFlagList(drug_items_sheet)
    # print(drug_is_display_flag_list)
    drug_tips_id_list = getTipsIdList(drug_items_sheet)
    # print(drug_tips_change_list)
    drug_dose_lower_limition_list = getDrugLibColList(drug_items_sheet, drug_items_sheet_title_dose_lower)
    # print(drug_dose_lower_limition_list)
    drug_dose_upper_limition_list = getDrugLibColList(drug_items_sheet, drug_items_sheet_title_dose_upper)
    # print(drug_dose_upper_limition_list)
    drug_concentration_lower_limition_list = getDrugLibColList(drug_items_sheet, drug_items_sheet_title_concentration_lower)
    # print(drug_concentration_lower_limition_list)
    drug_concentration_upper_limition_list = getDrugLibColList(drug_items_sheet, drug_items_sheet_title_concentration_uppwer)
    # print(drug_concentration_upper_limition_list)
    drug_color_list = getDrugColorList(drug_items_sheet)
    # print(drug_color_list)

    # check list length is equal
    sheet = getExcleSheet(drug_lib_filename, drug_items_sheet)
    for row in range(1, sheet.nrows):
        contens += ('    {' + (str)((int)(drug_index_list[row-1])).ljust(5) + ',' + (drug_is_display_flag_list[row-1]).ljust(24) + ','
                    + (drug_string_id_list[row-1]).ljust(26) + ',' + (drug_tips_id_list[row-1]).ljust(12) + ','
                    + (drug_dose_lower_limition_list[row - 1]).ljust(6) + ',' + (drug_dose_upper_limition_list[row - 1]).ljust(6) + ','
                    + (drug_concentration_lower_limition_list[row - 1]).ljust(6) + ',' + (drug_concentration_upper_limition_list[row - 1]).ljust(6) + ','
                    + (drug_color_list[row - 1])  + '},\n')

    contens += ('};\n\n')
    # print(contens)
    return contens

def createDrugLibCategory(language):
    if language == "simplified_chinese" or language == "english" or language == "spanish" or language == "portuguese":
        if language == "spanish" or language == "portuguese":
            contens = ('#ifdef SUPPORT_' + language.upper() + '\n')
        else:
            contens = ('')

        contens += ('static const TS_DEFAULT_DRUG_CATEGORY default_drug_categories_'+language+'[MAX_DRUG_CATEGORY_NUMBER] = \n{\n')

        # get drug index list in drug_lib.xlsx
        drug_category_index_list = getDrugLibColList(drug_category_sheet, drug_category_sheet_title_index)
        # print(drug_category_index_list)
        # get drug name list in drug_lib.xlsx
        drug_category_name_list = getDrugLibColList(drug_category_sheet, drug_category_sheet_title_english)
        # print(drug_category_name_list)
        drug_category_string_name_list = getDrugLibStringIdList(drug_category_name_list)
        # print(drug_category_string_name_list)
        drug_item_cnt_list = getDrugLibColList(drug_category_sheet, drug_category_sheet_title_count)
        # print(drug_item_cnt_list)
        drug_items_list = getDrugLibColList(drug_category_sheet, drug_category_sheet_title_items_index)
        # print(drug_items_list)
        drug_items_index_list = getDrugItemsIndex(drug_items_list)
        # print(drug_items_index_list)
        drug_category_color_list = getDrugColorList(drug_category_sheet)
        # print(drug_category_color_list)

        sheet = getExcleSheet(drug_lib_filename, drug_category_sheet)
        for row in range(1, sheet.nrows):
            contens += ('    {' + (str)((int)(drug_category_index_list[row-1])).ljust(3)  + ',' + (drug_category_string_name_list[row-1]).ljust(26) + ','
                        + (str)((int)(drug_item_cnt_list[row - 1])).ljust(3) + ',' + '{' + getDrugItemSortedIndexStr((drug_item_cnt_list[row - 1]), drug_items_index_list, language)+ '},  '
                        + (str)(drug_category_color_list[row - 1]) + '},\n')
        contens += ('};\n')

        if language == "spanish" or language == "portuguese":
            contens += ('#endif\n\n')
        else:
            contens += ('\n')
        # print(contens)
        return contens
    else:
        exit_p("language: %s does not exist" % language)
def getDrugLibVersion():
    sheet = getExcleSheet(drug_lib_filename, version_sheet)
    content  = ('#define DRUG_LIB_MAJOR_VERSION' + '       ' + (str)((int)(sheet.cell_value(version_major_row, version_major_col))) + '\n')
    content += ('#define DRUG_LIB_MINOR_VERSION' + '       ' + (str)((int)(sheet.cell_value(version_minor_row, version_minor_col))) + '\n')
    content += ('#define DRUG_LIB_REVISION_VERSION' + '    ' + (str)((int)(sheet.cell_value(version_revision_row, version_revision_col))) + '\n')
    content += ('\n\n')
    
    # print(content)
    return content

def getDrugLibContents():
    contens = ('#ifndef GENERATED_DEFAULT_DRUG_LIB_H\n#define GENERATED_DEFAULT_DRUG_LIB_H\n')
    contens += ('//generated by ' + self_name + ', do not edit it\n\n')
    contens += getDrugLibVersion()
    contens += createDrugLibItem()
    contens += createDrugLibCategory("simplified_chinese")
    contens += createDrugLibCategory("english")
    contens += createDrugLibCategory("spanish")
    contens += createDrugLibCategory("portuguese")
    contens += ('#endif\n')

    return contens
# generated default drug lib file
def createDefaultDrugLibFile():
    default_drug_lib_filename = 'generated_default_drug_lib.h'
    fh = open(default_drug_lib_filename, mode='w', encoding='utf-8')
    contens = getDrugLibContents()
    fh.write(contens)
    fh.close()

# create default drug lib file
createDefaultDrugLibFile()

# default_drug_lib_filename = 'generated_default_drug_lib.h'
# if os.path.exists(default_drug_lib_filename):
#     dst_filename = proj_path + '/Source/kernel/resource/' + default_drug_lib_filename
#     dst_path = proj_path + '/Source/kernel/resource'
#     # print(dst_filename)
#     if os.path.exists(dst_filename):
#         print("%s is exsit, and remove it" %(default_drug_lib_filename))
#         os.remove(dst_filename)
#     # move generated_dafault_drug_lib.h to $PROJ/Source/kernel/resource/generated_dafault_drug_lib.h
#     shutil.move(default_drug_lib_filename, dst_path)
#     print("move generated_dafault_drug_lib.h to $PROJ/Source/kernel/resource/generated_dafault_drug_lib.h")
# else:
#     print("%s is not exsit" %(default_drug_lib_filename))

