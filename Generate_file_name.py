#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# Author: Fei.LONG

import os
import xlwt


def get_c_file_name(file_dir):
    '''
    :function:get *.c file from file_dir
    :param file_dir:
    :return: c_file_list
    '''
    c_file_list = []
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            # os.path.splitext()函数将路径拆分为文件名+扩展名
            if os.path.splitext(file)[1] == '.c':
                c_file_list.append(os.path.join(file))
    return c_file_list


def get_h_file_name(file_dir):
    '''
    :function:get *.h file from file_dir
    :param file_dir:
    :return: h_file_list
    '''
    h_file_list = []
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if os.path.splitext(file)[1] == '.h':
                h_file_list.append(os.path.join(file))
    return h_file_list


def write_c_file_to_excel(wb_sheet, file_dir):
    c_file_list = get_c_file_name(file_dir)
    for row in range(0, len(c_file_list)):
        wb_sheet.write(row + 1, 0, c_file_list[row])


def write_h_file_to_excel(wb_sheet, file_dir):
    h_file_list = get_h_file_name(file_dir)
    for row in range(0, len(h_file_list)):
        wb_sheet.write(row + 1, 1, h_file_list[row])


def write_to_excel(file_dir):
    '''
    :function: write data to excel
    :param file_dir:
    :return: None
    '''
    # define file name
    file_name = 'C9_Source_File_list.xls'
    # create a workbook
    wb = xlwt.Workbook(encoding='utf-8')

    # UI_Bootloader file dir
    ui_bootloader_file_dir = file_dir + r'\UIBootloader\Source'
    # UI_Firmware file dir
    ui_firmeware_file_dir = file_dir + r'\UIFirmware\Source'
    # MAIN_Bootloader file dir
    main_bootloader_file_dir = file_dir + r'\MainBootloader\Source'
    # MAIN_Firmware file dir
    main_firmware_file_dir = file_dir + r'\MainFirmware\Source'
    # Safety_Firmware file dir
    safety_firmware_file_dir = file_dir + r'\SafetyFirmware\Source'

    # create sheets
    wb_sheet_ui_bootloader = wb.add_sheet('UI_Bootloader', cell_overwrite_ok=True)
    wb_sheet_ui_firmware = wb.add_sheet('UI_Firmware', cell_overwrite_ok=True)
    wb_sheet_main_bootloader = wb.add_sheet('MAIN_Bootloader', cell_overwrite_ok=True)
    wb_sheet_main_firmware = wb.add_sheet('MAIN_Firmware', cell_overwrite_ok=True)
    wb_sheet_safety_firmware = wb.add_sheet('SAFETY_Firmware', cell_overwrite_ok=True)

    # set style
    style = xlwt.XFStyle()  # Init style
    font = xlwt.Font()  # create font for style
    font.bold = True  # bold style
    style.font = font  # set style

    # write *.c file name into sheet
    wb_sheet_ui_bootloader.write(0, 0, '.c File', style)  # row 0,col 0
    wb_sheet_ui_firmware.write(0, 0, '.c File', style)
    wb_sheet_main_bootloader.write(0, 0, '.c File', style)
    wb_sheet_main_firmware.write(0, 0, '.c File', style)
    wb_sheet_safety_firmware.write(0, 0, '.c File', style)

    write_c_file_to_excel(wb_sheet_ui_bootloader, ui_bootloader_file_dir)
    write_c_file_to_excel(wb_sheet_ui_firmware, ui_firmeware_file_dir)
    write_c_file_to_excel(wb_sheet_main_bootloader, main_bootloader_file_dir)
    write_c_file_to_excel(wb_sheet_main_firmware, main_firmware_file_dir)
    write_c_file_to_excel(wb_sheet_safety_firmware, safety_firmware_file_dir)

    # write *.h file name into sheet
    wb_sheet_ui_bootloader.write(0, 1, '.h File', style)  # row 0,col 1
    wb_sheet_ui_firmware.write(0, 1, '.h File', style)
    wb_sheet_main_bootloader.write(0, 1, '.h File', style)
    wb_sheet_main_firmware.write(0, 1, '.h File', style)
    wb_sheet_safety_firmware.write(0, 1, '.h File', style)

    write_h_file_to_excel(wb_sheet_ui_bootloader, ui_bootloader_file_dir)
    write_h_file_to_excel(wb_sheet_ui_firmware, ui_firmeware_file_dir)
    write_h_file_to_excel(wb_sheet_main_bootloader, main_bootloader_file_dir)
    write_h_file_to_excel(wb_sheet_main_firmware, main_firmware_file_dir)
    write_h_file_to_excel(wb_sheet_safety_firmware, safety_firmware_file_dir)

    # save work book
    wb.save(file_name)


if __name__ == '__main__':
    # set the file_dir which needs to be written to excel as the argument for write_to_excel function.

    # manual generate by absolute path

    # write_to_excel(r'C:\01-WorkSpace\SourceCode\C9\CI_For_PV')

    # auto generate by cmd on window
    '''
    usage : 
    1> copy this .py file to the project dir and open cmd on window in current dir
    2> type cmd > Generate_file_name      
    '''
    write_to_excel(os.getcwd())
    os.system("pause")
