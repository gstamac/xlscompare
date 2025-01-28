import sys
import time
import re
import warnings
import io
import argparse

import pandas as pd
import numpy as np
import win32com.client as win32
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning, module=re.escape('openpyxl.styles.stylesheet'))

out_file = False

def main():
    global out_file
    start = time.time()

    parser = argparse.ArgumentParser(prog="compare", description="Compare Excel files")
    parser.add_argument('path1', help='first file or folder')
    parser.add_argument('path2', help='second file or folder')
    parser.add_argument('-o', '--outfile', help='output file')
    parser.add_argument('-s', '--skipunlock', action='store_true', help="skip unlocking Excel files")
    parser.print_help()
    args = parser.parse_args()

    xls_file_path_1 = Path(args.path1)
    xls_file_path_2 = Path(args.path2)
    if args.outfile:
        out_file = io.open(args.outfile, "w", encoding="utf-8", buffering=1)

    # xls_file_path_1 = Path(sys.argv[1])
    # xls_file_path_2 = Path(sys.argv[2])
    # if len(sys.argv) > 3:
    #     out_file = io.open(sys.argv[3], "w", encoding="utf-8", buffering=1)

    if xls_file_path_1.is_dir():
        compare_dirs(xls_file_path_1, xls_file_path_2, out_file)
    else:
        compare_files(xls_file_path_1, xls_file_path_2, True, out_file)

    end = time.time()

    out("DONE in %s seconds" % round(end - start))

    if out_file:
        out_file.close()

def out(s):
    global out_file
    print(s)
    if out_file:
        out_file.write(s + "\n")

def compare_dirs(dir_1, dir_2, full_comp):
    if not dir_1.exists():
        out("ERROR: Directory %s doesn't exist" % dir_2)
        return
    if not dir_2.exists():
        out("ERROR: Directory %s doesn't exist" % dir_2)
        return

    visited = []

    for file_path in dir_1.iterdir():
        visited.append(file_path.name)
        compare_dir_or_file(file_path, dir_2 / file_path.name, full_comp)
    for file_path in dir_2.iterdir():
        if not file_path.name in visited:
            compare_dir_or_file(dir_1 / file_path.name, file_path, full_comp)

    return

def compare_dir_or_file(file_path_1, file_path_2, full_comp):
    if file_path_1.is_dir():
        compare_dirs(file_path_1, file_path_2, full_comp)
    elif file_path_1.name.endswith(".xlsx") or file_path_1.name.endswith(".xls"):
        try:
            compare_files(file_path_1, file_path_2, False, full_comp)
        except Exception as error:
            out("FATAL ERROR: %s" % error)

def compare_files(xls_file_path_1, xls_file_path_2, cell_comp, full_comp):
    out("Comparing %s" % xls_file_path_1)
    out("      and %s" % xls_file_path_2)

    if not xls_file_path_1.exists():
        out("ERROR: File %s doesn't exist" % xls_file_path_1)
        return
    if not xls_file_path_2.exists():
        out("ERROR: File %s doesn't exist" % xls_file_path_2)
        return

    book1 = read_file(xls_file_path_1)
    book2 = read_file(xls_file_path_2)

    if len(book1) != len(book2):
        out("ERROR: Different number of sheets %d != %d" % (len(book1), len(book2)))

    for sheet_name, sheet1 in book1.items():
        if sheet_name in book2:
            sheet2 = book2[sheet_name]
            if len(sheet1.columns) != len(sheet2.columns):
                out("ERROR: [%s] Different number of columns (%s <> %s)" % (sheet_name, len(sheet1.columns), len(sheet2.columns)))
            if len(sheet1) != len(sheet2):
                out("ERROR: [%s] Different number of rows (%s <> %s)" % (sheet_name, len(sheet1), len(sheet2)))
            else:
                for col in sheet1:
                    column_name = col_num_to_name(col)
                    if col in sheet2:
                        column1 = sheet1[col]
                        column2 = sheet2[col]
                        sum1 = round(column1.map(convert_to_float).sum(), 4)
                        sum2 = round(column2.map(convert_to_float).sum(), 4)
                        if sum1 != sum2:
                            out("ERROR: [%s] Different column %s sum (%s != %s)" % (sheet_name, column_name, sum1, sum2))
                            if cell_comp or full_comp:
                                diffs = 0
                                for row_idx in range(min(len(column1), len(column2))):
                                    if column1[row_idx] == column1[row_idx] and column1[row_idx] != column2[row_idx]:
                                        diffs += 1
                                        out("ERROR: [%s] Different value in %s%d ('%s' <> '%s')" % (sheet_name, column_name, row_idx + 1, column1[row_idx], column2[row_idx]))
                                    if not full_comp and diffs > 3:
                                        break
                    else:
                        out("ERROR: [%s] Column %s (%s) missing" % (sheet_name, column_name, sheet1[col][0]))
        else:
            out("ERROR: %s not found" % sheet_name)

def read_file(file_path):
    # if is_locked(file_path):
    unlock(file_path)

    return pd.read_excel(file_path, index_col=None, header=None, sheet_name=None, dtype=object)

def is_locked(file_path):
    diff = abs(file_path.stat().st_mtime - file_path.stat().st_ctime)

    return diff < 30

def unlock(file_path):
    out("INFO: Unlocking %s" % file_path)
    try:
        xlapp = win32.DispatchEx('Excel.Application')
        xlapp.DisplayAlerts = False
        xlapp.Visible = True
        xlbook = xlapp.Workbooks.Open(file_path)
        xlbook.RefreshAll()
        xlbook.Save()
        xlbook.Close()
        xlapp.Quit()
    except:
        out("FATAL ERROR: Failed to unlock %s" % file_path)

    return

def col_num_to_name(col_num):
    (d, m) = divmod(col_num, 26)
    if d == 0:
        return chr(m + 65)
    return chr(d - 1 + 65) + chr(m + 65)

def float_or_zero(value):
    if isinstance(value, (int, float, complex)) and not isinstance(value, bool):
        return value
    return 0

def convert_to_float(value):
    if isinstance(value, (int, float, complex)) and not isinstance(value, bool):
        return value
    return len(str(value).strip())


main()
