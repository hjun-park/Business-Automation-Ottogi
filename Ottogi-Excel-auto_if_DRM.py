import xlwings as xw
import pandas as pd
import os, sys
import time, datetime

dirpath = "D:\\dev\\excel_result\\test1"
today = datetime.date.today()
day = today.strftime('%d')


def input_column_info():
    while True:
        column_match_list = []
        print("== 빈 값이 입력되면 진행됩니다 ==")
        print("복사할 행 / 붙여넣을 행 순서대로 입력")
        print("예시 : B A")

        while True:
            column_info = list(map(str, input().split()))

            if not column_info:
                break

            column_match_list.append(column_info)

        print("=========== 컬럼 정보 ===========")
        for column_info in column_match_list:
            print(f'[{column_info[0]}] ==> [{column_info[1]}]')

        return column_match_list


def print_filelist(flist):
    for num, file in enumerate(flist):
        print(f'\t\t[{num + 1}] ==> {file}')

    # Ask Continue
    while True:
        ANSWER = input("\n\nContinue ? (Y/N) ")

        if ANSWER == "n" or ANSWER == "N":
            sys.exit(0)
        elif ANSWER != "y" or ANSWER != "Y":
            break
        else:
            continue


def make_filelist(dirpath):
    print("=============== Directory Check ================")
    if os.path.isdir(dirpath) is False:
        print(f'{dirpath} doesn\'t exist. Check it')
        sys.exit(0)
    else:
        print(f"[{dirpath}] OK")

    flist = [os.path.join(dirpath, f) for f in os.listdir(dirpath) if os.path.isfile(os.path.join(dirpath, f))]

    return flist


def find_result_filename():
    temp = input("Saved result file: ")
    result_name = os.path.join(dirpath, temp)

    return result_name


if __name__ == "__main__":
    result_file_name = input("Result file name : ")

    temp_name = dirpath+"\\결과\\"+result_file_name
    print(temp_name)
    result_wb = xw.Book(temp_name)
    result_sht = result_wb.sheets(day)


    file_list = make_filelist(dirpath)
    print_filelist(file_list)

    print("=====================")
    print(file_list)
    print("=====================")

    input_column_list = input_column_info()

    for file in file_list:
        for column_info in input_column_list:
            copy_column = column_info[0]
            paste_column = column_info[1]

            start_row = f'{copy_column}2'
            wb = xw.Book(file)
            sht = wb.sheets[0]

            # num_row = result_wb.sheets[0].range('A' + str(result_wb.sheets[0].cells.last_cell.row)).end('up').row
            # print(f"{num_row} <== num_row")

            # TODO: wb.sheets[0] => sht
            max_copy_row = sht.range(copy_column + str(sht.cells.last_cell.row)).end('up').row
            copy_range = f'{start_row}:{copy_column}{int(max_copy_row)}'
            copied_value = sht[copy_range].options(pd.DataFrame).value

            # print(f'{copied_value} <==== copied_value')

            # result 파일의 경우, 엑셀 row 1번 컬럼에 데이터를 병합해서 채워놓을 것
            # TODO: result_wb.sheets[0] => result_sht
            max_paste_row = result_sht.range(paste_column + str(result_sht.cells.last_cell.row)).end('up').row
            max_paste_row = int(max_paste_row)
            max_paste_row += 1
            max_paste_row = str(max_paste_row)
            max_paste_start = f"{paste_column}{max_paste_row}"
            result_sht.range(max_paste_start).value = copied_value

            result_wb.save()
            wb.close()


