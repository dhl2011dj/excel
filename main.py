# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import xlrd2
import xlwt

def fomart_number(num):
    length = len(str(num))
    if 1 == length:
        sheet_name = '000' + str(num)
    elif 2 == length:
        sheet_name = '00' + str(num)
    elif 3 == length:
        sheet_name = '0' + str(num)
    elif 4 == length:
        sheet_name = str(num)
    else:
        exit(10)
    return sheet_name

def parse_write_xlsx(start, end, begin_row_index, begin_col_index):
    #
    # 准备数据结构：字典嵌套字典嵌套字典
    #

    #
    # 解析表格
    #
    population = {'Australia': 2000,
                  'Brazil': 2000,
                  'China': 2000,
                  'Great Britain': 2000,
                  'India': 2000,
                  'Japan': 2035,
                  'South Africa': 2000,
                  'USA': 2000}

    print("From sheet ", start, " to ", end)  # Press Ctrl+F8 to toggle the breakpoint.
    wb = xlrd2.open_workbook(
        '/home/hanlin/Downloads/20-065644-01_University_of_Cambridge_All_countries_wt__Internal_Client_remove106-109and56.xlsx')

    data = {}  # 空字典
    for isheet in range(start, end + 1):
        sheet_name = fomart_number(isheet)

        if 0 != wb.sheet_names().count(sheet_name):
            print("$$$$$$$$$$$$$ 表单[", isheet, "]解析完成 $$$$$$$$$$$$$")
            sheet = wb.sheet_by_name(sheet_name)
            # ----------------  检查国家和人数 ------------------
            if ("Australia" != str(sheet.cell_value(7, 2)).strip(" ") or
                    "Brazil" != str(sheet.cell_value(7, 3)).strip(" ") or
                    "China" != str(sheet.cell_value(7, 4)).strip(" ") or
                    "Great Britain" != str(sheet.cell_value(7, 5)).strip(" ") or
                    "India" != str(sheet.cell_value(7, 6)).strip(" ") or
                    "Japan" != str(sheet.cell_value(7, 7)).strip(" ") or
                    "South Africa" != str(sheet.cell_value(7, 8)).strip(" ") or
                    "USA" != str(sheet.cell_value(7, 9)).strip(" ")):
                print(" > ", " 国家名断言错误")
                exit(100)
            if (2000 != int(sheet.cell_value(9, 2)) or
                    2000 != int(sheet.cell_value(9, 3)) or
                    (2000 != int(sheet.cell_value(9, 4)) and 0 != int(sheet.cell_value(9, 4))) or
                    2000 != int(sheet.cell_value(9, 5)) or
                    2000 != int(sheet.cell_value(9, 6)) or
                    2035 != int(sheet.cell_value(9, 7)) or
                    2000 != int(sheet.cell_value(9, 8)) or
                    2000 != int(sheet.cell_value(9, 9))):
                print(" > ", " 人数断言错误")
                exit(101)
            # ----------------  检查国家和人数 ------------------

            # ----------------  将问卷数据写入数据结构 ------------------
            sheet_data = {}  # 存储一个国家对每个问题的数据<country, country_data>
            for icol in range(begin_col_index - 1, sheet.ncols):
                country_data = {}  # <option index, <option, count>>
                for irow in range(begin_row_index - 1, sheet.nrows):
                    if "" != str(sheet.cell_value(irow, 0)).strip(" "):  # 如果这一行第一列有内容
                        if ("-" == str(sheet.cell_value(irow, icol)).strip(" ") or
                                "N.A" == str(sheet.cell_value(irow, icol)).strip(" ")):
                            country_name = sheet.cell_value(7, icol)
                            option = int(sheet.cell_value(irow, 0))
                            count = "N.A."
                        else:
                            country_name = sheet.cell_value(7, icol)
                            option = int(sheet.cell_value(irow, 0))
                            count = int(sheet.cell_value(irow, icol))

                        country_data[option] = count
                        sheet_data[country_name] = country_data
            data[isheet] = sheet_data
            # ----------------  将问卷数据写入数据结构 ------------------
    # print(data)

    #
    #   写入表格
    #
    out_xlsx = xlwt.Workbook()
    sheet = out_xlsx.add_sheet('data')
    # 写入最左侧一列
    # for irow in range(0, 2000):
    #     sheet.write(irow,0,'我最帅')

    # 写入数据
    icol = 0
    for sheet_write_keys in data.keys():
        sheet_write_data = data[sheet_write_keys]
        print("column = ", icol)
        icol += 1
        sheet.write(0, icol, fomart_number(sheet_write_keys)+"号表单")
        start_row = 1
        incremental = 0
        for country_write_name in sheet_write_data.keys():
            country_write_data = sheet_write_data[country_write_name]
            for option in country_write_data.keys():
                # print(option, "-", country_write_data[option])
                if "N.A." == country_write_data[option]:
                    incremental = 0
                else:
                    incremental = int(country_write_data[option])

                if "China" == country_write_name and "N.A." == country_write_data[option]:
                    incremental = 2000
                    write_flag = "CHINA_NA"
                else:
                    write_flag = "normal"

                print("start_row = ", start_row, "incremental = ", incremental, "write_flag: ", write_flag)

                for i in range(start_row, start_row+incremental):
                    if "normal" == write_flag:
                        sheet.write(i, icol, option)
                    elif "CHINA_NA" == write_flag:
                        sheet.write(i, icol, "N.A.")
                    # if i <= 16035:
                    #     sheet.write(i, icol, option)

                start_row += incremental
                if "CHINA_NA" == write_flag:
                    break;

    out_xlsx.save('output.xlsx')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    parse_write_xlsx(28, 128, 14, 3)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
