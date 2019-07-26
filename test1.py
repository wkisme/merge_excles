from os import listdir
from os.path import isfile, join
import pandas as pd


# function
def multiple_dfs(df_list1, sheets, file_name, spaces):
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    writer1 = pd.ExcelWriter('index.xlsx', engine='xlsxwriter')
    row = 0
    row1 = 0
    row2 = 0
    for key, dataframe in df_list1.items():
        if row < 1000000:
            index1 = pd.DataFrame({'country': [key], 'locate_row': [row]})
            index1.to_excel(writer1, sheet_name='index', startrow=row1, startcol=0)
            row1 = row1 + len(index1.index) + 1

            insert_row = pd.DataFrame({'NTLC': [key, ''], 'NTLC_NAME': ['', ''], 'REV_CD': ['', '']})

            insert_row.to_excel(writer, sheet_name=sheets, startrow=row, startcol=0)
            row = row + len(insert_row.index) + spaces
            dataframe.to_excel(writer, sheet_name=sheets, startrow=row, startcol=0)
            row = row + len(dataframe.index) + spaces + 1
            print(row)
        else:
            index1 = pd.DataFrame({'country': [key], 'locate_row': [row2]})
            index1.to_excel(writer1, sheet_name='index', startrow=row1, startcol=0)
            row1 = row1 + len(index1.index) + 1

            insert_row = pd.DataFrame({'NTLC': [key, ''], 'NTLC_NAME': ['', ''], 'REV_CD': ['', '']})

            insert_row.to_excel(writer, sheet_name='test_1_2', startrow=row2, startcol=0)
            row2 = row2 + len(insert_row.index) + spaces
            dataframe.to_excel(writer, sheet_name='test_1_2', startrow=row2, startcol=0)
            row2 = row2 + len(dataframe.index) + spaces + 1
            print(row2)

    writer1.save()
    writer.save()


my_path = '/home/wangkuo/PycharmProjects/merge_excels/excel_resources/'

if __name__ == '__main__':
    only_files = [f for f in listdir(my_path) if isfile(join(my_path, f))]
    df_list = {}
    df_exception = list()
    df_exception.append('Canada0.0空值.xlsx')
    df_exception.append('Tunisia0.0空值2015.xlsx')
    df_exception.append('Burkina Faso0.0.xlsx')
    df_exception.append('Costa Rica0.0空值2016.xlsx')
    df_exception.append('China0.0.xlsx')
    for i in only_files:
        if i.endswith('.xlsx'):
            if i not in df_exception:
                print(i)
                data = pd.read_excel(my_path + i, sheet_name=0,
                                     usecols=[0, 1, 2], dtype=str, index_col=None, header=4)
                df_list[i] = data
                # print(data)

    print(len(df_list))
    multiple_dfs(df_list, 'test1_1', 'test1.xlsx', 1)





