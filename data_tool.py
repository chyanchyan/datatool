import os
import sys

try:
    import pandas as pd
except ImportError:
    os.system('pip install pandas -i https://pypi.tuna.tsinghua.edu.cn/simple some-package')
    import pandas as pd

try:
    import numpy as np
except ImportError:
    os.system('pip install numpy -i https://pypi.tuna.tsinghua.edu.cn/simple some-package')
    import numpy as np

# try:
#     import openpyxl
# except ImportError:
#     os.system('pip install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple some-package')
#     import openpyxl
#
# import math


def get_file_name_info(s):
    l = s.split('.')
    post_fix = l[-1]
    if '_' in s:
        index = l[-2].split('_')[-1]
        try:
            index = str(int(index))
        except ValueError:
            filename = ''.join(l[:-1])
            index = ''
        else:
            filename = ''.join(s.split('_')[:-1])
    else:
        filename = ''.join(l[:-1])
        index = ''

    return filename, index, post_fix


def next_file_index(s):
    filename, index, post_fix = get_file_name_info(s)
    if index == '':
        return filename + '_0.' + post_fix
    else:
        return filename + '_' + str(int(index) + 1) + '.' + post_fix

def main():
    res = []
    data_valid_flag = True
    simple_logic_mapping = {
        '求和': sum,
        '均值': np.mean,
        '中位数': np.median,
        '计数': len
    }

    config_path = 'data_tool_config.xlsx'

    data_sources = pd.DataFrame([])
    logics = pd.DataFrame([])
    data = dict()

    try:
        data_sources = pd.read_excel(config_path,
                                     sheet_name='data_sources',
                                     index_col=False,
                                     header=0,
                                     dtype=str)
        logics = pd.read_excel(config_path,
                               sheet_name='logics',
                               index_col=False,
                               header=0,
                               dtype=str)
        pd.DataFrame.dropna(data_sources, inplace=True)
    except FileNotFoundError:
        res.append(('sys_error', '设置文件路径不存在'))
        data_valid_flag = False

    if data_valid_flag:

        for index, row in data_sources.iterrows():
            try:
                data[str(row['data_id'])] = pd.read_csv(str(row['path']))
            except FileNotFoundError:
                res.append((str(row['data_id']), '%s 数据路径不存在' % row['path']))
                data_valid_flag = False

    if data_valid_flag:
        for index, logic in logics.iterrows():
            if not pd.isna(logic['计算逻辑']):
                re = simple_logic_mapping[logic['计算逻辑']](
                    data[str(logic['数据源'])][logic['数据列']]
                )
            else:
                if not pd.isna(logic['advance_mode']):
                    re = eval(logic['advance_mode'])
                else:
                    re = 'no logic rules applied'

            if not pd.isna(logic['输出格式']):
                re = format(re, str(logic['输出格式']))
            else:
                pass
            print(logic['项目'] + str(re))
            res.append((logic['项目'] + ':', re))

    res = pd.DataFrame(res)
    re_filename = '../helpers/data_tool/results.xlsx'
    i = 0
    while os.path.exists(sys.path[0] + '/' + re_filename):
        re_filename = next_file_index(re_filename)
        print(re_filename)
    res.to_excel(re_filename, index=False, header=['计算项目', '计算结果'])
    # book = openpyxl.load_workbook(config_path)
    # writer = pd.ExcelWriter(config_path, engine=openpyxl)
    # writer.book = book
    # writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    # res.to_excel(writer, sheet_name='results', index=False)
    # writer.save()




if __name__ == '__main__':
    main()

