import openpyxl
import os
import urllib.request
import urllib.error
import datetime
import time
import shutil
import re
from paddleocr import PaddleOCR

data_dic = {}
date = datetime.date.today()
main_save_path = '.\\{}{:0>2d}{:0>2d}信息xs2101'.format(date.year, date.month, date.day)
save_path = [main_save_path + '\\健康码', main_save_path + '\\行程码', main_save_path + '\\同行密接人员自查']
check_path = '.\\temp'
temp_path = [check_path + '\\health_code', check_path + '\\tra_code', check_path + '\\close_check']
ID = r'^([1-9]\d{5}[12]\d{3}(0[1-9]|1[012])(0[1-9]|[12][0-9]|3[01])\d{3}[0-9xX])$'
file_name = '{:0>2d}月{:0>2d}日信息xs2101“两码一查询”（收集结果）.xlsx'.format(date.month, date.day)


def save_data():
    main_book = openpyxl.load_workbook(file_name)
    main_sheet = main_book.active
    for i in range(2, main_sheet.max_row):
        if main_sheet.cell(i, 1).value is not None:
            name = main_sheet.cell(i, 3).value
            url = []
            for j in range(4, main_sheet.max_column):
                if main_sheet.cell(i, j).value is not None:
                    url.append(main_sheet.cell(i, j).hyperlink.target)
            data_dic.update({name: [url,
                                    [str(time.time()).replace('.', '_') + str(i) + '.jpeg'
                                     for i in range(3)]]})


def download_img():
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/94.0.4606.61 Safari/537.36 Edg/94.0.992.31"}
    for name in data_dic:
        index = 0
        for url in data_dic.get(name)[0]:
            request = urllib.request.Request(url, headers=head)
            try:
                response = urllib.request.urlopen(request)
                img_name = name + '.jpeg'
                if response.getcode() == 200 and not os.path.exists(save_path[index] + '\\' + img_name):
                    with open(save_path[index] + '\\' + img_name, 'wb') as file:
                        file.write(response.read())
                    # with open(temp_path[index] + '\\' + data_dic.get(name)[1][index], 'wb') as f:
                    #     f.write(response.read())
                shutil.copyfile(save_path[index] + '\\' + img_name,
                                temp_path[index] + '\\' + data_dic.get(name)[1][index])
            except urllib.error.URLError as e:
                if hasattr(e, 'code'):
                    print(e.code)
                if hasattr(e, 'reason'):
                    print(e.reason)
                print('failed!')
            index += 1
    print('搞定了')


def show_profile():
    print('这是一个临时用于打包打卡数据的python代码')
    print('----------------------------------')
    print('使用前请确保与 .py/.exe 文件同一目录下\n有且仅有一个 收集数据 .xlsx 文件')
    print('一定要使用腾讯文档导出的 .xlsx 文件')
    print('请严格按照QQ群中发送的收集表顺序填写')
    print('遇到bug，及时联系开发人员，比如对面宿舍')
    print('-----------------------------------')
    print('详细阅读后,键入y开始打包,n退出,其他无效')


def check():
    print('是否对生成文件进行身份证检查？\ny:是\t其他:否')
    is_stop = False
    problem = {}
    che = []
    while not is_stop:
        choice = str(input())
        if choice == 'y':
            for name_t in data_dic:
                error = []
                index = 0
                for path in data_dic.get(name_t)[1][::2]:
                    flag = False
                    print(path)
                    # reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)
                    # result = reader.readtext(temp_path[index] + '\\' + path)
                    ocr = PaddleOCR(use_angle_cls=True, lang='ch')
                    result = ocr.ocr(temp_path[index] + '\\' + path, cls=True)
                    for item in result:
                        if len(re.findall(ID, str(item[1][0]))) > 0:
                            flag = True
                    if not flag:
                        error.append(save_path[index].replace(main_save_path + '\\', ''))
                    index += 2
                if len(error) > 0:
                    problem.update({name_t: error})
            # for name_t in problem:
            #     index = 0
            #     for path in data_dic.get(name_t)[1][::2]:
            #         flag = []
            #         print(path)
            #         # reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)
            #         # result = reader.readtext(temp_path[index] + '\\' + path)
            #         ocr = PaddleOCR(use_angle_cls=True, lang='ch', use_gpu=True)
            #         result = ocr.ocr(temp_path[index] + '\\' + path, cls=True)
            #         for item in result:
            #             print(item)
            #             if len(re.findall(ID, str(item[1][0]))) > 0:
            #                 flag.append(True)
            #         if flag.count(True) == 2:
            #             che.append(name_t)
            #         index += 2
            is_stop = True
        else:
            is_stop = True
    print('------------------------------------------------------')
    print(problem)
    print('以上同学可能存在问题，请查看！')


if __name__ == '__main__':
    try:
        is_exit = False
        while not is_exit:
            show_profile()
            select = str(input())
            if select == 'y':
                if not os.path.exists(main_save_path):
                    os.mkdir(main_save_path)
                    for path in save_path:
                        os.mkdir(path)
                if not os.path.exists(check_path):
                    os.mkdir(check_path)
                    for temp in temp_path:
                        os.mkdir(temp)
                save_data()
                download_img()
                check()
                is_exit = True
            elif select == 'n':
                is_exit = True
    finally:
        for root, dirs, files in os.walk(check_path, topdown=False):
            for name in files:
                os.remove(os.path.join(root, name))
            for name in dirs:
                os.rmdir(os.path.join(root, name))
        os.rmdir(check_path)
