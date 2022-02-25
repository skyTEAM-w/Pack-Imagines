import openpyxl
import os
import urllib.request
import urllib.error
import datetime

file_name = '收集结果.xlsx'
data_dic = {}
date = datetime.date.today()
main_save_path = '.\\{}{:0>2d}{:0>2d}信息xs2101'.format(date.year, date.month, date.day)
save_path = [main_save_path + '\\健康码', main_save_path + '\\行程码', main_save_path + '\\同行密接人员自查']


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
            data_dic.update({name: url})


def download_img():
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/94.0.4606.61 Safari/537.36 Edg/94.0.992.31"}
    for name in data_dic:
        index = 0
        for url in data_dic.get(name):
            request = urllib.request.Request(url, headers=head)
            try:
                response = urllib.request.urlopen(request)
                img_name = name + '.jpeg'
                if response.getcode() == 200:
                    with open(save_path[index] + '\\' + img_name, 'wb') as file:
                        file.write(response.read())
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


if __name__ == '__main__':
    is_exit = False
    while not is_exit:
        show_profile()
        select = str(input())
        if select == 'y':
            if not os.path.exists(main_save_path):
                os.mkdir(main_save_path)
                for path in save_path:
                    os.mkdir(path)
            save_data()
            download_img()
        elif select == 'n':
            is_exit = True
