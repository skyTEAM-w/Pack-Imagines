import openpyxl
import os
import urllib.request

file_name = '打卡情况.xlsx'


def download(img_url):
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/94.0.4606.61 Safari/537.36 Edg/94.0.992.31"}
    request = urllib.request.Request(img_url, headers=head)
    try:
        response = urllib.request.urlopen(request)
        img_name = 'dd.png'
        path = ".\\" + img_name
        if response.getcode() == 200:
            with open(path, 'wb') as f:
                f.write(response.read())
            return path
    except:
        return 'failed'


if __name__ == '__main__':
    main_book = openpyxl.load_workbook('test（收集结果）_20220224.xlsx')
    main_sheet = main_book.active
    print(type(main_sheet.cell(2, 4).hyperlink.target))
    download("https://docimg10.docs.qq.com/image/OWINGk49ZusyQD2pdKNyog.jpeg?w=1152&h=2376&_type=jpeg")
