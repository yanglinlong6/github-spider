import requests, os, time
from bs4 import BeautifulSoup
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
excel_titles = ["內文簡介", "網址", "上傳日期", "資料夾名稱"]
ws.append(excel_titles)


def python_spider(start, finish, keyword):
    page = start

    for i in range(finish - start + 1):
        # 頁碼
        ws.append([""])
        ws.append([f"第{str(page)}頁"])

        # 主網頁
        URL = f"https://github.com/search?p={page}&q={keyword}&type=Repositories"
        response = requests.get(URL)
        soup = BeautifulSoup(response.text, 'html.parser')
        tag = "div.mt-n1.flex-auto"
        titles = soup.select(tag)

        for t in titles:
            username_filename = t.select('a.v-align-middle')
            content = t.select('p.mb-1')
            updated = t.select('relative-time')
            address = t.select('a.v-align-middle')

            name = username_filename[0].text.split('/')
            username = name[0]
            filename = name[1]

            try:
                print(content[0].text.strip())  # 內文
            except IndexError:
                print('-None-')

            # print(updated[0].text) #上傳時間
            print("https://github.com/" + address[0]['href'])  # 網址
            print('--------------------------------------------\n')

            # for Excel

            course = []
            # course.append(username)

            try:
                course.append(content[0].text.strip())
            except IndexError:
                course.append('-None-')

            course.append("https://github.com/" + address[0]['href'])
            course.append(updated[0].text)
            course.append(filename)
            ws.append(course)

        print(
            f'------------------------------------- Page {page} - Download complete ----------------------------------------- \n')
        page += 1

        time.sleep(breaktime)


# 抓取資料
try:
    # keyword = "Python" #關鍵字
    keyword = input()  # 關鍵字
    start = 1  # 起始頁
    end = 2  # 終止頁
    breaktime = 10  # 每頁間隔時間

    print(f'Github 關鍵字搜尋: {keyword}')
    python_spider(start, end, keyword)
except KeyboardInterrupt:
    print('Stop')


def mkdir(path):
    isExists = os.path.exists(path)  # 判断路径是否存在，存在则返回true
    if not isExists:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)

        print(path + ' 创建成功')
        return True
    else:
        # 如果目录存在则不创建，并提示目录已存在
        print(path + ' 目录已存在')
        return True


# Excel存檔
path = f'.//{keyword}'
try:
    if mkdir(path):
        os.chdir(path)
        wb.save(f'Github Spider - {keyword}.xlsx')
        print('Finished!')
except PermissionError:
    print('檔案已被開啟，關閉後按enter鍵')
    input('')
    if path:
        os.chdir(path)
        wb.save(f'Github Keyword - {keyword}.xlsx')
        print('Finished!')
