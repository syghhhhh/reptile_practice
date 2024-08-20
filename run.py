# -*- coding: utf-8 -*-
# @Time    : 2024/8/20 14:46
# @Author  : 施昀谷
# @File    : run.py


def get_result_list(html_path, page_num, save_file):
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from bs4 import BeautifulSoup
    from tqdm import tqdm
    import pandas as pd
    try:
        # 初始化一个空列表来存储结果
        results = []
        # 启动chrome驱动 这步需要等待一段时间
        driver = webdriver.Chrome()
        # 进入想要爬取的网页
        driver.get(html_path)
        for i in tqdm(range(page_num)):
            # 通过唯一id检索到想要的元素
            element = driver.find_element(By.ID, "flData")
            # 提取元素信息文本
            html_text = element.get_attribute('innerHTML')
            # 使用BeautifulSoup解析HTML
            soup = BeautifulSoup(html_text, 'html.parser')
            # 查找所有的<tr>标签
            for tr in soup.find_all('tr', class_='list-b'):
                # 提取URL
                url = tr.find('li', class_='l-wen')['onclick'].split('(')[1].split(')')[0]
                # 删去额外的引号
                url = url[1: -1]
                # 提取标题
                title = tr.find('li', class_='l-wen').text
                # 提取其他信息
                authority = tr.find('h2', class_='l-wen1').text if tr.find('h2', class_='l-wen1') else ''
                type_of_law = tr.find_all('h2', class_='l-wen1')[1].text if len(tr.find_all('h2', class_='l-wen1')) > 1 else ''
                status = tr.find_all('h2', class_='l-wen1')[2].text if len(tr.find_all('h2', class_='l-wen1')) > 2 else ''
                date = tr.find_all('h2', class_='l-wen1')[3].text if len(tr.find_all('h2', class_='l-wen1')) > 3 else ''
                # 将信息添加到列表中
                results.append([url, title, authority, type_of_law, status, date])
            # 判断是否要点击下一页
            if i < page_num - 1:
                # 找到下一页的符号
                button = driver.find_element(By.CLASS_NAME, "layui-laypage-next")
                # 模拟鼠标左键点击
                button.click()
        driver.quit()
        result_list_pro = [['https://flk.npc.gov.cn' + x[0][1:], x[1], x[2], x[3], x[4], x[5]] for x in results]
        df = pd.DataFrame(columns=['url', 'title', 'authority', 'type_of_law', 'status', 'date'], data=result_list_pro)
        # 保存为表格文件
        df.to_excel(save_file)
    except Exception as e:
        print(e)


def download_docx(xlsx_file_path, download_folder, user_download_folder):
    from shutil import move
    from os.path import join
    from tqdm import tqdm
    import pandas as pd
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from time import sleep
    try:
        df = pd.read_excel(xlsx_file_path, index_col=0)
        df_now = df[df['status'] == '有效 '].reset_index(drop=True)
        # docx_base_url = 'https://wb.flk.npc.gov.cn/flfg/WORD/replace.docx'
        # 启动chrome驱动 这步需要等待一段时间
        driver = webdriver.Chrome()
        for i in tqdm(range(int(df_now.shape[0]))):
            data = df_now.iloc[i]
            # 进入想要爬取的网页
            driver.get(data['url'])
            # 通过唯一id检索到想要的元素
            element = driver.find_element(By.ID, "codeMa")
            # 提取文档id
            docs_id = element.get_attribute('src').split('/')[-1][:-4]
            # # 生成下载地址
            # download_url = docx_base_url.replace('replace', docs_id)
            # # 使用wget下载并重命名
            # os.system(f'wget {download_url} -O {os.path.join(download_folder, data["title"] + ".docx")}')
            # 通过唯一class检索到下载按钮
            button = driver.find_element(By.CLASS_NAME, "xia-z")
            # 模拟鼠标左键点击
            button.click()
            # 停顿5s保证下载完
            sleep(5)
            # 从默认下载地址移动到指定地址
            move(join(user_download_folder, docs_id + '.docx'), join(download_folder, data['title'] + '.docx'))
        driver.quit()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    # task = 1
    task = 2
    url = 'https://flk.npc.gov.cn/list.html?sort=true&type=flfg&xlwj=02,03,04,05,06,07,08'
    user_download_folder = r'C:\Users\70473\Downloads'
    if task == 1:
        get_result_list(url, 45, r'D:\project\reptile\url.xlsx')
    elif task == 2:
        download_docx(r'D:\project\reptile\url.xlsx', r'D:\project\reptile\docs', user_download_folder)



