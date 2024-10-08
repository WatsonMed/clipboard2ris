
import re
import time
from tkinter import Tk, messagebox
from bs4 import BeautifulSoup
import win32clipboard
import os
import requests
from lxml import html


def get_clipboard_data(): # 获取剪贴板数据
    win32clipboard.OpenClipboard()
    data = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    return data

def is_doi(text):
    if re.match(r'10\.\d{4,9}/[-._;()/:A-Z0-9a-z]+', text) is not None:
        return True
    else:
        return False

# text = '10.1016/j.lungcan.2020.11.005' # 10.1177/17588359221146134  # 10.1016/j.lungcan.2020.11.005
# is_doi(text)

#
# def get_doi_by_title(title): # 测试成功
#     # PubMed搜索URL
#     search_url = "https://pubmed.ncbi.nlm.nih.gov/?term="
#
#     # 对标题进行URL编码
#     title_encoded = requests.utils.quote(title)
#
#     # 拼接完整的搜索URL
#     full_url = search_url + title_encoded
#
#     # 发起GET请求
#     response = requests.get(full_url)
#
#     # 检查请求是否成功
#     if response.status_code == 200:
#         # 解析HTML内容
#         soup = BeautifulSoup(response.text, 'html.parser')
#
#         # 查找所有包含DOI的<a>标签
#         doi_tags = soup.find_all('a', href=True)
#
#         # 遍历所有<a>标签，找到DOI
#         for tag in doi_tags:
#             href = tag['href']
#             if 'doi.org' in href:
#                 # 提取DOI号
#                 print("DOI:", href)
#                 # 提取 10\.\d{4,9}/[-._;()/:A-Z0-9a-z]+
#                 doi = re.search(r'10\.\d{4,9}/[-._;()/:A-Z0-9a-z]+', href).group()
#                 return doi
#         return "DOI not found"
#     else:
#         return "Error: Unable to fetch data from PubMed"
#


def get_doi_by_title(title):
    # 将title的\n替换为空格，两个以上的空格替换为单个空格，字符串前后都去掉空格
    title = re.sub(r'\s+|\n', ' ', title).strip()

    # PubMed搜索URL
    search_url = "https://pubmed.ncbi.nlm.nih.gov/?term="

    # 对标题进行URL编码
    title_encoded = requests.utils.quote(title)

    # 拼接完整的搜索URL
    full_url = search_url + title_encoded

    # 发起GET请求
    response = requests.get(full_url)

    # 检查请求是否成功
    if response.status_code == 200:
        # 使用lxml解析HTML内容
        tree = html.fromstring(response.content)

        # 判断页面是否有特定的XPath元素 //div[contains(@class, 'top-citations')]
        # specific_element = tree.xpath('//*[@id="search-results"]/section[1]/div[2]/article/div/a')
        specific_element = tree.xpath('//div[contains(@class, "top-citations")]')

        if specific_element:
            html_source = html.tostring(specific_element[0], encoding='unicode')
            # 提取href值并构建新的URL
            href_value = re.search('href="/(\d+)/"', html_source).group(1)
            new_url = f"https://pubmed.ncbi.nlm.nih.gov/{href_value}"
            response = requests.get(new_url)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                # 查找所有包含DOI的<a>标签
                doi_tags = soup.find_all('a', href=True)
                # 遍历所有<a>标签，找到DOI
                for tag in doi_tags:
                    href = tag['href']
                    if 'doi.org' in href:
                        # 提取DOI号
                        print("DOI:", href)
                        # 提取 10\.\d{4,9}/[-._;()/:A-Z0-9a-z]+
                        doi = re.search(r'10\.\d{4,9}/[-._;()/:A-Z0-9a-z]+', href).group()
                        return doi
                else:
                    return "DOI not found"
            else:
                return "Error: Unable to fetch data from the new URL"
        else:
            # 如果没有特定的XPath元素，则直接从页面获取DOI
            soup = BeautifulSoup(response.text, 'html.parser')
            doi_tags = soup.find_all('a', href=True)
            for tag in doi_tags:
                href = tag['href']
                if 'doi.org' in href:
                    doi = re.search(r'10\.\d{4,9}/[-._;()/:A-Z0-9a-z]+', href).group()
                    return doi
            return "DOI not found"
    else:
        return "Error: Unable to fetch data from PubMed"


# 测试函数
# title = "Your Paper Title Here"
# doi = get_doi_by_title(title)
# print(f"The DOI of the paper is: {doi}")

def get_ris_from_doi(doi):
    # Crossref API endpoint for DOI to RIS conversion
    url = f"https://api.crossref.org/works/{doi}/transform/application/x-research-info-systems"

    # Send a GET request to the API
    response = requests.get(url)

    # Check if the request was successful
    if response.status_code == 200:
        # 创建一个文件，用于保存RIS数据
        doi_name = doi.replace('/', ' ')
        ris_content = response.text
        ris_title = re.search(r'TI  -(.+)\n', ris_content).group(1)
        ris_title = ris_title.replace('/', ' ').replace(':', ' ')
        ris_file_name = f"{doi_name}_{ris_title}"
        # 创建ris_file_name RIS文件
        # Write the RIS data to a file
        with open(f"{ris_file_name}.RIS", "w") as file:
            file.write(response.text)
        print(f"RIS file for DOI {doi} has been saved.")
        return True
    else:
        print(f"Failed to retrieve RIS data for DOI {doi}. Status code: {response.status_code}")
        return False



def show_notification(message, title="Notification"):
    # 创建一个隐藏的主窗口。
    # 显示一个消息框，标题为title，内容为message。
    # 销毁主窗口。
    # 等待1秒，确保消息框已关闭。这里实际上是为了给用户时间查看消息，因为showinfo是阻塞调用，窗口关闭后才继续执行。
    root = Tk()
    root.withdraw()  # 非阻塞模式
    messagebox.showinfo(title, message)
    root.destroy()
    time.sleep(1)  # 等待弹窗关闭

# def show_notification(message, title="Notification"): # 显示提示框，等待用户关闭
#     root = Tk()
#     # 移除下面这行，不隐藏主窗口
#     # root.withdraw()  # 非阻塞模式
#     messagebox.showinfo(title, message)
#     root.mainloop()  # 启动事件循环，等待用户关闭窗口

def main():
    previous_clipboard = ''
    while True:
        try:
            current_clipboard = get_clipboard_data()
            current_clipboard = title = re.sub(r'\s+|\n', ' ', current_clipboard).strip()
            if current_clipboard != previous_clipboard:
                if is_doi(current_clipboard):
                    if get_ris_from_doi(current_clipboard):
                        show_notification(f"RIS file for DOI {current_clipboard} has been saved.")
                    else:
                        show_notification("Failed to retrieve RIS data for DOI {}".format(current_clipboard))
                else:
                    doi = get_doi_by_title(current_clipboard)
                    if doi != "DOI not found" and get_ris_from_doi(doi):
                        show_notification(f"RIS file for title '{current_clipboard}' has been saved.")
                    else:
                        show_notification("DOI not found or failed to retrieve RIS data for title {}".format(current_clipboard))

                break  # 完成操作后退出循环
            previous_clipboard = current_clipboard
        except Exception as e:
            show_notification(f"An error occurred: {e}")
            break
        time.sleep(1)


if __name__ == "__main__":
    main()


