import re
from datetime import datetime

from pandas import DataFrame


def txt_to_dict() -> dict[str]:
    path = input("请将txt文件拖到此处并回车：")
    print('正在分析txt文件……')
    try:
        with open(path, "r", encoding="utf-8") as f:
            data_list = f.readlines()
    except FileNotFoundError as e:
        print(f"此文件不存在：\n{e}")
        input()
        exit()

    data_dict = {}
    for i in range(len(data_list)):
        data = data_list[i]
        split_list = data.split()
        if len(split_list) < 5:
            print(f"第{i+1}行内容错误：{data}")
            continue
        name_and_uid = split_list[2]
        gift = split_list[4]
        pattern = re.compile(r"(.+)\((\d+?)\)$")
        match = re.search(pattern, name_and_uid)
        if match:
            name = match.group(1)
            uid = match.group(2)
        else:
            print(f"匹配失败：{data}")
            continue
        if uid not in data_dict:
            data_dict[uid] = {
                "name": name,
                gift: 1,
            }
        elif gift not in data_dict[uid]:
            data_dict[uid][gift] = 1
        else:
            data_dict[uid][gift] += 1

    return data_dict


def dict_to_xlsx(data_dict: dict):
    print('正在写入xlsx文件……')
    dataframe = DataFrame(data_dict).T
    while True:
        time_now = datetime.now().strftime("%Y-%m-%dT%H_%M_%S")
        file_name = f"data_{time_now}.xlsx"
        try:
            dataframe.to_excel(file_name, "data", na_rep=0)
            break
        except PermissionError as e:
            print(f"写入文件失败: 权限不足\n可能是存在同名文件‘{file_name}’或此文件夹无写入权限")
            input(f"{e}\n回车退出程序")
            exit()
    return file_name


if __name__ == "__main__":
    file_name = dict_to_xlsx(txt_to_dict())
    input(f"完成！请查看与本程序同文件夹的‘{file_name}’。\n回车退出程序")
