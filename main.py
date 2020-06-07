import os
import docx # python-docx
import openpyxl

"""
    Author: coasxu
    Version: v0.1
    Update time: 2020/6/7 09:27
"""

def cmp(x):
    return x[0]


if __name__ == "__main__":
    # 获取当年目录下唯一的xlsx文件
    files = os.listdir('.')
    xlpath = None
    for file in files:
        print(file)
        if file.endswith(".xlsx") and "~" not in file:
            xlpath = file
            break

    print("已加载文件 '%s'" % xlpath)
    # 读取xlsx
    wb=openpyxl.load_workbook(xlpath)
    ws = wb.active

    peoplelist = []
    i = 0
    while True:
        row = 2+i
        wname = ws['C%d'% row].value
        nid = ws['A%d'% row].value
        nname = ws['B%d'% row].value
    #     print(int(nid), nname, wname)
        if wname is None:
            break
        peoplelist.append([int(nid), nname, wname])
        i+=1

    peoplelist = sorted(peoplelist, key=cmp)


    # 获取当前目录下唯一的xlsx文件
    files = os.listdir()
    path = None
    for file in files:
        if file.endswith(".docx") or file.endswith(".doc"):
            if "~" not in file:
                path = file
                break

    print("已加载文件 '%s'" % path)
    file=docx.Document(path)

    infos = {}
    usernow = None
    for para in file.paragraphs:
        if usernow is None and para.text != "":
            usernow = para.text.split(":")[0]
            if usernow not in infos:
                infos[usernow] = []
        elif usernow is not None:
            if para.text == "":
                usernow = None
            else:
                infos[usernow].append(para.text)

    # 匹配姓名
    successlist = []
    faillist = []
    for i,v in infos.items():
        peop = None
        for index in range(len(peoplelist)):
            people = peoplelist[index]
            if people[2] == i:
                peop = people.copy()
                peop.insert(0, index)
            
        if peop is None:
            faillist.append([i, v])
        else:
            peop.extend([v])
            successlist.append(peop)

    kickcardlist = []

    # 输出打卡记录
    for f in successlist:
        shougong = 0
        kaigong = 0
        for ff in f[4]:
            if "开工" in ff:
                kaigong = 1
            if "收工" in ff:
                shougong = 1
        kickcardlist.append([f[0], f[1],f[2], kaigong, shougong])

    print("******************成功匹配名单******************")
    print("共%d人" % len(kickcardlist))

    print("学号 姓名 开工 收工")
    for line in kickcardlist:
        print(line[1],end=" ")
        print(line[2],end=" ")
        print(line[3],end=" ")
        print(line[4],end=" ")
        print()
    print("******************无法匹配名单******************")
    print("复制后的名称 聊天记录")
    for line in faillist:
        print(line[0],end=" ")
        print(line[1],end=" ")
        print()
    print()
    print("开始修改文件 '%s'" % xlpath)
    # 读取xlsx
    wb=openpyxl.load_workbook(xlpath)
    ws=wb.worksheets[0]
    for card in kickcardlist:
        row = card[0] + 2
        s_id = card[1]
        if card[3] == 1:
            ws["D%d" % row] = 1
        if card[4] == 1:
            ws["E%d" % row] = 1
    wb.save(filename=xlpath)

    print("修改文件 '%s' 完成，已保存！" % xlpath)
    print("各位辛苦啦！")


    line = input("输入任意字符退出")