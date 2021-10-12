import xlwt
import re
import sys
import os
import time

# 支持文件拖入启动
# 如果直接点开, 则转换全部符合的文件
file_paths = sys.argv[1:]
if not file_paths:
    dirs = os.walk("./")
    for files in dirs:
        file = files[2]
        for f in file:
            if re.search(".*[.]txt$", f):
                file_paths.append(f)

for path in file_paths:
    f = None
    excelFile = None
    excelSheet = None
    try:
        f = open(path)
    except:
        print(path + '读取文件失败')
        continue
    try:
        excelFile = xlwt.Workbook()
        excelSheet = excelFile.add_sheet('1')
    except:
        print(path + "写文件失败")
        continue
    excelSheet.col(0).width = 22 * 256
    excelSheet.col(1).width = 10 * 256
    excelSheet.col(2).width = 14 * 256
    excelSheet.col(3).width = 17 * 256

    # 设置单元格格式
    font1 = xlwt.Font()
    font1.bold = True
    font2 = font1
    font1.height = 20 * 14  # 11是字号

    # 居中, 加粗, 放大
    style1 = xlwt.XFStyle()
    a1 = xlwt.Alignment()
    a1.horz = a1.HORZ_CENTER
    style1.alignment = a1
    style1.font = font1

    # 加粗, 不居中
    style2 = xlwt.XFStyle()
    style2.font = font2

    try:
        contentLines = f.readlines()
    except:
        print(1)
    isTitleLine = True
    excelLine = 1
    styleIsRight = False
    for line in contentLines:
        if line.startswith('+--'):
            styleIsRight = True
            continue
        elif isTitleLine and styleIsRight:
            title = line
            title = re.findall('[|]\s*(.*?)\s*[|]', title)
            excelSheet.write_merge(0, 0, 0, 3, title, style1)
            isTitleLine = False
        elif styleIsRight:
            strList = re.findall('[|]\s*(.*?)\s*[|]\s*(.*?)\s*[|]', line)
            st = strList[0]
            if st[1].isdigit():
                materialNum = int(st[1])
                excelSheet.write(excelLine, 0, st[0])
                excelSheet.write(excelLine, 1, st[1])
                if materialNum >= 64:  # 算组数
                    if materialNum % 64 == 0:
                        excelSheet.write(excelLine, 2, (str(materialNum // 64) + " set(s)"))
                    else:
                        excelSheet.write(excelLine, 2, (str(materialNum // 64) + " set(s) + " + str(materialNum % 64)))
                    if materialNum > (64 * 27):
                        excelSheet.write(excelLine, 3, str(materialNum // (64 * 27), ) + ' box(s) +' + str((materialNum % (64 * 27)) // 64 + 1) + " set(s)")
                    else:
                        excelSheet.write(excelLine, 3, "-----------")
                else:
                    excelSheet.write(excelLine, 2, "< 1")
                    excelSheet.write(excelLine, 3, "-----------")
                # 算盒数
            else:
                materialNum = 0
                excelSheet.write(excelLine, 2, "SetNum", style2)
                excelSheet.write(excelLine, 3, "BoxNum", style2)
                excelSheet.write(excelLine, 1, st[1], style2)
                excelSheet.write(excelLine, 0, st[0], style2)
            excelLine += 1

    if not styleIsRight:
        print(path + '不是对应格式的文件')
        continue
    else:
        try:
            excelFile.save(path + ".xls")
        except:
            print(path + "保存文件失败, 文件可能在别处被打开")
    f.close()
    print(path + '转换完成')
time.sleep(1.5)
