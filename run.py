import xlwt
import re
import sys
import os
import time
import chardet

# 支持文件拖入启动
# 如果直接点开, 则转换全部符合的文件
file_paths = sys.argv[1:]
if not file_paths:
    files = os.listdir(os.getcwd())
    for f in files:
        if re.search(".*[.]txt$", f):
            file_paths.append(f)

for path in file_paths:
    f = None
    encod = ''
    excelFile = None
    excelSheet = None
    try:
        with open(path, mode='rb') as m:
            inp = m.read()
            encod = chardet.detect(inp)['encoding']
        f = open(path, mode='r', encoding=encod)

    except:
        print(path + '读取文件失败')
        continue
    try:
        excelFile = xlwt.Workbook()
        excelSheet = excelFile.add_sheet('material')
    except:
        print(path + "写文件失败")
        f.close()
        continue
    excelSheet.col(0).width = 22 * 256
    excelSheet.col(1).width = 10 * 256
    excelSheet.col(2).width = 14 * 256
    excelSheet.col(3).width = 17 * 256

    # 设置单元格格式
    font0 = xlwt.Font()
    font0.bold = True
    font3 = font0
    font0.height = 20 * 14  # 11是字号

    # 白色背景
    pattern1 = xlwt.Pattern()
    pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern1.pattern_fore_colour = 1
    # 灰色背景
    pattern2 = xlwt.Pattern()
    pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern2.pattern_fore_colour = 67

    # 居中, 加粗, 放大
    style0 = xlwt.XFStyle()
    a0 = xlwt.Alignment()
    a0.horz = a0.HORZ_CENTER
    style0.alignment = a0
    style0.font = font0
    style0.pattern = pattern2

    style1 = xlwt.XFStyle()
    style2 = xlwt.XFStyle()
    style1.pattern = pattern1
    style2.pattern = pattern2
    # 加粗, 不居中
    style3 = xlwt.XFStyle()
    style3.font = font3

    try:
        contentLines = f.readlines()
    except:
        print("读文件失败")
        f.close()
        continue
    isTitleLine = True
    excelLine = 1
    styleIsRight = False
    for line in contentLines:
        # 每行背景色
        if excelLine % 2 == 0:
            style = style2
        else:
            style = style1

        if line.startswith('+--'):
            styleIsRight = True
            continue
        elif isTitleLine and styleIsRight:
            title = line
            title = re.findall('[|]\s*(.*?)\s*[|]', title)
            excelSheet.write_merge(0, 0, 0, 3, title, style0)
            isTitleLine = False
        elif styleIsRight:
            strList = re.findall('[|]\s*(.*?)\s*[|]\s*(.*?)\s*[|]', line)
            st = strList[0]
            if st[1].isdigit():
                materialNum = int(st[1])
                excelSheet.write(excelLine, 0, st[0], style)
                excelSheet.write(excelLine, 1, st[1], style)
                if materialNum >= 64:  # 算组数
                    if materialNum % 64 == 0:
                        excelSheet.write(excelLine, 2, (str(materialNum // 64) + " set(s)"), style)
                    else:
                        excelSheet.write(excelLine, 2, (str(materialNum // 64) + " set(s) + " + str(materialNum % 64)), style)
                    if materialNum > (64 * 27):
                        excelSheet.write(excelLine, 3, str(materialNum // (64 * 27), ) + ' box(s) +' + str((materialNum % (64 * 27)) // 64 + 1) + " set(s)", style)
                    else:
                        excelSheet.write(excelLine, 3, "----------", style)
                else:
                    excelSheet.write(excelLine, 2, "----------", style)
                    excelSheet.write(excelLine, 3, "----------", style)
                # 算盒数
            else:
                materialNum = 0
                excelSheet.write(excelLine, 2, "SetNum", style3)
                excelSheet.write(excelLine, 3, "BoxNum", style3)
                excelSheet.write(excelLine, 1, st[1], style3)
                excelSheet.write(excelLine, 0, st[0], style3)
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
            continue
        print(path + '转换完成')
    f.close()

time.sleep(1.5)
