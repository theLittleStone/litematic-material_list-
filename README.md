# litematic模组材料清单的转换脚本
因为litematic mod 导出的材料列表为txt文件, 很难看, 所以把它转换成excel文件

顺便把材料数量转换成了xx组, xx盒这样的模式

# 使用方法:
把需要转换的txt文件拖到run.py上, 或者直接双击运行run.py

前者只转换拖动的文件, 后者默认转换当前文件夹下全部符合要求的txt文件

转换完的xls文件与原文件位于同一目录下

## 注意: 
运行前需要安装python并下载xlwt库, 并设置python.exe为.py文件的默认启动方式 

**需要安装xlwt包**

下载方法:
```
pip install xlwt
```


# 效果:

## 原文件: 

![原文件](https://github.com/theLittleStone/litematic-material_list-/blob/main/pictures/before.png)

## 转换后:

![转换后](https://github.com/theLittleStone/litematic-material_list-/blob/main/pictures/after.png)
