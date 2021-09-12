

# python with excel



## 1. excel常见库文件能力对比

![image-20210912210746044](https://tu-chuang-1253216127.cos.ap-beijing.myqcloud.com/20210912210746.png)



个人推荐学习的库文件:

1. openpyxl
2. xlwings



## 2. 安装库文件

```python
pip install xlwings
```





## 3. xlwing学习

`xlwings`开源免费，能够非常方便的读写Excel文件中的数据，并且能够进行单元格格式的修改。

`xlwings`还可以和matplotlib、numpy以及pandas无缝连接，支持读写numpy、pandas数据类型，将matplotlib可视化图表导入到excel中。

最重要的是`xlwings`可以调用Excel文件中VBA写好的程序，也可以让VBA调用用Python写的程序。

![img](https://imgconvert.csdnimg.cn/aHR0cHM6Ly91cGxvYWQtaW1hZ2VzLmppYW5zaHUuaW8vdXBsb2FkX2ltYWdlcy8yOTc5MTk2LTRhMmFiMGJhZjllMjZkNjcucG5n?x-oss-process=image/format,png)

app：应用 一个xlwing程序

book: excel程序 工作簿

sheets： 工作表

range：范围

### 3.1 模块导入

```text
import xlwings as xw
```

### 3.2 写入excel的基本操作

```python
import xlwings as xw
# 创建一个应用  visible : 操作的时候是否可视化，  add_book：打开excel是否新建excel表
# visible and add_book default value all is true.
app = xw.App(visible=True, add_book=False)
# 工作簿
workbook = app.books.add()
# 工作表
sheet = workbook.sheets['sheet1']
# 写入数据
sheet.range('A1').value = "天天学习"

# 保存工作表
workbook.save(r'./excel_sheet/cds_test2.xlsx')
# 关闭工作表
workbook.close()
# 退出app
app.quit()

```

![image-20210912232610859](https://tu-chuang-1253216127.cos.ap-beijing.myqcloud.com/20210912232610.png)

### 3.3 写入excel常见操作

```python
import xlwings as xw
# 创建一个应用  visible : 操作的时候是否可视化，  add_book：打开excel是否新建excel表
# visible and add_book default value all is true.
app = xw.App(visible=True, add_book=False)
# 工作簿
workbook = app.books.add()
# 工作表
sheet = workbook.sheets['sheet1']
# 指定单元格写入数据
sheet.range('A1').value = "天天学习"

# 指定单元格写入
sheet.range("A2").value = "new value"

# 指定单元格写入
sheet.range("A2").value = "new value"

# 插入一列 transpose：翻转
sheet.range("C5:C8").options(transpose=True).value = [5, 6, 7]

# 保存
workbook.save(r'./excel_sheet/demo3.xlsx')
workbook.close()
app.quit()
```

![image-20210912232826397](https://tu-chuang-1253216127.cos.ap-beijing.myqcloud.com/20210912232826.png)



### 3.4  常规读取

```python
import xlwings as xw
# 创建一个应用  visible : 操作的时候是否可视化，  add_book：打开excel是否新建excel表
# visible and add_book default value all is true.
app = xw.App(visible=True, add_book=False)
# 工作簿
workbook = app.books.open(r'./excel_sheet/demo3.xlsx')
# 工作表
sheet = workbook.sheets['sheet1']

print(sheet.range("a2").value)


workbook.save()
# 关闭工作表
workbook.close()
# 退出app
app.quit()

```



![image-20210912232724432](C:\Users\cds\AppData\Roaming\Typora\typora-user-images\image-20210912232724432.png)

### 3.5 读取excel常用方式

```python
import xlwings as xw
# 创建一个应用  visible : 操作的时候是否可视化，  add_book：打开excel是否新建excel表
# visible and add_book default value all is true.
app = xw.App(visible=True, add_book=False)
# 工作簿
workbook = app.books.open(r'./excel_sheet/demo3.xlsx')
# 工作表
sheet = workbook.sheets['sheet1']

# 读取某个位置的数值
print(sheet.range("a2").value)

# 读取某行
print(sheet.range('c4:g4').value)

# 读一列
print(sheet.range('c4:c7').value)

# 读行列
print(sheet.range('c4:g5').value)


# 保存
workbook.save()
# 关闭工作表
workbook.close()
# 退出app
app.quit()
```

### 3.6 更多读取操作

```python
import xlwings as xw
# 创建一个应用  visible : 操作的时候是否可视化，  add_book：打开excel是否新建excel表
# visible and add_book default value all is true.
app = xw.App(visible=True, add_book=False)
# 工作簿
workbook = app.books.open(r'./excel_sheet/demo3.xlsx')
# 工作表
sheet = workbook.sheets['sheet1']


```



![image-20210912233118264](https://tu-chuang-1253216127.cos.ap-beijing.myqcloud.com/20210912233118.png)

```python
# 读取一列
print(sheet.range('a1').expand('down').value)

# 读取一列
print(sheet.range('b1').expand('down').value)

# 读取一列
print(sheet.range('c1').expand('down').value)

# 读取所有关联的
print(sheet.range('c4').expand('table').value)

# 读取所有关联的
print(sheet.range('A1').options(expand='table').value)
```



![image-20210912233110414](https://tu-chuang-1253216127.cos.ap-beijing.myqcloud.com/20210912233110.png)