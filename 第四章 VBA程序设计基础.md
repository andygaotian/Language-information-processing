# 第四章 VBA程序设计基础

### VBA常用函数

- len()
- right()
- left()
- mid()
- split()



- Msgbox()
- Inputbox()

### VBA 基础知识

#### 变量



#### 循环



#### 判断





### Excel 对象引用语法

- 行的引用： `Row(i)` i为行号
- 列的引用：`Column(i)` i为列号
- 单元格的引用：
  - `Cells(row, col)` row和col分别代表行、列号数字
  - `Range("A5")` 分别指定列号字母和行号数字，如`Range("B" & i)