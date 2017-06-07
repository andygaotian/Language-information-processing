# 第四章 VBA程序设计基础

> 相关参考资料见 `\参考资料` 文件夹
>
> 另推荐一个Excel VBA学习网站：[Excel Home](http://club.excelhome.net/)
>
> 推荐一本书：[《ExcelVBA实战技巧精粹》](https://pan.baidu.com/share/link?shareid=1336292023&uk=2835369685)

## 1. Excel 对象引用语法

### 1.1 行的引用： 

 	`Rows(i)` i为行号

### 1.2 列的引用：

	`Columns(i)` i为列号


### 1.3单元格的引用：

  - `Cells(row, col)` row和col分别代表行、列号数字

  - `Range("A5")` 分别指定列号字母和行号数字，如`Range("B" & i)

### 1.4 工作簿的引用

`Worksheets("sheet1")`
例：`Worksheets("sheet2").cells(2,1)` 指`sheet2`工作表的`A2`单元格
### 1.5 格式
#### 1.5.1 字号
```vbscript
Cells(1, 1).Font.Size = 12
```
#### 1.5.2 粗体、斜体
```vbscript
Cells(1, 1).Font.Bold = True    '设为粗体`
Cells(1, 1).Font.Italic = True   '设为粗体`
```
#### 1.5.3 颜色

##### 1.5.3.1 文本颜色

```vbscript
Cells(1, 1).Font.Color =vbRed
Cells(1, 1).Font.ColorIndex = 3
```
##### 1.5.3.2 背景色

```vbscript
Cells(1, 1).Interior.ColorIndex = 36
```

##### 1.5.3.3 vba颜色代码（ColorIndex）

![vba颜色代码](/参考资料/vba颜色代码.png)

### 1.6 边框

#### 1.6.1 设定边框的对象

可针对行（如`Rows(3)`）、列（如`Columns(3)`）、单元格（如`Range("A5")`、`Cells(2,1)`）和Range区域（如`Range("A1:B3")`）指定表格边框。

#### 1.6.2 指定边框线

- 所有边框 `Borders`，如`Cells(2,2).Borders`将设定单元格	`B2`的四个边框线。
- 上框线：`Borders(xlEdgeTop)`，如`Cells(2,2).Borders(xlEdgeTop`设定上框线。
- 下框线：`Borders(xlEdgeBottom)`。
- 上框线：`Borders(xlEdgeLeft)`。
- 上框线：`Borders(xlEdgeRight)`。
- 内部横线：`xlInsideHorizontal`。
- 内部竖线：`xlInsideVertical`。
#### 1.6.3 指定线型 `.LineStyle`
| 名称              | 值     | 说明       |
| --------------- | ----- | -------- |
| xlContinuous    | 1     | 实线。      |
| xlDash          | -4115 | 虚线。      |
| xlDashDot       | 4     | 点划相间线。   |
| xlDashDotDot    | 5     | 划线后跟两个点。 |
| xlDot           | -4118 | 点式线。     |
| xlDouble        | -4119 | 双线。      |
| xlLineStyleNone | -4142 | 无线条。     |
| xlSlantDashDot  | 13    | 倾斜的划线。   |

#### 1.6.4 线条粗细`.Weight`

| 名称         | 说明   |
| ---------- | ---- |
| xlHairline | 最细   |
| xlThin     | 细    |
| xlMedium   | 中等   |
| xlThick    | 粗    |

#### 1.6.5 示例

```vbscript
With Cells(1,1).Borders(xlEdgeLeft)
  .LineStyle = xlContinuous
  .ColorIndex = 0
  .Weight = xlThin
End With
```

## 2. VBA 基础知识

### 2.1 数据类型

1. **`Integer`：** 整型，存储整数，范围为-32768——32767
2. **Long：** 长整型，存储整数，范围：-2,147,483,648——2,147,483,647
3. **Single：** 单精度浮点型，存储小数，最大精度约为小数点后6位
4. **Double：** 双精度浮点型，存储小数，最大精度约为小数点后14位
5. **`String`：** 字符串型，存储文本字符串，需用半角双引号包裹，如 `"abcd"`
6. **Boolean：** 布尔型，存储真假两种逻辑值，其值只有两种：`True`和`False`

### 2.2 变量

使用VBA变量前一般先声明（指定）其数据类型，亦可不声明（隐式声明，由系统根据变量的值自动判断其数据类型）。变量声明语法：

`Dim var As Integer`

`Dim str As String`

### 2.2 运算符

#### 2.2.1 数学运算符
| 加    | 减    | 乘    | 除    | 整除   | 取模   | 乘方   |
| ---- | ---- | ---- | ---- | ---- | ---- | ---- |
| +    | -    | *    | /    | \    | Mod  | ^    |


​    整除：`7 \ 3 = 2`

​    取模：除法运算的余数，例：`7 Mod 3 = 1`

​    乘方：`2 ^ 4 = 16`

#### 2.2.2 连接运算符`&`和`+`

- `&`运算符用于将两个字符串连接在一起。

  例：`"abc" & “def" = "abcdef`

  当运算对象不是字符串型时，系统会先将非字符串数据转换为字符串，然后再做字符串连接运算，如：`12 & "34" = "1234"`、 `12 & 34 = "1234"`

- `+`运算符亦可用于字符串连接。

  当两个运算对象都是数字类型时，执行数学运算，当其中之一是字符串型时，现将数字类型转换为字符串，然后执行字符串连接运算。由于`+`运算符具有两种运算功能，当运算对象的数据类型不明确时，容易产生非预期错误，因此，不建议使用`+`运算符做字符串连接。

#### 2.2.3 比较运算符

- 大于：`>`

- 小于：`<`

- 等于：`=`

- 大于等于：`>=`

- 小于等于：`<=`

- 不等于：`<>`

比较运算符运算结果为`Boolean`型，即为`True`或`False`，因此比较运算符常在判断语句中使用。例：`3 > 4 = False`

#### 2.2.4 逻辑运算符

- **与：**`And` 执行逻辑与运算，返回一个逻辑值（`True` 或`False`），当两个运算表达式都为真时，返回`True`，否则返回`False`。例： 

  - `2>1 And True = True`
  - `3>1 And 2=3 = False`
  - `False And False = False`

- **或：**`Or` 执行逻辑或运算，当两个运算表达式中有一个为真时，即返回`True`。

- **非：** `Not`，执行逻辑非运算，当表达式为`True`时，返回`False`，反之。例：

  - `Not 2>1 = False`
  - `Not Cells(1,1).Font.Bold`如果单元格A1字体为粗体，则返回`False`，反之。

#### 2.2.5 赋值运算符`=`

将值赋给变量，如：

```visual basic
Dim a As Integer
a = 3   '将数字3赋给变量a
a = a + 1 '将a的值加1后再赋给a，因此此句执行后a = 4
```

#### 2.2.6 运算符的优先级

- 算术运算符>连接运算符>比较运算符>逻辑运算符

  如：表达式 `1+2 & "3" = "33" Or "ab" <> "c"` 的结果为 True

- 同算数运算中一样，当需要指定复杂表达式的运算顺序时，可使用`()`进行人工定义。

  如：`(1 + 2) * 3 = 8 And True  `

### 2.3 循环语句

#### 2.3.1 `For`循环

基本语法：

```vbscript
For i = 1 To 30
	Debug.Print i
Next i
'以上代码i取值为1，2，3……30
```

进阶语法：

```vbscript
For i = 1 to 100 Step 3
	Debug.Print i
Next i
'Step用于指定步进值，因此i的取值为1，4，7，11……100
```

```vbscript
For i = 100 to 1 Step -1
	Debug.Print i
Next i
'步进值可以为负数，因此i的取值为100，99，98……1
```

#### 2.3.2 `Do` 循环

For循环适用于知道确切的循环次数的情况，但当循环次数不确定（如条件不同则循环次数不同）时，For循环就不适用了。Do循环可以根据给定的条件灵活确定循环的次数。

##### 2.3.2.1 格式1

先判断条件是否为真，真则执行内部语句，否则跳出循环。

```vbscript
Do {while |until} condition
	Statements
	Exit do
	Statements
Loop
```

例1：

```VBScript
'sum = 1+2+3+…………n，求sum小于100最大值
sum = 0
i = 1
Do While sum < 100
	sum = sum + i
	i = i + 1
Loop
Debug.Print i-2
```

例2：

```vbscript
'sum = 1+2+3+…………n，求sum小于100最大值
sum = 0
i = 1
Do Until sum > 100
	sum = sum + i
	i = i + 1
Loop
Debug.Print i-2
```

##### 2.3.2.2 格式2

格式二同格式一的差别在于，先执行操作，再判断条件是否满足。因此，格式一可能一次也不执行循环内部的操作，格式二则至少会执行一次。

```vbscript
Do                          
	Statements
	Exit do
	Statements
Loop {while |until} condition
```

### 2.4 判断语句

#### 2.4.1 `If`语句

语法：

```vbscript
If condition Then
	do something
Elseif condition then
    do something
Elseif condition Then
    do something
Else
    do something
End if
```

`Elseif`语句可根据需要重复多次，也可不出现，`Else`最多出现一次，也可不出现。

例子：

```vbscript
'最简单的判断语句
If a > 3 Then
    b = a + 1
End if
```

```vbscript
'二值判断
if a >3 then
	a = a + 1
Else
	a = a + 2
End if
```

```vbscript
'多值判断
s = cells(i,2)
If s >= 96 And s <= 100 Then
        Range("X" & i).Value = "A+"
 ElseIf s >= 90 And s <= 95 Then
         Range("X" & i).Value = "A"
 ElseIf s >= 86 And s <= 89 Then
        Range("X" & i).Value = "B+"
 ElseIf s >= 80 And s <= 85 Then
         Range("X" & i).Value = "B"
 ElseIf s >= 76 And s <= 79 Then
        Range("X" & i).Value = "C+"
 ElseIf s >= 70 And s <= 75 Then
         Range("X" & i).Value = "C"
 ElseIf s >= 66 And s <= 69 Then
        Range("X" & i).Value = "D+"
 ElseIf s >= 60 And s <= 65 Then
         Range("X" & i).Value = "D"
 ElseIf s < 60 Then
         Range("X" & i).Value = "F"
End If
```

#### 2.4.2 `Case`语句

`If`语句已能满足所有条件判断要求，但当判断条件较多时，语句较为繁琐，因此有语法相对简洁的`Case`语句可选：

例：

```vbscript
s = Range("W" & i).Value
Select Case s
     Case 96 To 100
            Range("X" & i).Value = "A+"
     Case 90 To 95
             Range("X" & i).Value = "A"
     Case 86 To 89
            Range("X" & i).Value = "B+"
     Case 80 To 85
             Range("X" & i).Value = "B"
     Case 76 To 79
            Range("X" & i).Value = "C+"
     Case 70 To 75
             Range("X" & i).Value = "C"
     Case 66 To 69
            Range("X" & i).Value = "D+"
     Case 60 To 65
             Range("X" & i).Value = "D"
     Case Else
             Range("X" & i).Value = "F"
End Select
```

### 2.5 练习

1. 在A列中，从第一行至第10行，填充以下序列：100，95，90……

2. 在第一行中，从第一列到第6列，填充以下序列：2，4，8，16……

3. 在`数据样本\第一章\六甲成绩单-上学期.xlsm`的`排名`表中，将操行评定为A和A+的记录设为红色，操行评定为C的设为绿色。

4. 使用vba在excel中生成以下乘法表：

   ![乘法表](E:\CODE\Github\Language-information-processing\数据样本\第四章\乘法表.png)







## 3. VBA字符串处理

- len()
- right()
- left()
- mid()
- split()




- - ​