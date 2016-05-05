VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Game of Life"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   13560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      BackColor       =   &H0080FFFF&
      Caption         =   "完成选点"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080FFFF&
      Caption         =   "界面选点"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton command6 
      BackColor       =   &H0080FF80&
      Caption         =   "周期设定"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   10680
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "随机选点"
      Height          =   495
      Left            =   8520
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "下一步"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "暂停"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "执行"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "初始化"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Game of Life"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1095
      Left            =   8760
      TabIndex        =   13
      Top             =   6480
      Width           =   4695
   End
   Begin VB.Label Label7 
      Caption         =   "2"
      Height          =   255
      Left            =   9360
      TabIndex        =   12
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "当前周期：  秒"
      Height          =   255
      Left            =   8520
      TabIndex        =   11
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   $"Form1.frx":1CCA
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   4455
      Left            =   9960
      TabIndex        =   10
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "单步执行："
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   8520
      TabIndex        =   6
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "连续执行："
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   8520
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   8760
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "存活单元格数量："
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   8280
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'ACM软件软件程序设计大赛决赛作品

Dim x1 As Long       '绘制网格所用坐标
Dim y1 As Long
Dim draw_x As Long   '用于过程drawing中的点的绘制（活细胞）
Dim draw_y As Long   '同上
Dim i As Integer     '二维数组纵向
Dim j As Integer     '二维数组横向
Dim temp_square(80, 80) As Boolean  '定义一个辅助性数组存储临时数据，最后用于给赋值
Dim square(80, 80) As Boolean  '布尔型二维数组
Dim flag As Integer
Dim cell_num As Long
Dim life_time As Variant
Dim xforarray As Single
Dim yforarray As Single


Private Sub Command1_Click() '初始化按钮

redrawing   '用于重绘图形
revariable   '用于初始化变量
initialize_array
cell_count
Label2.Caption = cell_num
Timer1.Enabled = False
Command5.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command8.Enabled = flase
Command7.Enabled = True

flag = 0
End Sub


Private Sub Command2_Click() '生成细胞按钮
redrawing
random_array
drawing
cell_count
Label2.Caption = cell_num
flag = 1
End Sub

Private Sub Command3_Click()   '执行 按钮
If flag = 1 Then
Command5.Enabled = False
Timer1.Enabled = True
End If

End Sub

Private Sub Command4_Click()     '暂停 按钮
If flag = 1 Then
Timer1.Enabled = False
Command5.Enabled = True
End If
End Sub

Private Sub Command5_Click()   '下一步 按钮
If flag = 1 Then
redrawing
calculate_array
drawing
cell_count
Label2.Caption = cell_num
End If
End Sub

Private Sub command6_Click()
life_time = InputBox("请输入细胞变化周期" + Chr(13) + "        单位：秒", "周期设定", 2)
Label7.Caption = life_time
End Sub



Private Sub Command7_Click()    '界面选点按钮



cell_num = 0
Command1_Click  '初始化按钮
flag = 2 '定义标记变量
Command7.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command8.Enabled = True





End Sub

Private Sub Command8_Click()

If flag = 2 Then
flag = 1
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command8.Enabled = flase
Command7.Enabled = True
End If


End Sub

Private Sub Form_Load()
Command8.Enabled = False

Timer1.Enabled = False
flag = 0
life_time = 2
redrawing            '重绘图像
initialize_array     '数组的初始化，应加入图形绘制（不要重绘）
  
End Sub


'过程 代码区
'------------------------------------------
Sub calculate_array()     '定义一个调用数组计算下一周期细胞数的函数
For i = 1 To 80
For j = 1 To 80
Call judge(i, j) '将数组下表传递过去，用于判断
Next j
Next i

'判断完毕，赋值给原来数组

For i = 1 To 80
For j = 1 To 80
square(i, j) = temp_square(i, j)   '将计算后的结果复制给square数组，实现细胞状态的变化
Next j
Next i

End Sub
Public Sub judge(i As Integer, j As Integer)    '核心区域，用于判断每个单元格的生死
Dim count As Integer    '用于存储细胞周围存活数量
Dim m As Integer
Dim n As Integer
count = 0   '初始化计数变量

If i = 1 And j > 1 And j < 80 Then       '四个边缘的点
For m = i To i + 1            '循环判断
For n = j - 1 To j + 1
If m <> i Or n <> j Then
If square(m, n) = True Then
count = count + 1
End If
End If
Next n
Next m
End If


If i = 80 And j > 1 And j < 80 Then
For m = i - 1 To i           '循环判断
For n = j - 1 To j + 1
If m <> i Or n <> j Then
If square(m, n) = True Then
count = count + 1
End If
End If
Next n
Next m
End If



If j = 1 And i > 1 And i < 80 Then
For m = i - 1 To i + 1            '循环判断
For n = j To j + 1
If m <> i Or n <> j Then
If square(m, n) = True Then
count = count + 1
End If
End If
Next n
Next m
End If


If j = 80 And i > 1 And i < 80 Then
For m = i - 1 To i + 1            '循环判断
For n = j - 1 To j
If m <> i Or n <> j Then
If square(m, n) = True Then
count = count + 1
End If
End If
Next n
Next m
End If



If i = 1 And j = 1 Then   '四个角落的点
For m = i To i + 1            '循环判断
For n = j To j + 1
If m <> i Or n <> j Then
If square(m, n) = True Then
count = count + 1
End If
End If
Next n
Next m
End If

If i = 1 And j = 80 Then
For m = i To i + 1            '循环判断
For n = j - 1 To j
If m <> i Or n <> j Then
If square(m, n) = True Then
count = count + 1
End If
End If
Next n
Next m
End If

If i = 80 And j = 1 Then
For m = i - 1 To i            '循环判断
For n = j To j + 1
If m <> i Or n <> j Then
If square(m, n) = True Then
count = count + 1
End If
End If
Next n
Next m
End If

If i = 80 And j = 80 Then
For m = i - 1 To i             '循环判断
For n = j - 1 To j
If m <> i Or n <> j Then
If square(m, n) = True Then
count = count + 1
End If
End If
Next n
Next m
End If

If i > 1 And j > 1 And i < 80 And j < 80 Then   '防止越界
For m = i - 1 To i + 1            '循环判断
For n = j - 1 To j + 1

If m <> i Or n <> j Then

If square(m, n) = True Then
count = count + 1
End If

End If

Next n
Next m

End If


Select Case count          '多分支选择判断
Case Is < 2
temp_square(i, j) = False
Case Is = 2
If square(i, j) = True Then
temp_square(i, j) = True
Else
temp_square(i, j) = False
End If
Case Is = 3
temp_square(i, j) = True
Case Is > 3
temp_square(i, j) = False
End Select

End Sub


Sub redrawing()
'在这里写重绘代码，用于网格的重新绘制（初始化）
Form1.Cls
DrawWidth = 1
CurrentX = 0
CurrentY = 0
x1 = CurrentX
y1 = CurrentY + 100
For i = 1 To 80
x1 = X
For j = 1 To 80
x1 = x1 + 100
Line (x1, y1)-(x1 - 100, y1 - 100), vbBlue, B
Next j
Print Chr(13)
y1 = y1 + 100
Next i

End Sub

Sub drawing() '在这里写绘制图像的代码，可能要加参。。。通过数组的值绘制!!!最重要！！！


For i = 1 To 80
For j = 1 To 80

If square(i, j) = True Then
'这里为算法核心区域

DrawWidth = 5
draw_x = (j - 1) * 100 + 50
draw_y = (i - 1) * 100 + 50
PSet (draw_x, draw_y), vbRed

End If

Next j
Next i

End Sub

Sub revariable()   '变量的初始化
CurrentX = 0
CurrentY = 0
x1 = CurrentX
y1 = CurrentY
'在这里添加所有变量的复位赋值


End Sub

Sub initialize_array()
'定义一个过程对数组进行初始化
For i = 1 To 80
For j = 1 To 80
square(i, j) = False
Next j
Next i
End Sub

Sub random_array()   '随机为数组赋值
For i = 1 To 80
For j = 1 To 80
Randomize
square(i, j) = Int(Rnd * 2) '产生随机数0-1并赋值，产生大约网格数量一半的“活细胞”
Next j
Next i
End Sub

Sub cell_count()
cell_num = 0
 For i = 1 To 80
 For j = 1 To 80
 If square(i, j) = True Then
 cell_num = cell_num + 1
 End If
 Next j
 Next i
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'这里为鼠标单击事件，按下鼠标选定方格
DrawWidth = 5
If flag = 2 Then
If Button = 1 Then
If X > 0 And X < 80 * 100 And Y > 0 And Y < 80 * 100 Then '限定单机事件的有效范围

j = Int(X / 100) + 1
i = Int(Y / 100) + 1
square(i, j) = True
PSet (((j - 1) * 100 + 50), ((i - 1) * 100 + 50)), vbRed
cell_num = cell_num + 1
Label2.Caption = cell_num

End If
End If
End If

End Sub

Private Sub Timer1_Timer()
Timer1.Interval = Val(life_time) * 1000
Call Command5_Click
Call cell_count
End Sub

