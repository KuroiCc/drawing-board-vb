VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   1530
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      Height          =   400
      Index           =   8
      Left            =   2790
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "选择"
      Top             =   555
      Width           =   400
   End
   Begin VB.OptionButton Option1 
      Height          =   400
      Index           =   7
      Left            =   2250
      Picture         =   "Form3.frx":2F85
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "多边形"
      Top             =   555
      Width           =   400
   End
   Begin VB.OptionButton Option1 
      Height          =   400
      Index           =   6
      Left            =   3330
      Picture         =   "Form3.frx":786A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "圆"
      Top             =   90
      Width           =   400
   End
   Begin VB.OptionButton Option1 
      Height          =   400
      Index           =   5
      Left            =   2790
      Picture         =   "Form3.frx":9F46
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "矩形"
      Top             =   90
      Width           =   400
   End
   Begin VB.OptionButton Option1 
      Height          =   400
      Index           =   3
      Left            =   915
      Picture         =   "Form3.frx":DC9E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "填充：单击画布上某个区域用前景色进行填充"
      Top             =   615
      Width           =   400
   End
   Begin VB.OptionButton Option1 
      Height          =   400
      Index           =   2
      Left            =   420
      Picture         =   "Form3.frx":10234
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "取色器：从图片中选取颜色并将其用于前景色"
      Top             =   615
      Width           =   400
   End
   Begin VB.CheckBox Check2 
      Caption         =   "填充"
      Height          =   330
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "是否填充，填充的颜色为背景色"
      Top             =   600
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "轮廓"
      Height          =   330
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "是否有轮廓，轮廓的颜色为前景色"
      Top             =   150
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Height          =   400
      Index           =   4
      Left            =   2250
      Picture         =   "Form3.frx":12660
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "直线"
      Top             =   90
      Width           =   400
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   1980
      ScaleHeight     =   960
      ScaleWidth      =   2025
      TabIndex        =   4
      Top             =   60
      Width           =   2025
   End
   Begin VB.OptionButton Option1 
      Height          =   400
      Index           =   1
      Left            =   915
      Picture         =   "Form3.frx":14A14
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "铅笔：用选定的线宽画一个任意形状的线条"
      Top             =   105
      Width           =   400
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form3.frx":17908
      Left            =   6015
      List            =   "Form3.frx":17921
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "为所选择的工具选择宽度"
      Top             =   510
      Width           =   1440
   End
   Begin VB.OptionButton Option1 
      Height          =   400
      Index           =   0
      Left            =   420
      Picture         =   "Form3.frx":1793B
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "橡皮：擦除图像一部分，并用背景色代替"
      Top             =   105
      Width           =   400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工具"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   675
      TabIndex        =   13
      Top             =   1125
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "形状"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3210
      TabIndex        =   12
      Top             =   1125
      Width           =   360
   End
   Begin VB.Line Line2 
      X1              =   5040
      X2              =   5040
      Y1              =   45
      Y2              =   1400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "粗细："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5250
      TabIndex        =   0
      Top             =   540
      Width           =   720
   End
   Begin VB.Line Line1 
      X1              =   1845
      X2              =   1845
      Y1              =   45
      Y2              =   1400
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check2_Click()
    If Check2.Value = 1 Then Form1.Picture1.FillStyle = 0 Else Form1.Picture1.FillStyle = 1
End Sub

Private Sub Combo1_Change()
    Form1.Picture1.DrawWidth = Val(Combo1.Text)
    If Combo1.Text > 10 Then Combo1.Text = 10
    
End Sub

Private Sub Combo1_Click()
    Form1.Picture1.DrawWidth = Val(Combo1.Text)
End Sub

Private Sub Label1_Click()
    Label1.BorderStyle = 1
End Sub

Private Sub Form_Activate()
    Dim i%
    Dim c!
    c = 255
    For i = 1 To Me.ScaleHeight
        Me.Line (0, i)-(Me.ScaleWidth, i), RGB(c, c, c)
        c = c - 35 / Me.ScaleHeight
    Next
    c = 240
    For i = 1 To Me.ScaleHeight
        Picture1.Line (0, i)-(Me.ScaleWidth, i), RGB(c, c, c)
        c = c + 15 / Me.ScaleHeight
    Next
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            SelectedTool = 0
        Case 1
            SelectedTool = 1
        Case 2
            SelectedTool = 2
        Case 3
            SelectedTool = 3
        Case 4
            SelectedTool = 4
        Case 5
            SelectedTool = 5
            Form1.Shape1.Shape = 0
        Case 6
            SelectedTool = 6
            Form1.Shape1.Shape = 2
        Case 7
            SelectedTool = 7
            Isfirst = True
        Case 8
            SelectedTool = 8
    End Select
End Sub
