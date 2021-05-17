VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4493.991
   ScaleMode       =   0  'User
   ScaleWidth      =   7345.275
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   9210
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   1
      Top             =   6015
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   15
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Shape SelectA 
         BorderStyle     =   3  'Dot
         Height          =   495
         Left            =   285
         Top             =   1905
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   3  'Dot
         Height          =   495
         Left            =   255
         Top             =   1185
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   3
         X2              =   83
         Y1              =   2
         Y2              =   34
      End
      Begin VB.Shape Eraser 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   105
         Left            =   300
         Top             =   345
         Visible         =   0   'False
         Width           =   105
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FirstX, FirstY, LastX, LastY

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    x1 = X
    y1 = Y
    SelectA.Visible = False
    SelectAr = False
    If SelectedTool = 1 Then '铅笔画点
        Picture1.Line (X, Y)-(X, Y), DbFcolor
        Picture1.PSet (X, Y), DbFcolor
    End If
    If SelectedTool = 3 Then '填充颜色
        Filling DbFcolor, X, Y
    End If
    If SelectedTool = 7 And Button = 1 And Isfirst = False Then '多边形预览
        Line1.BorderWidth = Val(Form3.Combo1.Text)
        Line1.BorderColor = DbFcolor
        Line1.x1 = LastX: Line1.y1 = LastY
        Line1.X2 = X: Line1.Y2 = Y
        Line1.Visible = True
    End If
    If SelectedTool = 8 Then '选区
        Selectedx1 = X: Selectedy1 = Y
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SelectedTool = 0 Then '橡皮跟随
        Eraser.Width = Val(Form3.Combo1.Text) * 5 + 5
        Eraser.Height = Val(Form3.Combo1.Text) * 5 + 5
        Eraser.Left = X - Eraser.Width / 2
        Eraser.Top = Y - Eraser.Height / 2
        Eraser.Visible = True
    Else
        Eraser.Visible = False
    End If
    If Button = 1 Then
        If SelectedTool = 0 Then Picture1.Line (Eraser.Left, Eraser.Top)-(Eraser.Left + Eraser.Width, Eraser.Top + Eraser.Height), DbBcolor, BF '橡皮擦除
        If SelectedTool = 1 Then Picture1.Line -(X, Y), DbFcolor '画笔画线
        If SelectedTool = 2 Then '取色
            DbFcolor = Picture1.Point(X, Y)
            Form2.Picture1.BackColor = DbFcolor
        End If
        If SelectedTool = 4 Then '直线预览
            Line1.BorderWidth = Val(Form3.Combo1.Text)
            Line1.BorderColor = DbFcolor
            Line1.x1 = x1: Line1.y1 = y1
            Line1.X2 = X: Line1.Y2 = Y
            Line1.Visible = True
        End If
        If SelectedTool = 5 Or SelectedTool = 6 Then '矩形和圆预览
            Shape1.Left = Smaller(x1, X)
            Shape1.Top = Smaller(y1, Y)
            Shape1.Width = Abs(X - x1)
            Shape1.Height = Abs(Y - y1)
            Shape1.Visible = True
        End If
        If SelectedTool = 7 Then '多边形预览
            Line1.BorderWidth = Val(Form3.Combo1.Text)
            Line1.BorderColor = DbFcolor
            If Isfirst = True Then
                Line1.x1 = x1: Line1.y1 = y1
                Line1.X2 = X: Line1.Y2 = Y
            Else
                Line1.x1 = LastX: Line1.y1 = LastY
                Line1.X2 = X: Line1.Y2 = Y
            End If
            Line1.Visible = True
        End If
        If SelectedTool = 8 Then '选区预览
            SelectA.Left = Smaller(x1, X)
            SelectA.Top = Smaller(y1, Y)
            SelectA.Width = Abs(X - x1)
            SelectA.Height = Abs(Y - y1)
            SelectA.Visible = True
            SelectAr = True
        End If
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Line1.Visible = False
    Shape1.Visible = False
    If Button = 1 Then
        If SelectedTool = 4 Then '画直线
            Picture1.Line (x1, y1)-(X, Y), DbFcolor
        End If
        If SelectedTool = 5 Then '画矩形
            Picture1.FillColor = DbBcolor
            If Form3.Check1.Value = 1 Then
                Picture1.Line (x1, y1)-(X, Y), DbFcolor, B
            Else
                If Form3.Check2.Value = 1 Then Picture1.Line (x1, y1)-(X, Y), DbFcolor, BF
            End If
        End If
        If SelectedTool = 6 Then '画圆
            Picture1.FillColor = DbBcolor
            If Form3.Check1.Value = 1 Then
                Picture1.Circle ((x1 + X) / 2, (y1 + Y) / 2), Bigger(Shape1.Width, Shape1.Height) / 2, DbFcolor, , , Shape1.Height / Shape1.Width
            Else
                Picture1.Circle ((x1 + X) / 2, (y1 + Y) / 2), Bigger(Shape1.Width, Shape1.Height) / 2, DbBcolor, , , Shape1.Height / Shape1.Width
            End If
        End If
        If SelectedTool = 7 Then '画多边形
            If Isfirst = True Then
                Picture1.Line (x1, y1)-(X, Y), DbFcolor
                Isfirst = False
                FirstX = x1: FirstY = y1
                LastX = X: LastY = Y
            Else
                If Sqr((FirstX - X) ^ 2 + (FirstY - Y) ^ 2) <= 10 Then
                    Picture1.Line (LastX, LastY)-(FirstX, FirstY), DbFcolor
                    Isfirst = True
                Else
                    Picture1.Line (LastX, LastY)-(X, Y), DbFcolor
                    LastX = X: LastY = Y
                End If
            End If
        End If
        If SelectedTool = 8 Then '选区
            Selectedx2 = X: Selectedy2 = Y
        End If
    End If
End Sub
