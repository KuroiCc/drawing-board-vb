Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function SelectObject& Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long)
Public Declare Function DeleteObject& Lib "GDI32" (ByVal hObject As Long)
Public Declare Function CreateBitmap& Lib "GDI32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any)
Public Declare Function CreatePatternBrush& Lib "GDI32" (ByVal hbitmap As Long)
Public Declare Function ExtFloodFill& Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long)
Private Const FLOODFILLBORDER& = 0
Private Const FLOODFILLSURFACE& = 1

Public x1
Public y1
Public DbFcolor
Public DbBcolor
Public SelectedTool%
Public Selectedx1
Public Selectedx2
Public Selectedy1
Public Selectedy2
Public SelectAr As Boolean
Public Isfirst As Boolean

Public ARRY(1 To 16) As Integer
Public hbitmap&
Public oldbrush&
Public newbrush&
Public thiscolor&
Public dl

Public Sub Reset_Darea()
    Form1.Picture1.BackColor = DbBcolor
End Sub

Public Sub Filling(a_color, X, Y)
    oldbrush& = SelectObject(Form1.Picture1.hDC, newbrush)
    Form1.Picture1.ForeColor = a_color
    thiscolor = Form1.Picture1.Point(X, Y)
    dl = ExtFloodFill(Form1.Picture1.hDC, X, Y, thiscolor, FLOODFILLSURFACE)
    dl = SelectObject(Form1.Picture1.hDC, oldbrush)
End Sub

Public Function Smaller(ByVal a, ByVal b)
    If a <= b Then Smaller = a
    If b < a Then Smaller = b
End Function

Public Function Bigger(ByVal a, ByVal b)
    If a >= b Then Bigger = a
    If b > a Then Bigger = b
End Function
