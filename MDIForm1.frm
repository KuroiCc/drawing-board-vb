VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "��ͼv0.9��"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   7680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '��Ļ����
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   3300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu m1 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu new 
         Caption         =   "�½�(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu save 
         Caption         =   "����(&S)"
         Shortcut        =   ^O
      End
      Begin VB.Menu fgx1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "�˳�(&X)      Alt+F4"
      End
   End
   Begin VB.Menu m2 
      Caption         =   "�༭(&E)"
      Begin VB.Menu cut 
         Caption         =   "����(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "����(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu plst 
         Caption         =   "ճ��(&P)"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu m3 
      Caption         =   "ͼ��(&I)"
      Begin VB.Menu Flip_Horizontal 
         Caption         =   "ˮƽ��ת(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu Flip_Vertical 
         Caption         =   "��ֱ��ת(&V)"
         Shortcut        =   ^F
      End
      Begin VB.Menu clear 
         Caption         =   "���ͼ��(&C)"
      End
   End
   Begin VB.Menu m4 
      Caption         =   "��ɫ(&C)"
      Begin VB.Menu coloredit 
         Caption         =   "�༭��ɫ(&E)..."
      End
   End
   Begin VB.Menu m5 
      Caption         =   "����(&H)"
      Begin VB.Menu about 
         Caption         =   "���ڱ����..."
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As String, xx, yy, sWidth, sHeight

Private Sub about_Click()
    MsgBox "�����е�����ϢְҵѧУ" & vbCrLf & "13���1��  �³�" & vbCrLf & "2015-3-5", vbInformation, "��ͼv0.9�� ����"
End Sub

Private Sub clear_Click()
    Call new_Click
End Sub

Private Sub coloredit_Click()
    CommonDialog1.ShowColor
    DbFcolor = CommonDialog1.Color
    Form2.Picture1.BackColor = DbFcolor
End Sub

Private Sub copy_Click()
    If SelectAr = True Then
        xx = Smaller(Selectedx1, Selectedx2)
        yy = Smaller(Selectedy1, Selectedy2)
        sWidth = Abs(Selectedx1 - Selectedx2)
        sHeight = Abs(Selectedy1 - Selectedy2)
        Form4.Picture1.Width = sWidth * 15
        Form4.Picture1.Height = sHeight * 15
        Form4.Picture1.PaintPicture Form1.Picture1.Image, 0, 0, , , xx, yy, sWidth, sHeight, &HCC0020
    Else
        MsgBox "����ѡ������", vbOKOnly + vbExclamation, "����"
    End If
End Sub

Private Sub cut_Click()
    If SelectAr = True Then
        xx = Smaller(Selectedx1, Selectedx2)
        yy = Smaller(Selectedy1, Selectedy2)
        sWidth = Abs(Selectedx1 - Selectedx2)
        sHeight = Abs(Selectedy1 - Selectedy2)
        Form4.Picture1.Width = sWidth * 15
        Form4.Picture1.Height = sHeight * 15
        Form4.Picture1.PaintPicture Form1.Picture1.Image, 0, 0, , , xx, yy, sWidth, sHeight, &HCC0020
        Form1.Picture1.Line (xx, yy)-(sWidth + xx, sHeight + yy), DbBcolor, BF
    Else
        MsgBox "����ѡ������", vbOKOnly + vbExclamation, "����"
    End If
End Sub

Private Sub exit_Click()
    a = MsgBox("�Ƿ񱣴��ļ�?", vbYesNo, "��ȷ��")
    If a = 6 Then
        CommonDialog1.Filter = "bmp�ļ�|*.bmp|�����ļ�|*.*"
        CommonDialog1.Action = 2
        f$ = CommonDialog1.FileName
        If f$ <> "" Then
            SavePicture Form1.Picture1.Image, f$
        End If
    End If
    dl = DeleteObject(newbrush)
    dl = DeleteObject(oldbrush)
    End
End Sub

Private Sub Flip_Horizontal_Click()
    SavePicture Form1.Picture1.Image, App.Path & "\backup.bmp"
    Form1.Picture2.Picture = LoadPicture(App.Path & "\backup.bmp")
    Form1.Picture1.PaintPicture Form1.Picture2.Picture, Form1.Picture1.ScaleWidth, 0, -1 * Form1.Picture1.ScaleWidth, Form1.Picture1.ScaleHeight
End Sub

Private Sub Flip_Vertical_Click()
    SavePicture Form1.Picture1.Image, App.Path & "\backup.bmp"
    Form1.Picture2.Picture = LoadPicture(App.Path & "\backup.bmp")
    Form1.Picture1.PaintPicture Form1.Picture2.Picture, 0, Form1.Picture1.ScaleHeight, Form1.Picture1.ScaleWidth, -1 * Form1.Picture1.ScaleHeight
End Sub

Private Sub MDIForm_Load()
    Form3.Show
    Form3.Top = 0
    Form3.Left = 0
    Form1.Show
    Form1.Left = 0
    Form1.Top = Form3.Height
    Form2.Show
    Form2.Left = 0
    Form2.Top = Form3.Height + Form1.Height
    DbFcolor = Form2.Picture1.BackColor
    DbBcolor = Form2.Picture2.BackColor
    SelectedTool = 1
    Form4.Hide
    hbitmap& = CreateBitmap(8, 8, 1, 1, ARRY(1))
    newbrush& = CreatePatternBrush(hbitmap)
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    a = MsgBox("�Ƿ񱣴��ļ�?", vbYesNo, "��ȷ��")
    If a = 6 Then
        CommonDialog1.Filter = "bmp�ļ�|*.bmp|�����ļ�|*.*"
        CommonDialog1.Action = 2
        f$ = CommonDialog1.FileName
        If f$ <> "" Then
            SavePicture Form1.Picture1.Image, f$
        End If
    End If
    dl = DeleteObject(newbrush)
    dl = DeleteObject(oldbrush)
End Sub

Private Sub new_Click()
    a = MsgBox("�Ƿ񱣴��ļ�?", vbYesNo, "��ȷ��")
    If a = 6 Then
        CommonDialog1.Filter = "bmp�ļ�|*.bmp|�����ļ�|*.*"
        CommonDialog1.Action = 2
        f$ = CommonDialog1.FileName
        If f$ <> "" Then
            SavePicture Form1.Picture1.Image, f$
        End If
    End If
    Form1.Picture1.Line (0, 0)-(Form1.Picture1.Width, Form1.Picture1.Height), RGB(255, 255, 255), BF '����
End Sub

Private Sub plst_Click()
    If SelectAr = True Then
        xx = Smaller(Selectedx1, Selectedx2)
        yy = Smaller(Selectedy1, Selectedy2)
        sWidth = Abs(Selectedx1 - Selectedx2)
        sHeight = Abs(Selectedy1 - Selectedy2)
        Form1.Picture1.PaintPicture Form4.Picture1.Image, xx, yy, sWidth, sHeight, 0, 0, Form4.Picture1.Width / 15, Form4.Picture1.Height / 15, &HCC0020
    Else
        MsgBox "����ѡ������", vbOKOnly + vbExclamation, "����"
    End If
End Sub

Private Sub save_Click()
    CommonDialog1.Filter = "bmp�ļ�|*.bmp|�����ļ�|*.*"
    CommonDialog1.Action = 2
    f$ = CommonDialog1.FileName
    If f$ <> "" Then
        SavePicture Form1.Picture1.Image, f$
    End If
End Sub

