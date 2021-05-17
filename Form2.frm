VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   960
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Height          =   240
      Left            =   585
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "将前景色和背景色交换"
      Top             =   105
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   1080
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   28
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00808080&
      Height          =   240
      Index           =   1
      Left            =   1350
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   27
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00008080&
      Height          =   240
      Index           =   2
      Left            =   1890
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   26
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00000080&
      Height          =   240
      Index           =   3
      Left            =   1620
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   25
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00008000&
      Height          =   240
      Index           =   4
      Left            =   2160
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   24
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00808000&
      Height          =   240
      Index           =   5
      Left            =   2430
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   23
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00800080&
      Height          =   240
      Index           =   6
      Left            =   2970
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   22
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00800000&
      Height          =   240
      Index           =   7
      Left            =   2700
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   21
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H0000C0C0&
      Height          =   240
      Index           =   8
      Left            =   3240
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   20
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H0000C000&
      Height          =   240
      Index           =   9
      Left            =   3510
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   19
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00C00000&
      Height          =   240
      Index           =   10
      Left            =   4050
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   18
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00C0C000&
      Height          =   240
      Index           =   11
      Left            =   3780
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   17
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00C000C0&
      Height          =   240
      Index           =   12
      Left            =   4320
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   16
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00004080&
      Height          =   240
      Index           =   13
      Left            =   4590
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   15
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   1080
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   14
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00C0C0C0&
      Height          =   240
      Index           =   15
      Left            =   1350
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   13
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H0000FFFF&
      Height          =   240
      Index           =   16
      Left            =   1890
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   12
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H000000FF&
      Height          =   240
      Index           =   17
      Left            =   1620
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   11
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H0000FF00&
      Height          =   240
      Index           =   18
      Left            =   2160
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   10
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00FFFF00&
      Height          =   240
      Index           =   19
      Left            =   2430
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   9
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00FF00FF&
      Height          =   240
      Index           =   20
      Left            =   2970
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   8
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00FF0000&
      Height          =   240
      Index           =   21
      Left            =   2700
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   7
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H0080FFFF&
      Height          =   240
      Index           =   22
      Left            =   3240
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   6
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H0080FF80&
      Height          =   240
      Index           =   23
      Left            =   3510
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   5
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00FF8080&
      Height          =   240
      Index           =   24
      Left            =   4050
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   4
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00FFFF80&
      Height          =   240
      Index           =   25
      Left            =   3780
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   3
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00FF80FF&
      Height          =   240
      Index           =   26
      Left            =   4320
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   2
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H0080C0FF&
      Height          =   240
      Index           =   27
      Left            =   4590
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   1
      Top             =   480
      Width           =   240
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   735
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   300
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   29
         Top             =   300
         Width           =   300
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   30
         Top             =   450
         Width           =   300
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim t
    t = DbFcolor: DbFcolor = DbBcolor: DbBcolor = t
     Picture1.BackColor = DbFcolor
     Picture2.BackColor = DbBcolor
End Sub

Private Sub Picture0_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Picture1.BackColor = Picture0(Index).BackColor
    If Button = 2 Then Picture2.BackColor = Picture0(Index).BackColor
    DbFcolor = Picture1.BackColor
    DbBcolor = Picture2.BackColor
End Sub
