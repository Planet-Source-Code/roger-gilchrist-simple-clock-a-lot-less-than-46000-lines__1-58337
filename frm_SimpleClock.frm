VERSION 5.00
Begin VB.Form frm_SimpleClock 
   BackColor       =   &H8000000C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alarm Clock v1.1"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Interval        =   1000
      Left            =   3840
      Top             =   0
   End
   Begin VB.Image imgNum4 
      Height          =   600
      Index           =   6
      Left            =   3720
      Picture         =   "frm_SimpleClock.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum4 
      Height          =   600
      Index           =   5
      Left            =   3720
      Picture         =   "frm_SimpleClock.frx":0F42
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum4 
      Height          =   105
      Index           =   4
      Left            =   3120
      Picture         =   "frm_SimpleClock.frx":1E84
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgNum4 
      Height          =   105
      Index           =   3
      Left            =   3120
      Picture         =   "frm_SimpleClock.frx":2CD6
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgNum4 
      Height          =   105
      Index           =   2
      Left            =   3120
      Picture         =   "frm_SimpleClock.frx":3B28
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgNum4 
      Height          =   600
      Index           =   1
      Left            =   3000
      Picture         =   "frm_SimpleClock.frx":497A
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum4 
      Height          =   600
      Index           =   0
      Left            =   3000
      Picture         =   "frm_SimpleClock.frx":58BC
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum3 
      Height          =   600
      Index           =   0
      Left            =   2040
      Picture         =   "frm_SimpleClock.frx":67FE
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum3 
      Height          =   600
      Index           =   1
      Left            =   2040
      Picture         =   "frm_SimpleClock.frx":7740
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum3 
      Height          =   105
      Index           =   2
      Left            =   2115
      Picture         =   "frm_SimpleClock.frx":8682
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgNum3 
      Height          =   105
      Index           =   3
      Left            =   2160
      Picture         =   "frm_SimpleClock.frx":94D4
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgNum3 
      Height          =   105
      Index           =   4
      Left            =   2160
      Picture         =   "frm_SimpleClock.frx":A326
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgNum3 
      Height          =   600
      Index           =   5
      Left            =   2760
      Picture         =   "frm_SimpleClock.frx":B178
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum3 
      Height          =   600
      Index           =   6
      Left            =   2760
      Picture         =   "frm_SimpleClock.frx":C0BA
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgDot 
      Height          =   180
      Index           =   1
      Left            =   1800
      Picture         =   "frm_SimpleClock.frx":CFFC
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgDot 
      Height          =   180
      Index           =   0
      Left            =   1800
      Picture         =   "frm_SimpleClock.frx":D1EE
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgNum2 
      Height          =   600
      Index           =   0
      Left            =   1080
      Picture         =   "frm_SimpleClock.frx":D3E0
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum2 
      Height          =   600
      Index           =   1
      Left            =   1080
      Picture         =   "frm_SimpleClock.frx":E322
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum2 
      Height          =   105
      Index           =   2
      Left            =   1125
      Picture         =   "frm_SimpleClock.frx":F264
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgNum2 
      Height          =   105
      Index           =   3
      Left            =   1200
      Picture         =   "frm_SimpleClock.frx":100B6
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgNum2 
      Height          =   105
      Index           =   4
      Left            =   1200
      Picture         =   "frm_SimpleClock.frx":10F08
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgNum2 
      Height          =   600
      Index           =   5
      Left            =   1680
      Picture         =   "frm_SimpleClock.frx":11D5A
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum2 
      Height          =   600
      Index           =   6
      Left            =   1680
      Picture         =   "frm_SimpleClock.frx":12C9C
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblAM 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "AM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgNum1 
      Height          =   600
      Index           =   6
      Left            =   960
      Picture         =   "frm_SimpleClock.frx":13BDE
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum1 
      Height          =   600
      Index           =   5
      Left            =   960
      Picture         =   "frm_SimpleClock.frx":14B20
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum1 
      Height          =   105
      Index           =   4
      Left            =   360
      Picture         =   "frm_SimpleClock.frx":15A62
      Stretch         =   -1  'True
      Top             =   1410
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgNum1 
      Height          =   105
      Index           =   3
      Left            =   360
      Picture         =   "frm_SimpleClock.frx":168B4
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgNum1 
      Height          =   105
      Index           =   2
      Left            =   360
      Picture         =   "frm_SimpleClock.frx":17706
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgNum1 
      Height          =   600
      Index           =   1
      Left            =   240
      Picture         =   "frm_SimpleClock.frx":18558
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgNum1 
      Height          =   600
      Index           =   0
      Left            =   240
      Picture         =   "frm_SimpleClock.frx":1949A
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape shpFace 
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frm_SimpleClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  'just to make sure they are out of step with each other

  imgDot(0).Visible = Not imgDot(1).Visible

End Sub

Private Function InArray(ByVal Test As Long, _
                         ParamArray arr() As Variant) As Boolean

  Dim P As Variant

  For Each P In arr
    If Test = P Then
      InArray = True
    End If
  Next P

End Function

Private Sub LED(ArrayImg As Variant, _
                LngVal As Long)

  ' use this to work out what turns on
  '  2
  '0  5
  ' 3
  '1  6
  ' 4

  ArrayImg(0).Visible = InArray(LngVal, 4, 5, 6, 8, 9, 0)
  ArrayImg(1).Visible = InArray(LngVal, 2, 6, 8, 0)
  ArrayImg(2).Visible = InArray(LngVal, 2, 3, 5, 6, 7, 8, 9, 0)
  ArrayImg(3).Visible = InArray(LngVal, 2, 3, 4, 5, 6, 8, 9)
  ArrayImg(4).Visible = InArray(LngVal, 2, 3, 5, 6, 8, 9, 0)
  ArrayImg(5).Visible = InArray(LngVal, 1, 2, 3, 4, 7, 8, 9, 0)
  ArrayImg(6).Visible = InArray(LngVal, 1, 3, 4, 5, 6, 7, 8, 9, 0)

End Sub

Private Sub tmrUpdate_Timer()

  Dim strTime As String

  strTime = Format$(Time, "hh:mm AM/PM")
  LED imgNum1, Val(Left$(strTime, 1))
  LED imgNum2, Val(Mid$(strTime, 2, 1))
  LED imgNum3, Val(Mid$(strTime, 4, 1))
  LED imgNum4, Val(Mid$(strTime, 5, 1))
  imgDot(0).Visible = Not imgDot(0).Visible
  imgDot(1).Visible = Not imgDot(1).Visible
  If lblAM.Caption <> Right$(strTime, 2) Then
    lblAM.Caption = Right$(strTime, 2)
  End If

End Sub

':)Code Fixer V2.8.9 (18/01/2005 2:07:50 PM) 6 + 59 = 65 Lines Thanks Ulli for inspiration and lots of code.

