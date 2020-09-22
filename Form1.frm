VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hover Buttons"
   ClientHeight    =   2895
   ClientLeft      =   1560
   ClientTop       =   900
   ClientWidth     =   8100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Default         =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ListBox List3 
      Height          =   1815
      ItemData        =   "Form1.frx":0000
      Left            =   6240
      List            =   "Form1.frx":000D
      TabIndex        =   8
      Top             =   480
      Width           =   1695
   End
   Begin VB.PictureBox Btn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   1815
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
      Begin VB.Label txt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.PictureBox Btn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   1815
      TabIndex        =   4
      Top             =   960
      Width           =   1815
      Begin VB.Label txt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.ListBox List2 
      Height          =   1815
      ItemData        =   "Form1.frx":0035
      Left            =   4440
      List            =   "Form1.frx":0042
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "Form1.frx":0057
      Left            =   2640
      List            =   "Form1.frx":0064
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.PictureBox Btn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.Label txt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Message on Click"
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Selected Text"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Normal Text"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nr

Private Sub Btn_Click(Index As Integer)
MsgBox List3.List(Index)
End Sub

Private Sub Btn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Selected Index
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Dim A
For A = 0 To Btn.Count - 1
MakeBtn Btn(A), A, List1.List(A), "Verdana", 9, 140
Next A
End Sub
Sub MakeBtn(PicButton As PictureBox, Index, LbText, LbFont, LbFontSize, LabelY)
PicButton.BackColor = &HC0C0C0
txt(Index).ForeColor = vbBlack
txt(Index).Top = LabelY
txt(Index).Font = LbFont
txt(Index).FontSize = LbFontSize
txt(Index).FontBold = True
txt(Index).Caption = LbText
PicButton.Line (PicButton.Width, PicButton.Height - 20)-(0, PicButton.Height - 20)
PicButton.Line (PicButton.Width, 0)-(0, 0), vbWhite
PicButton.Line (0, PicButton.Height)-(0, -20), vbWhite
PicButton.Line (PicButton.Width - 20, 0)-(PicButton.Width - 20, PicButton.Height - 20)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Restore
End Sub
Sub Restore()
Dim A
If Nr <> "" Then
A = Mid(Nr, 3, Len(Nr) - 2)
MakeBtn Btn(A), A, List1.List(A), "Verdana", 9, 140
Nr = ""
Else
End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Restore
End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Restore
End Sub

Private Sub List3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Restore
End Sub

Private Sub txt_Click(Index As Integer)
MsgBox List3.List(Index)
End Sub
Sub Selected(Index As Integer)
Restore
Nr = "NR" & Index
txt(Index).Caption = List2.List(Index)
txt(Index).ForeColor = vbRed
Btn(Index).Line (Btn(Index).Width, Btn(Index).Height - 20)-(0, Btn(Index).Height - 20), vbWhite
Btn(Index).Line (Btn(Index).Width, 0)-(0, 0)
Btn(Index).Line (0, Btn(Index).Height)-(0, -20)
Btn(Index).Line (Btn(Index).Width - 20, 0)-(Btn(Index).Width - 20, Btn(Index).Height - 20), vbWhite
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Selected Index
End Sub
