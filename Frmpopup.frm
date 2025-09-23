VERSION 5.00
Begin VB.Form Frmpopup 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1680
      Top             =   1560
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Picture         =   "Frmpopup.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "TrÝch khÊu h&ao"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1220
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "§¸nh gi¸ l¹&i"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   810
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Gi¶&m TSC§"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   400
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080FF80&
      Caption         =   "T¨ng T&SC§"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Frmpopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FrmChungtu.Show2
End Sub

Private Sub Command2_Click()
FrmChungtu.Show3
End Sub

Private Sub Command3_Click()
FrmChungtu.Show4
End Sub

Private Sub Command4_Click()
Hide
End Sub
Public Sub closeMain()
Timer1.Enabled = True
End Sub
Private Sub Command7_Click()
FrmChungtu.Show1
End Sub
Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Ð?i màu n?n khi di chu?t
    Command7.BackColor = 8438015
    Command1.BackColor = &H80FF80
    Command2.BackColor = &H80FF80
    Command3.BackColor = &H80FF80
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Ð?i màu n?n khi di chu?t
    Command1.BackColor = 8438015
    Command7.BackColor = &H80FF80
    Command2.BackColor = &H80FF80
    Command3.BackColor = &H80FF80
     Command4.BackColor = &H80FF80
End Sub
Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Ð?i màu n?n khi di chu?t
    Command2.BackColor = 8438015
    Command1.BackColor = &H80FF80
    Command7.BackColor = &H80FF80
    Command3.BackColor = &H80FF80
     Command4.BackColor = &H80FF80
End Sub
Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Ð?i màu n?n khi di chu?t
    Command3.BackColor = 8438015
    Command1.BackColor = &H80FF80
    Command7.BackColor = &H80FF80
    Command2.BackColor = &H80FF80
    Command4.BackColor = &H80FF80
End Sub
Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Ð?i màu n?n khi di chu?t
    Command4.BackColor = 8438015
    Command1.BackColor = &H80FF80
    Command7.BackColor = &H80FF80
    Command2.BackColor = &H80FF80
    Command3.BackColor = &H80FF80
End Sub
Private Sub Form_Load()
   ' Gi? s? b?n mu?n d?t form ? v? trí 10% t? trái và 10% t? trên
' Gi? s? b?n mu?n d?t form con ? 10% t? trái và 10% t? trên c?a FrmChungtu
Dim percentLeft As Double
Dim percentTop As Double

percentLeft = 0.003 ' 10% t? trái
percentTop = 0.36 ' 10% t? trên

' Ð?t v? trí cho form con
Me.Left = FrmChungtu.Left + (FrmChungtu.Width * percentLeft)
Me.Top = FrmChungtu.Top + (FrmChungtu.Height * percentTop)
End Sub
 
Private Sub Timer1_Timer()
Timer1.Enabled = False
Hide
End Sub
