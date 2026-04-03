VERSION 5.00
Begin VB.Form FrmGetStr2 
   BackColor       =   &H00FFFFC0&
   Caption         =   """NhËp m· sè cµi ®Æt"""
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   LinkTopic       =   "Form4"
   ScaleHeight     =   1200
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   5
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   9
      Top             =   360
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   4
      Left            =   4680
      MaxLength       =   6
      TabIndex        =   7
      Top             =   360
      Width           =   1000
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   100
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   960
      MaxLength       =   13
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   2
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   2
      Top             =   360
      Width           =   1000
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   3
      Left            =   3600
      MaxLength       =   6
      TabIndex        =   1
      Top             =   360
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      Height          =   255
      Left            =   6840
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "NhËp m· sè cµi ®Æt"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Liªn hÖ nhµ cung cÊp ®Ó cã m· sè cµi ®Æt"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3495
   End
End
Attribute VB_Name = "FrmGetStr2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Clipboard.Clear
    Clipboard.SetText Trim(Text1.Text)
End Sub

Private Sub Text_Change(Index As Integer)
    Dim a() As String
    On Error GoTo Error_Handler
    If (Len(Text(0).Text) > 13) Then
        a = Split(Text(0).Text, "*")
        Text(0).Text = a(0)
        Text(1).Text = a(1)
        Text(2).Text = a(2)
        Text(3).Text = a(3)
        Dim a2() As String
        a2 = Split(a(4), "-")
        Text(4).Text = a2(0)
        Text(5).Text = a2(1)
    End If
Error_Handler:
End Sub
Private Sub Form_Load()
    Me.KeyPreview = True
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13: Me.Hide

    Case 27: Text(0).Text = ""
        Me.Hide
    End Select
End Sub
