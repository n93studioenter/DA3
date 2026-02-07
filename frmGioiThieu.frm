VERSION 5.00
Begin VB.Form frmgioithieu 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   105
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGioiThieu.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5460
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFakeTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   13575
      TabIndex        =   13
      Top             =   0
      Width           =   13575
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9360
         TabIndex        =   15
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "§¨ng nhËp"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   11
         Left            =   600
         TabIndex        =   14
         Top             =   0
         Width           =   4455
      End
      Begin VB.Image picIcon 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   120
         Picture         =   "frmGioiThieu.frx":57E2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   8550
         Index           =   0
         Left            =   840
         Picture         =   "frmGioiThieu.frx":5A9F
         Stretch         =   -1  'True
         Top             =   240
         Width           =   7890
      End
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   8280
      Picture         =   "frmGioiThieu.frx":115BC
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "&Return"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "PhÇn mÒm hoµn toµn miÔn phÝ cho doanh nghiÖp nhá, cã sè chøng tõ nhá h¬n 500 vµ doanh thu nhá h¬n  05 tû."
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   2280
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- Xin vui lßng liªn hÖ  090 575 7799 Mr. H­ng. "
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2520
      TabIndex        =   11
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"frmGioiThieu.frx":129DE
      Height          =   615
      Index           =   6
      Left            =   360
      TabIndex        =   10
      Top             =   3840
      Width           =   9135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Giíi thiÖu"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Index           =   5
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   9135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- Lµm ®­îc nhiÒu c«ng ty trªn mét phÇn mÒm, thªm c«ng ty míi ®¬n gi¶n, phï hîp cho ng­êi dïng lµm dÞch vô kÕ to¸n."
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- ChuyÓn trùc tiÕp c¸c b¸o c¸o ra m· v¹ch."
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- Më c¸c tµi kho¶n chi tiÕt ®¬n gi¶n, ®¨ng ký theo dâi ®èi t­îng nhanh."
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- PhÇn mÒm chØ nhËp mét lÇn lµ ra tÊt c¶ c¸c sæ, tæng hîp, chi tiÕt, vËt t­, c«ng nî, c«ng tr×nh..."
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- Tæng hîp, ng©n hµng, vËt t­, b¸n hµng, TSC§, dông cô, kÕt chuyÓn tù ®éng..."
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "PhÇn mÒm gåm c¸c ph©n hÖ:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   7575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"frmGioiThieu.frx":12A91
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   9015
   End
End
Attribute VB_Name = "frmgioithieu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Height = Me.Height + 350 + 10
    picFakeTitle.Width = Me.ScaleWidth
    picFakeTitle.Height = 325
    picIcon(1).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2
    lblTitle(11).Left = picIcon(1).Left + picIcon(1).Width + 90
    lblTitle(11).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2 + 15
    lblClose.Top = 55
    AnControl Me
    Caption = "Giíi thiÖu"
    lblTitle(11).Caption = Caption
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub
Private Sub picFakeTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
Private Sub lblTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picFakeTitle_MouseDown Button, Shift, X, Y
End Sub

Private Sub Command_Click(Index As Integer)
Unload Me
End Sub

