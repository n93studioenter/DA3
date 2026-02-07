VERSION 5.00
Begin VB.Form frmFontsize 
   Caption         =   "Form4"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   7800
   LinkTopic       =   "Form4"
   ScaleHeight     =   5655
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2880
      Width           =   5415
   End
   Begin VB.ComboBox CboLoai 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2280
      Width           =   5415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Bold"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Width           =   5415
   End
   Begin VB.ComboBox CboLoai 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Control  size"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Control  Font"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7320
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label2 
      Caption         =   "Menu size"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Menu Font"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmFontsize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
