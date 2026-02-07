VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form GetDate2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   555
   ClientLeft      =   3090
   ClientTop       =   3285
   ClientWidth     =   3675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GetDate2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFakeTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   13575
      TabIndex        =   4
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
         Left            =   3240
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   0
         Width           =   4455
      End
      Begin VB.Image picIcon 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   120
         Picture         =   "GetDate2.frx":57E2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   8550
         Index           =   0
         Left            =   840
         Picture         =   "GetDate2.frx":5A9F
         Stretch         =   -1  'True
         Top             =   240
         Width           =   7890
      End
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "®Õn ngµy"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Tag             =   "to"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tõ ngµy"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Tag             =   "From"
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "GetDate2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ngay(0 To 1) As Date
Dim esc As Integer

Public Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27:
            esc = 1
            Unload Me
    End Select
End Sub
Private Sub lblClose_Click()
    Unload Me
End Sub
Private Sub picFakeTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
Private Sub lblTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picFakeTitle_MouseDown Button, Shift, X, Y
End Sub
Private Sub Form_Load()
    Caption = " TÝnh gi¸ vèn"
    lblTitle(11).Caption = Caption
    lblTitle(11).AutoSize = True
    Me.Height = Me.Height + 350 + 10
    picFakeTitle.Width = Me.ScaleWidth
    picFakeTitle.Height = 325
    picIcon(1).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2
    lblTitle(11).Left = picIcon(1).Left + picIcon(1).Width + 90
    lblTitle(11).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2 + 15
    lblClose.Top = 55
    AnControl Me
    Dim chi_so As Integer
    esc = 0
    For chi_so = 0 To 1
        InitDateVars MedNgay(chi_so), ngay(chi_so)
    Next
End Sub

Private Sub MedNgay_GotFocus(Index As Integer)
    AutoSelect MedNgay(Index)
End Sub

Public Sub MedNgay_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Index
            Case 0: RFocus MedNgay(1)
            Case 1:
                MedNgay_LostFocus 1
                Unload Me
        End Select
    End If
End Sub

Private Sub MedNgay_LostFocus(Index As Integer)
    If IsDate(MedNgay(Index).Text) Then
        ngay(Index) = CVDate(MedNgay(Index).Text)
    Else
        RFocus MedNgay(Index)
    End If
End Sub

Public Function GetDate(s As String, d1 As Date, d2 As Date, Optional setndau As Integer = 0) As Boolean
    Me.Caption = s
    If setndau > 0 Then
        ngay(0) = d1
        MedNgay(0).Text = Format(d1, Mask_D)
        MedNgay(0).Enabled = False
    End If
    Me.Show vbModal
    If esc = 1 Then
        GetDate = False
    Else
        d1 = ngay(0)
        d2 = ngay(1)
        GetDate = True
    End If
    lblTitle(11).Caption = Caption
End Function

Public Function GetDate1(s As String, d1 As Date) As Boolean
    Me.Caption = s
    MedNgay(1).Visible = False
    Me.Show vbModal
    If esc = 1 Then
        GetDate1 = False
    Else
        d1 = ngay(0)
        GetDate1 = True
    End If
    lblTitle(11).Caption = Caption
End Function

