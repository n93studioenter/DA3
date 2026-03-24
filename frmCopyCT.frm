VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSCCT 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   LinkTopic       =   "Form4"
   ScaleHeight     =   825
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFakeTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   13575
      TabIndex        =   6
      Top             =   0
      Width           =   13575
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
         TabIndex        =   8
         Top             =   0
         Width           =   4455
      End
      Begin VB.Image picIcon 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   120
         Picture         =   "frmCopyCT.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   8550
         Index           =   0
         Left            =   840
         Picture         =   "frmCopyCT.frx":02BD
         Stretch         =   -1  'True
         Top             =   240
         Width           =   7890
      End
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
         Left            =   6240
         TabIndex        =   7
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Thùc hiÖn"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   350
      Width           =   1455
   End
   Begin VB.ComboBox CboThang 
      Height          =   315
      Index           =   1
      ItemData        =   "frmCopyCT.frx":BDDA
      Left            =   3480
      List            =   "frmCopyCT.frx":BDE1
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.ComboBox CboThang 
      Height          =   315
      Index           =   0
      ItemData        =   "frmCopyCT.frx":BDEF
      Left            =   960
      List            =   "frmCopyCT.frx":BDF6
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      VariousPropertyBits=   746604571
      Size            =   "3413;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "®Õn th¸ng"
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
      Left            =   2640
      TabIndex        =   2
      Tag             =   "to"
      Top             =   400
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "Tõ th¸ng"
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
      Index           =   6
      Left            =   120
      TabIndex        =   1
      Tag             =   "to"
      Top             =   400
      Width           =   735
   End
End
Attribute VB_Name = "frmSCCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_ktra164 As Recordset
Public thangbd As Integer
Public MaCT As Integer
Public Tuthang As Integer
Public denthang As Integer
Public tt As Integer
Dim TITLE_HEIGHT As Long      ' twips


Private Sub Command1_Click()

    Tuthang = val(Replace(LCase$(CboThang(0).Text), "tháng", ""))
    denthang = val(Replace(LCase$(CboThang(1).Text), "tháng", ""))
    'Kiem tra xem loai ngay cua chung tu la cuoi thang ko
    Dim Query64 As String
    Query64 = "SELECT * FROM ChungTu WHERE MaCT=" & MaCT & ""
    Set rs_ktra164 = DBKetoan.OpenRecordset(Query64, dbOpenSnapshot)
    If Not rs_ktra164.EOF Then
        'Neu la tong hop
        If rs_ktra164!maloai = 0 Or rs_ktra164!maloai = 4 Then
            Dim lastDay As Date
            lastDay = DateSerial(Year(rs_ktra164!NgayCT), month(rs_ktra164!NgayCT) + 1, 0)
            Dim GetDate As Date
            Dim getsh As String
            Dim getDG As String
            If lastDay = rs_ktra164!NgayCT Then
                GetDate = DateSerial(Year(rs_ktra164!NgayCT), Tuthang, 0)
            Else
                GetDate = DateSerial(Year(rs_ktra164!NgayCT), Tuthang - 1, 1)
            End If
            FrmChungtu.bakNgay = GetDate
            'Tang so hieu hien tai5 len den gatedate
            Dim currentThang As Integer
            currentThang = month(rs_ktra164!NgayCT)
            Dim i As Integer
            getsh = rs_ktra164!sohieu
            TextBox1.Text = VniToUnicode(rs_ktra164!diengiai)
            getDG = TextBox1.Text
            For i = currentThang To (Tuthang - 1)
                getsh = FrmChungtu.NextCode(getsh)
                Dim Test As String
                Test = TextBox1.Text
                getDG = FrmChungtu.NextMonthText(TextBox1.Text)
            Next i
            FrmChungtu.bakSoHieu = getsh
            FrmChungtu.bakNoidung = getDG
        End If
    End If


    FrmChungtu.bodemafter = Tuthang
    FrmChungtu.Afterclick
    Unload Me
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
Public Sub AnControl(frm As Form)
    Dim ctl As Control

    For Each ctl In frm.Controls
        Select Case TypeName(ctl)
        Case "Label", "TextBox", "ComboBox", "PictureBox", _
             "CommandButton", "Frame", "CheckBox", _
             "OptionButton", "ListBox", "Grid", _
             "MSHFlexGrid", "DataGrid", "Outline", "Line", "SSTab", "MaskEdBox"

            If ctl.Name <> "picFakeTitle" _
               And ctl.Name <> "lblTitle" _
               And ctl.Name <> "lblClose" Then

                If TypeName(ctl) = "Line" Then
                    ctl.y1 = ctl.y1 + TITLE_HEIGHT
                    ctl.y2 = ctl.y2 + TITLE_HEIGHT
                Else
                    ctl.Top = ctl.Top + TITLE_HEIGHT
                End If

            End If
        End Select
    Next
End Sub
Private Sub Form_Load()
    TITLE_HEIGHT = 200
    lblTitle(11).AutoSize = True
    Me.Height = Me.Height + 350 + 10
    picFakeTitle.Width = Me.ScaleWidth
    picFakeTitle.Height = 300
    picIcon(1).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2
    lblTitle(11).Left = picIcon(1).Left + picIcon(1).Width + 90
    lblTitle(11).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2 + 15
    lblClose.Top = 55
    lblTitle(11).Caption = "Nh©n b¶n chøng tõ"
    AnControl Me
    With CboThang(0)
        .Clear
        .AddItem "Tháng 1"
        .AddItem "Tháng 2"
        .AddItem "Tháng 3"
        .AddItem "Tháng 4"
        .AddItem "Tháng 5"
        .AddItem "Tháng 6"
        .AddItem "Tháng 7"
        .AddItem "Tháng 8"
        .AddItem "Tháng 9"
        .AddItem "Tháng 10"
        .AddItem "Tháng 11"
        .AddItem "Tháng 12"
        .ListIndex = 0    ' Ch?n tháng 1 m?c d?nh
    End With
    With CboThang(1)
        .Clear
        .AddItem "Tháng 1"
        .AddItem "Tháng 2"
        .AddItem "Tháng 3"
        .AddItem "Tháng 4"
        .AddItem "Tháng 5"
        .AddItem "Tháng 6"
        .AddItem "Tháng 7"
        .AddItem "Tháng 8"
        .AddItem "Tháng 9"
        .AddItem "Tháng 10"
        .AddItem "Tháng 11"
        .AddItem "Tháng 12"
        .ListIndex = 0    ' Ch?n tháng 1 m?c d?nh
    End With

    If thangbd <> 12 Then
        CboThang(0).ListIndex = thangbd
        CboThang(1).ListIndex = 11
    End If
End Sub

