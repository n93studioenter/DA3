VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmKhachHang 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   7080
   ClientLeft      =   1035
   ClientTop       =   675
   ClientWidth     =   11010
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmKhachHang.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Liability Items"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.PictureBox picFakeTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   13575
      TabIndex        =   52
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
         TabIndex        =   54
         Top             =   0
         Width           =   4455
      End
      Begin VB.Image picIcon 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   120
         Picture         =   "FrmKhachHang.frx":57E2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   8550
         Index           =   0
         Left            =   840
         Picture         =   "FrmKhachHang.frx":5A9F
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
         Left            =   10560
         TabIndex        =   53
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   315
      Left            =   4560
      TabIndex        =   51
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Left            =   360
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   6960
      MaxLength       =   20
      TabIndex        =   2
      Text            =   "..."
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtVT 
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   100
      TabIndex        =   4
      Text            =   "..."
      Top             =   840
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtVT 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   10
      Left            =   7440
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "FrmKhachHang.frx":115BC
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   6960
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "..."
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtVT 
      Height          =   285
      Index           =   2
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   6
      Text            =   "..."
      Top             =   1200
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   6960
      MaxLength       =   20
      TabIndex        =   8
      Text            =   "..."
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   9240
      MaxLength       =   20
      TabIndex        =   10
      Text            =   "..."
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   6960
      MaxLength       =   20
      TabIndex        =   9
      Text            =   "..."
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtVT 
      Height          =   285
      Index           =   7
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   12
      Text            =   "..."
      Top             =   3240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtVT 
      Height          =   285
      Index           =   8
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   14
      Text            =   "..."
      Top             =   3840
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtVT 
      Height          =   285
      Index           =   9
      Left            =   1080
      MaxLength       =   100
      TabIndex        =   18
      Text            =   "..."
      Top             =   1920
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.ComboBox CboNT 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FrmKhachHang.frx":115C0
      Left            =   9240
      List            =   "FrmKhachHang.frx":115C2
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   9480
      Picture         =   "FrmKhachHang.frx":115C4
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "&Return"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   8280
      Picture         =   "FrmKhachHang.frx":129E6
      Style           =   1  'Graphical
      TabIndex        =   30
      Tag             =   "&Delete"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   7080
      Picture         =   "FrmKhachHang.frx":13EC8
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "&Save"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   5880
      Picture         =   "FrmKhachHang.frx":152F6
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "&Add"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton SSCmdF 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   29
      Top             =   6640
      Width           =   255
   End
   Begin VB.OptionButton SSOpt 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tªn KH"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   27
      Tag             =   "Name"
      Top             =   6640
      Width           =   855
   End
   Begin VB.OptionButton SSOpt 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sè hiÖu"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Tag             =   "Code"
      Top             =   6640
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.ComboBox CboLoai 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.ListBox LstVt 
      Height          =   6105
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
   Begin VB.OptionButton SSOpt 
      BackColor       =   &H00FFFFC0&
      Caption         =   "MST"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   28
      Tag             =   "T. Code"
      Top             =   6640
      Width           =   700
   End
   Begin MSForms.TextBox txtSearch 
      Height          =   375
      Left            =   2880
      TabIndex        =   55
      Top             =   6600
      Width           =   1215
      VariousPropertyBits=   679495707
      Size            =   "2143;661"
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtTK 
      Height          =   320
      Left            =   6960
      TabIndex        =   13
      Top             =   3840
      Width           =   3615
      VariousPropertyBits=   679495707
      Size            =   "6376;564"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtDaiDien 
      Height          =   300
      Left            =   6960
      TabIndex        =   11
      Top             =   3360
      Width           =   3615
      VariousPropertyBits=   679495707
      Size            =   "6376;529"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtGhiChu 
      Height          =   300
      Left            =   6960
      TabIndex        =   17
      Top             =   5880
      Width           =   3615
      VariousPropertyBits=   679495707
      Size            =   "6376;529"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtDiaChi 
      Height          =   300
      Left            =   6960
      TabIndex        =   5
      Top             =   1440
      Width           =   3615
      VariousPropertyBits=   679495707
      Size            =   "6376;529"
      SpecialEffect   =   0
      FontName        =   "Times New Roman"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtName 
      Height          =   300
      Left            =   6960
      TabIndex        =   3
      Top             =   960
      Width           =   3615
      VariousPropertyBits=   679495707
      Size            =   "6376;529"
      SpecialEffect   =   0
      FontName        =   "Times New Roman"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "D­ cã"
      Height          =   255
      Index           =   19
      Left            =   8880
      TabIndex        =   50
      Tag             =   "Current Balance"
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "D­ nî"
      Height          =   255
      Index           =   18
      Left            =   6480
      TabIndex        =   49
      Tag             =   "Current Balance"
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label LbTon 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   7440
      TabIndex        =   48
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè hiÖu"
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   47
      Tag             =   "Code"
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tªn"
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   46
      Tag             =   "Name"
      Top             =   960
      Width           =   735
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   6960
      X2              =   8280
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line Line 
      Index           =   1
      X1              =   6960
      X2              =   10560
      Y1              =   1285
      Y2              =   1285
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "§Þa chØ"
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   45
      Tag             =   "Address"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè d­ hiÖn thêi"
      Height          =   255
      Index           =   8
      Left            =   6120
      TabIndex        =   44
      Tag             =   "Current Balance"
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label LbTon 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   43
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Line Line 
      Index           =   2
      X1              =   6960
      X2              =   10560
      Y1              =   1755
      Y2              =   1755
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè d­ tèi ®a"
      Height          =   255
      Index           =   12
      Left            =   6120
      TabIndex        =   42
      Tag             =   "Credit"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Line Line 
      Index           =   6
      X1              =   7440
      X2              =   8520
      Y1              =   4605
      Y2              =   4605
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MS ThuÕ"
      Height          =   255
      Index           =   10
      Left            =   6120
      TabIndex        =   41
      Tag             =   "Tax Code"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Line Line 
      Index           =   3
      X1              =   6960
      X2              =   8280
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Line Line 
      Index           =   4
      X1              =   6960
      X2              =   8280
      Y1              =   2685
      Y2              =   2685
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tel"
      Height          =   255
      Index           =   0
      Left            =   6120
      TabIndex        =   40
      Top             =   2400
      Width           =   375
   End
   Begin VB.Line Line 
      Index           =   5
      X1              =   9240
      X2              =   10560
      Y1              =   2685
      Y2              =   2685
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fax"
      Height          =   255
      Index           =   11
      Left            =   8640
      TabIndex        =   39
      Top             =   2400
      Width           =   375
   End
   Begin VB.Line Line 
      Index           =   7
      X1              =   6960
      X2              =   8280
      Y1              =   3165
      Y2              =   3165
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email"
      Height          =   255
      Index           =   13
      Left            =   6120
      TabIndex        =   38
      Top             =   2880
      Width           =   615
   End
   Begin VB.Line Line 
      Index           =   8
      X1              =   6960
      X2              =   10560
      Y1              =   3675
      Y2              =   3675
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "§¹i diÖn"
      Height          =   255
      Index           =   14
      Left            =   6120
      TabIndex        =   37
      Tag             =   "Representative"
      Top             =   3360
      Width           =   735
   End
   Begin VB.Line Line 
      Index           =   9
      X1              =   6960
      X2              =   10560
      Y1              =   4155
      Y2              =   4155
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  T.kho¶n"
      Height          =   255
      Index           =   6
      Left            =   6000
      TabIndex        =   36
      Tag             =   "Bank Acc."
      Top             =   3840
      Width           =   855
   End
   Begin VB.Line Line 
      Index           =   10
      X1              =   6960
      X2              =   10560
      Y1              =   6195
      Y2              =   6195
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ghi chó"
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   35
      Tag             =   "Notes"
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Theo dâi b»ng"
      Height          =   255
      Index           =   15
      Left            =   9000
      TabIndex        =   34
      Tag             =   "by Currency"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nguyªn tÖ"
      Height          =   255
      Index           =   16
      Left            =   6480
      TabIndex        =   33
      Tag             =   "F. Currency"
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label LbTon 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   9480
      TabIndex        =   32
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6375
      Index           =   17
      Left            =   5640
      TabIndex        =   31
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   5
      Left            =   5880
      TabIndex        =   24
      Top             =   180
      Width           =   4935
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   4
      Left            =   5880
      TabIndex        =   23
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label 
      BackColor       =   &H00808080&
      Height          =   5955
      Index           =   9
      Left            =   165
      TabIndex        =   22
      Top             =   540
      Width           =   5295
   End
End
Attribute VB_Name = "FrmKhachHang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_KH As Recordset
Dim ThemMoi As Integer          ' =1 neu them moi, -1 neu sua cu
Dim okh As New ClsKhachHang       ' vat tu duoc tham chieu
Dim doiloai As Integer               ' =1 neu co thay doi loai vat tu dang sua doi
Dim MaDaTim As Long
Dim xT As Integer
Dim xSH As String
'======================================================================================
' Liet ke cac vat tu trong loai vat tu duoc chon
'======================================================================================
Private Sub CboLoai_Click()
    If ThemMoi <> -1 Then
        Me.MousePointer = 11
        Int_RecsetToCbo "SELECT MaSo As F2, SoHieu + Chr(9) + Ten As F1 FROM KhachHang WHERE MaPhanLoai=" + CStr(CboLoai.ItemData(CboLoai.ListIndex)) + " ORDER BY SoHieu", LstVt
        ThemMoi = 0
        doiloai = 0
        Me.MousePointer = 0
    Else
        doiloai = 1
    End If
End Sub


Public Sub Command_Click(Index As Integer)
    Dim vt As New ClsKhachHang, i As Integer
    
    If (User_Right = 2) And (Index < 3) Then
        HienThongBao "Kh«ng cã quyÒn truy cËp!", 1
        GoTo XongVT
    End If
    
    Me.MousePointer = 11
    If Index < 3 Then
        If CboLoai.ListIndex < 0 Then
                ErrMsg er_PhanLoai
                GoTo XongVT
        End If
    End If
    
    Select Case Index
        Case 0:
            txtVT(0).Text = SoHieuVTMoi(CboLoai.ItemData(CboLoai.ListIndex), 2)
            For i = 1 To 9
                txtVT(i).Text = "..."
            Next
            CboNT.ListIndex = 0
            RFocus txtVT(0)
            ThemMoi = 1
        Case 1:
            Select Case ThemMoi
                Case 1:
                    If Not KiemTraSoLieu Then GoTo XongVT
                    If okh.GhiKhachHang = 0 Then
                        LstVt.AddItem okh.sohieu + Chr(9) + okh.Ten
                        LstVt.ItemData(LstVt.NewIndex) = okh.MaSo
                        LstVt.ListIndex = LstVt.NewIndex
                    Else
                        ErrMsg er_SoHieu
                        vt.InitKhachHangSohieu txtVT(0).Text
                        If vt.MaPhanLoai = CboLoai.ItemData(CboLoai.ListIndex) Then
                            SetListIndex LstVt, vt.MaSo
                        End If
                    End If
                    ThemMoi = 0
                Case 0:
                    If LstVt.ListIndex < 0 Then GoTo XongVT
                    If Not KiemTraSoLieu Then GoTo XongVT
                    
                    If okh.SuaKH = 0 Then
                        If doiloai = 1 Then
                            CboLoai_Click
                            doiloai = 0
                        Else
                            LstVt.List(LstVt.ListIndex) = okh.sohieu + Chr(9) + okh.Ten
                        End If
                    Else
                        vt.InitKhachHangSohieu txtVT(0).Text
                        ErrMsg er_SoHieu
                        If vt.MaPhanLoai = CboLoai.ItemData(CboLoai.ListIndex) Then SetListIndex LstVt, vt.MaSo
                    End If
                    ThemMoi = 0
            End Select
            RFocus LstVt
        Case 2:
            i = LstVt.ListIndex
            If i < 0 Then GoTo XongVT
            If okh.XoaKH = 0 Then
                LstVt.RemoveItem i
                If LstVt.ListCount > 0 Then LstVt.ListIndex = i - 1
            Else
                ErrMsg er_CoPS1
            End If
            RFocus LstVt
        Case 3:
            Hide
    End Select
XongVT:
    Set vt = Nothing
    Me.MousePointer = 0
End Sub

Private Sub Command1_Click()
    LstVt.Clear
    Dim Query As String
    Dim rs_KH As DAO.Recordset

    ' Ki?m tra xem h?p van b?n có d? li?u không
    If Len(Trim(txtF.Text)) = 0 Then
        MsgBox "Vui lòng nh?p tên khách hàng d? tìm ki?m.", vbExclamation
        Exit Sub
    End If

    ' T?o truy v?n v?i di?u ki?n LIKE
    Query = "SELECT * FROM KhachHang WHERE Ten LIKE '*" & txtF.Text & "*'"

    ' M? Recordset
    Set rs_KH = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

    ' Ki?m tra xem Recordset có d? li?u không
    If Not rs_KH.EOF Then
        ' Duy?t danh sách b?ng Do While
        Do While Not rs_KH.EOF
            ' L?y giá tr? t? m?t tru?ng c? th?, ví d? "MaKH"
            'MsgBox rs_KH!Ten
            LstVt.AddItem rs_KH!sohieu + Chr(9) + rs_KH!Ten
            LstVt.ItemData(LstVt.NewIndex) = rs_KH!MaSo
            'LstVt.ListIndex = LstVt.NewIndex
            ' Di chuy?n d?n b?n ghi ti?p theo
            rs_KH.MoveNext
        Loop
    Else
        MsgBox "Không tìm th?y khách hàng nào phù h?p.", vbInformation
    End If

    ' Ðóng Recordset
    rs_KH.Close
    Set rs_KH = Nothing
End Sub

Private Sub Form_Activate()
    If Me.tag < 0 Then
        SetListIndex CboLoai, -Me.tag
        Me.tag = 0
    End If
    If ThemMoi = 0 And Me.tag = 1 Then RFocus LstVt
    If xT = 1 Then
        If xSH <> "" Then SetListIndex CboLoai, LayMaPhanLoai(xSH, "KhachHang")
        Command_Click 0
        txtVT(0).Text = xSH
    End If
End Sub

'Private Sub Form_Deactivate()
'    If CmdD.tag <> 0 Then CmdD_Click
'End Sub

'======================================================================================
' Xu ly cac phim nong
'======================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyV:
                RFocus Command(3)
                Command_Click 3
            Case vbKeyT:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyX:
                RFocus Command(2)
                Command_Click 2
            Case vbKeyG:
                RFocus Command(1)
                Command_Click 1
        End Select
    End If
    If KeyCode = vbKeyEscape Then Hide
End Sub

Private Sub picFakeTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
Private Sub lblTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picFakeTitle_MouseDown Button, Shift, X, Y
End Sub
Private Sub lblClose_Click()
    Unload Me
End Sub
Private Sub txtName_Change()
    txtVT(1).Text = UnicodeToVni(txtName.Text)
End Sub
Private Sub txtDiaChi_Change()
    txtVT(2).Text = UnicodeToVni(txtDiaChi.Text)
End Sub
Private Sub txtGhichu_Change()
    txtVT(9).Text = UnicodeToVni(txtGhiChu.Text)
End Sub
Private Sub txtDaiDien_Change()
    txtVT(7).Text = UnicodeToVni(txtDaiDien.Text)
End Sub
Private Sub txtTK_Change()
    txtVT(8).Text = UnicodeToVni(txtTK.Text)
End Sub

'======================================================================================
' Khoi tao form
'======================================================================================
Private Sub txtSearch_Change()
    txtF.Text = UnicodeToVni(txtSearch.Text)
End Sub
Private Sub Form_Load()
    lblTitle(11).AutoSize = True
    Me.Height = Me.Height + 350 + 10
    picFakeTitle.Width = Me.ScaleWidth
    picFakeTitle.Height = 325
    picIcon(1).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2
    lblTitle(11).Left = picIcon(1).Left + picIcon(1).Width + 90
    lblTitle(11).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2 + 15
    lblClose.Top = 80
    AnControl Me



    ThemMoi = 0
    doiloai = 0
    Caption = "Danh s¸ch kh¸ch hµng" + " - " + CStr(pNamTC)
    'Int_RecsetToCbo "SELECT DISTINCTROW MaSo As F2,SoHieu + ' - '  + TenPhanLoai As F1 FROM PhanLoaiKhachHang WHERE PLCon=0 AND LEFT(SoHieu,1)<>'#' ORDER BY SoHieu", CboLoai
    Int_RecsetToCbo "SELECT DISTINCTROW MaSo As F2,SoHieu + ' - '  + TenPhanLoai As F1 FROM PhanLoaiKhachHang WHERE PLCon=0  ORDER BY SoHieu", CboLoai
    Int_RecsetToCbo "SELECT MaSo As F2,KyHieu As F1 FROM NguyenTe WHERE KyHieu<>'" + pTienStr + "' ORDER BY KyHieu", CboNT
    CboNT.AddItem pTienStr, 0
    CboNT.ItemData(0) = 0
    CboNT.ListIndex = 0
    lblTitle(11).Caption = Caption
    SetFont Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set okh = Nothing
End Sub

'======================================================================================
' Khoi tao vat tu duoc chon
'======================================================================================
Private Sub LstVt_Click()
    okh.InitKhachHangMaSo LstVt.ItemData(LstVt.ListIndex)
    ShowChitiet okh
End Sub
'======================================================================================
' Thu tuc hien thong tin chi tiet
'======================================================================================
Private Sub ShowChitiet(otk As ClsKhachHang)
    Dim n As Double, c As Double, nt As Double
    
    txtVT(0).Text = okh.sohieu
    txtVT(1).Text = okh.Ten
    txtName.Text = VniToUnicode(okh.Ten)
    
    txtVT(2).Text = okh.DiaChi
    txtDiaChi.Text = VniToUnicode(okh.DiaChi)
    txtVT(3).Text = okh.mst
    txtVT(4).Text = okh.Tel
    txtVT(5).Text = okh.Fax
    txtVT(6).Text = okh.email
    txtVT(7).Text = okh.DaiDien
    txtDaiDien.Text = VniToUnicode(okh.DaiDien)
    txtVT(8).Text = okh.taikhoan
    txtTK.Text = VniToUnicode(okh.taikhoan)
    txtVT(9).Text = okh.GhiChu
    txtGhiChu.Text = VniToUnicode(okh.GhiChu)
    txtVT(10).Text = Format(okh.DuMax, Mask_0)
    SetListIndex CboNT, okh.MaNT
    okh.SoDuKH ThangCuoiNamTC, n, c, nt
    If n - c >= 0 Then
        n = n - c
        c = 0
    Else
        c = c - n
        n = 0
    End If
    LbTon(0).Caption = Format(n, Mask_0)
    LbTon(1).Caption = Format(c, Mask_0)
    LbTon(2).Caption = Format(nt, Mask_2)
End Sub
'======================================================================================
' Thu tuc chon vat tu
' sh: so hieu vat tu can chon
' Tra ve so hieu vat tu duoc chon
'======================================================================================
Public Function ChonKhachHang(sh As String) As String
    Dim mpl As Long, shtk As String
    Dim j As Integer, i As Integer, pos As Integer, Length As Integer
    
    If Len(sh) > 0 Then
        shtk = "SELECT DISTINCTROW TOP 1 MaPhanLoai AS F1 FROM KhachHang WHERE SoHieu LIKE '" + sh + "*' ORDER BY SoHieu"
        mpl = SelectSQL(shtk)
        If mpl > 0 And CboLoai.ListIndex >= 0 Then
            If CboLoai.ItemData(CboLoai.ListIndex) <> mpl Then SetListIndex CboLoai, mpl
        End If
         i = 0
         j = LstVt.ListCount - 1
         pos = 0
         Length = Len(sh)
         Do While i <= j - 1
                pos = Fix(0.5 + (i + j) / 2)
                shtk = Left(LstVt.List(pos), Length)
                If UCase(sh) = UCase(shtk) Then
                    i = pos - 1
                    Do While (UCase(sh) = UCase(Left(LstVt.List(i), Length))) And (i > 0)
                        i = i - 1
                    Loop
                    pos = i + 1
                    Exit Do
                End If
                If UCase(sh) > UCase(shtk) Then
                    i = pos
                Else
                    If j = 1 Then
                        pos = 0
                        Exit Do
                    Else
                        If j = pos Then Exit Do
                        j = pos
                    End If
                End If
        Loop
        If LstVt.ListCount > 0 Then LstVt.ListIndex = pos
    End If
    Me.tag = 1
    On Error Resume Next
    Me.Show 1
    On Error GoTo 0
    If okh.MaSo > 0 Then
        ChonKhachHang = okh.sohieu
    Else
        ChonKhachHang = ""
    End If
End Function

Private Sub LstVt_DblClick()
    If Me.tag = 1 Then Hide
End Sub

Private Sub LstVt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then LstVt_DblClick
End Sub
'======================================================================================
' Thu tuc kiem tra va dua so lieu vao object
'======================================================================================
'Private Function KiemTraSoLieu() As Boolean
Public Function KiemTraSoLieu() As Boolean
    KiemTraSoLieu = False
    
With okh
    If ThemMoi = 1 Then .MaSo = 0
    .MaPhanLoai = CboLoai.ItemData(CboLoai.ListIndex)
    .sohieu = txtVT(0).Text
    .Ten = txtVT(1).Text
    .DiaChi = txtVT(2).Text
    .mst = txtVT(3).Text
    .Tel = txtVT(4).Text
    .Fax = txtVT(5).Text
    .email = txtVT(6).Text
    .DaiDien = txtVT(7).Text
    .taikhoan = txtVT(8).Text
    .GhiChu = txtVT(9).Text
    .DuMax = Cdbl5(txtVT(10).Text)
    .MaNT = CboNT.ItemData(CboNT.ListIndex)
    If .mst <> "..." And SelectSQL("SELECT MaSo AS F1 FROM KhachHang WHERE MST='" + .mst + "' AND MaSo<>" + CStr(.MaSo)) > 0 Then
       ' If MsgBox("M· sè thuÕ ®· cã, cho phÐp nhËp?", vbYesNo + vbCritical, App.ProductName) = vbNo Then Exit Function
       If MsgBox("M· sè thuÕ ®· cã, cho phÐp nhËp?", vbYesNo + vbCritical, App.ProductName) = vbNo Then Exit Function
    End If
End With
    KiemTraSoLieu = True
End Function

Private Sub LstVt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sh As String, m As Long
    
    If Button = 2 And LstVt.ListIndex >= 0 And ThemMoi = 0 Then
        sh = FrmGetStr.GetString("ChuyÓn " + VString(okh.sohieu + " - " + okh.Ten) + " sang ph©n lo¹i cã sè hiÖu:", App.ProductName)
        If Len(sh) > 0 Then
            m = SelectSQL("SELECT MaSo AS F1 FROM PhanLoaiKhachHang WHERE PLCon=0 AND SoHieu='" + sh + "'")
            If m > 0 And m <> okh.MaPhanLoai Then
                ExecuteSQL5 "UPDATE KhachHang SET MaPhanLoai=" + CStr(m) + " WHERE MaSo = " + CStr(okh.MaSo)
                CboLoai_Click
            End If
        End If
    End If
End Sub

Private Sub SSCmdF_Click()
    Dim sql As String
    
    If Len(txtF.Text) = 0 Then
        RFocus txtF
        Exit Sub
    End If
    
    Me.MousePointer = 11
    sql = "SELECT DISTINCTROW Top 1 SoHieu AS F1 FROM KhachHang WHERE MaSo>" + CStr(MaDaTim)
    If SSOpt(0).Value Then sql = sql + " AND SoHieu LIKE '" + txtF.Text + "*'"
    If SSOpt(1).Value Then sql = sql + " AND InStr(Ten,'" + txtF.Text + "')>0"
    If SSOpt(2).Value Then sql = sql + " AND MST LIKE '" + txtF.Text + "*'"
    sql = CStr(SelectSQL(sql))
    If sql <> "0" Then
        ChonKhachHang sql
        MaDaTim = okh.MaSo
    Else
        MaDaTim = 0
    End If
    Me.MousePointer = 0
End Sub

Private Sub txtF_GotFocus()
    AutoSelect txtF
    MaDaTim = 0
End Sub

Private Sub Txtvt_GotFocus(Index As Integer)
    AutoSelect txtVT(Index)
End Sub

Private Sub TxtVT_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0: If KeyAscii = 32 Or KeyAscii = 35 Or KeyAscii = 39 Or KeyAscii = 42 Then KeyAscii = 0
        Case 10, 11, 12: KeyProcess txtVT(Index), KeyAscii
    End Select
End Sub

Private Sub TxtVT_LostFocus(Index As Integer)
    Select Case Index
        Case 0:
            txtVT(0).Text = UCase(txtVT(0).Text)
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
            If Len(txtVT(Index).Text) = 0 Then txtVT(Index).Text = "..."
        Case 10, 11, 12:
            txtVT(Index).Text = Format(txtVT(Index).Text, Mask_2)
    End Select
End Sub

Public Function ThemKhachHang(sh As String) As String
    If xT = 1 Then Exit Function
    Me.tag = 1
    xT = 1
    xSH = sh
    Me.Show 1
    xT = 0
    xSH = ""
    ThemKhachHang = okh.sohieu
End Function

