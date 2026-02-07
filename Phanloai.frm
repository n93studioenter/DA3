VERSION 5.00
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPhanLoai 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   5325
   ClientLeft      =   1560
   ClientTop       =   720
   ClientWidth     =   6615
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
   Icon            =   "Phanloai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Classification"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5325
   ScaleWidth      =   6615
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
      TabIndex        =   7
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
         Left            =   6120
         TabIndex        =   9
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   8550
         Index           =   0
         Left            =   840
         Picture         =   "Phanloai.frx":57E2
         Stretch         =   -1  'True
         Top             =   240
         Width           =   7890
      End
      Begin VB.Image picIcon 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   120
         Picture         =   "Phanloai.frx":112FF
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
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
         TabIndex        =   8
         Top             =   0
         Width           =   4455
      End
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   3
      Left            =   5400
      Picture         =   "Phanloai.frx":115BC
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Return"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   2
      Left            =   5400
      Picture         =   "Phanloai.frx":129DE
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "&Delete"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   5400
      Picture         =   "Phanloai.frx":13EC0
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "&Save"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   0
      Left            =   5400
      Picture         =   "Phanloai.frx":152EE
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "&Add"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   0
      Left            =   5520
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text 
      Height          =   320
      Index           =   1
      Left            =   120
      MaxLength       =   15
      TabIndex        =   0
      Top             =   4920
      Width           =   1155
   End
   Begin MSOutl.Outline Outline 
      Height          =   4455
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   7858
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Phanloai.frx":16848
      Style           =   2
      PicturePlus     =   "Phanloai.frx":16864
      PictureMinus    =   "Phanloai.frx":1695E
      PictureLeaf     =   "Phanloai.frx":16A58
      PictureOpen     =   "Phanloai.frx":16B52
      PictureClosed   =   "Phanloai.frx":16C4C
   End
   Begin MSForms.TextBox txtName 
      Height          =   320
      Left            =   1440
      TabIndex        =   10
      Top             =   4920
      Width           =   3735
      VariousPropertyBits=   679495707
      Size            =   "6588;564"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   3
      X1              =   5280
      X2              =   5280
      Y1              =   120
      Y2              =   4800
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   120
      X2              =   5280
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   5280
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "frmPhanLoai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type tpPhanLoai
      MaSo As Long
      Ten As String
      sohieu As String
      CapTren As Long
      cap As Integer
End Type
Dim PhanLoai As tpPhanLoai
Dim pPhanLoai As Integer
Dim TenBang As String
Dim tmpSoHieu As String                       ' L­u c¸c th«ng tin cña cÊp trªn (®Ó thªm míi)
Dim tmpMaSo As Long
Dim tmpCap As Integer
Dim sql As String
'======================================================================================
' FORM
'======================================================================================
' Activate
Private Sub Form_Activate()
    If Me.tag > 0 Then
        Select Case Me.tag
        Case 1:
            If pNN = 0 Then Me.Caption = "Ph©n lo¹i tµi s¶n"
            TenBang = "LoaiTaiSan"
        Case 2:
            If pNN = 0 Then Me.Caption = "Ph©n lo¹i chøng tõ"
            TenBang = "LoaiChungTu"
        End Select
        Caption = "Ph©n lo¹i" + " - " + CStr(pNamTC)
        Me.Refresh
        LayDanhSachPhanLoai
        pPhanLoai = Me.tag
        Me.tag = 0
        lblTitle(11).Caption = Caption
    End If
End Sub
' Key Down
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyInsert Then
            Outline.ListIndex = -1
            tmpMaSo = 0
            tmpCap = 0
            tmpSoHieu = ""
            RFocus Command(0)
            DoEvents
            Command_Click (0)
      End If
      If (Shift And vbAltMask) > 0 Then
            Select Case KeyCode
                  Case vbKeyT: RFocus Command(0): DoEvents: Command_Click (0)
                  Case vbKeyG:  RFocus Command(1): DoEvents: Command_Click (1)
                  Case vbKeyX: RFocus Command(2): DoEvents: Command_Click (2)
                  Case vbKeyV:  RFocus Command(3): DoEvents: Command_Click (3)
            End Select
      End If
      If KeyCode = vbKeyEscape Then Unload frmPhanLoai
End Sub
Private Sub lblClose_Click()
    Unload Me
End Sub
Private Sub txtName_Change()
    Text(0).Text = UnicodeToVni(txtName.Text)
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

    SetFont Me
End Sub

'======================================================================================
' OUTLINE
'======================================================================================
' Click
Private Sub Outline_Click()
      ' L­u d÷ liÖu cho yªu cÇu thªm míi hay xo¸
      tmpMaSo = Outline.ItemData(Outline.ListIndex)
      tmpCap = Outline.indent(Outline.ListIndex)
      sql = "SELECT SoHieu AS F1 FROM " + TenBang + " WHERE MaSo = " _
                                                                                                                                                                  + CStr(tmpMaSo)
      tmpSoHieu = CStr(SelectSQL(sql))
End Sub
' DblClick
Private Sub Outline_DblClick()
      ' ChØ ®Þnh ®èi t­îng theo m· sè
      ChiDinh Outline.ItemData(Outline.ListIndex)
      ' KiÓm tra cÊp cña ph©n lo¹i ®­îc chän (cÊp trªn cïng kh«ng thÓ söa ®æi hay xo¸)
'      If PhanLoai.cap = 1 Then
'            KhoiTao False
'      Else
            Text(0).Text = PhanLoai.Ten
            txtName.Text = VniToUnicode(PhanLoai.Ten)
            Text(1).Text = PhanLoai.sohieu
            RFocus Text(0)
'      End If
End Sub
' GotFocus
Private Sub Outline_GotFocus()
      KhoiTao False
End Sub
' KeyDown
Private Sub Outline_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then Outline_DblClick
End Sub
'======================================================================================
' command
'======================================================================================
Private Sub Command_Click(Index As Integer)
    Select Case Index
    Case 0      ' Míi
        KhoiTao True
    Case 1      ' Ghi
        If HopLe = 0 Then
            If PhanLoai.MaSo = 0 Then
                If ThemMoi = 0 Then KhoiTao True
            Else
                Dim vi_tri As Integer
                If SuaDoi = 0 Then
                    If pPhanLoai = 1 Then
                        SuaDoi
                        CapNhatSoHieu PhanLoai.cap, PhanLoai.MaSo, PhanLoai.sohieu
                    End If
                    vi_tri = Outline.ListIndex
                    LayDanhSachPhanLoai
                    Outline.ListIndex = vi_tri
                    KhoiTao False
                End If
            End If
        End If
    Case 2      ' Xo¸
        If Outline.ListIndex < 0 Then Exit Sub
        If Outline.indent(Outline.ListIndex) = 1 Then
            Beep
            'MsgBox "Kh«ng ®­îc phÐp xo¸ ph©n lo¹i cÊp trªn cïng", vbCritical
            Dim s As String
            s = ChrW(75) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(233) & ChrW(112) & ChrW(32) & ChrW(120) & ChrW(111) & ChrW(225) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(108) & ChrW(111) & ChrW(7841) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(7845) & ChrW(112) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(249) & ChrW(110) & ChrW(103)
            MessageBoxW Me.hWnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

        Else
            If vbNo = MsgBox("Xo¸ ph©n lo¹i hiÖn t¹i", vbYesNo + vbQuestion) Then Exit Sub
            If Outline.ListIndex + 1 < Outline.ListCount Then
                If Outline.indent(Outline.ListIndex + 1) > Outline.indent(Outline.ListIndex) Then
                    Beep
                    'MsgBox "VÉn cßn c¸c ph©n lo¹i cÊp d­íi", vbCritical
                    s = ChrW(86) & ChrW(7851) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(242) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(108) & ChrW(111) & ChrW(7841) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(7845) & ChrW(112) & ChrW(32) & ChrW(100) & ChrW(432) & ChrW(7899) & ChrW(105)
                    MessageBoxW Me.hWnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

                    Exit Sub
                End If
            End If
            If xoa = 0 Then KhoiTao False
        End If
    Case 3      ' Trë vÒ
        ' SendKeys "{Escape}", False
        Unload Me
    End Select
End Sub
'======================================================================================
' TEXT
'======================================================================================
' Got Focus
Private Sub Text_GotFocus(Index As Integer)
      AutoSelect Text(Index)
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If KeyAscii = 32 Or KeyAscii = 35 Or KeyAscii = 39 Or KeyAscii = 42 Then KeyAscii = 0
    End If
End Sub

' Lost Focus
Private Sub Text_LostFocus(Index As Integer)
      If Len(Text(Index).Text) = 0 Then Text(Index).Text = "(...)"
      Select Case Index
            Case 0:  PhanLoai.Ten = Text(0).Text
            Case 1:  PhanLoai.sohieu = Text(1).Text
      End Select
End Sub
'======================================================================================
' FUNCTION HopLe
'======================================================================================
Private Function HopLe()
    Dim thong_bao As String
    Dim s As String

    '      If tmpMaSo = 0 Then thong_bao = "Ch­a chØ ®Þnh ph©n lo¹i cÊp trªn": GoTo Err_InValidate
    With PhanLoai
        If Len(.Ten) = 0 Then Text_LostFocus (0)
        If Len(.sohieu) = 0 Then Text_LostFocus (1)
        If .Ten = "(...)" Then s = ChrW(84) & ChrW(104) & ChrW(105) & ChrW(7871) & ChrW(117) & ChrW(32) & ChrW(116) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(108) & ChrW(111) & ChrW(7841) & ChrW(105): GoTo Err_InValidate
        If .sohieu = "(...)" Then s = ChrW(84) & ChrW(104) & ChrW(105) & ChrW(7871) & ChrW(117) & ChrW(32) & ChrW(115) & ChrW(7889) & ChrW(32) & ChrW(104) & ChrW(105) & ChrW(7879) & ChrW(117) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(108) & ChrW(111) & ChrW(7841) & ChrW(105): GoTo Err_InValidate
        ' NÕu lµ thªm míi th× nhËn c¸c thuéc tÝnh cña ph©n lo¹i cÊp trªn
        If .MaSo = 0 Then
            .cap = tmpCap + 1
            .CapTren = tmpMaSo
            ' KiÓm tra cÊp vµ sè hiÖu
            If (pPhanLoai = 1 And .cap > 3) Or (pPhanLoai = 2 And .cap > 2) Then _
                   s = ChrW(83) & ChrW(7889) & ChrW(32) & ChrW(99) & ChrW(7845) & ChrW(112) & ChrW(32) & ChrW(118) & ChrW(432) & ChrW(7907) & ChrW(116) & ChrW(32) & ChrW(113) & ChrW(117) & ChrW(225) & ChrW(32) & ChrW(113) & ChrW(117) & ChrW(121) & ChrW(32) & ChrW(273) & ChrW(7883) & ChrW(110) & ChrW(104): GoTo Err_InValidate
            If Not Left(.sohieu, Len(tmpSoHieu)) = tmpSoHieu Then s = ChrW(83) & ChrW(7889) & ChrW(32) & ChrW(104) & ChrW(105) & ChrW(7879) & ChrW(117) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(273) & ChrW(250) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(113) & ChrW(117) & ChrW(121) & ChrW(32) & ChrW(273) & ChrW(7883) & ChrW(110) & ChrW(104): GoTo Err_InValidate
        Else
            Dim shieu_ctren As String
            ' KiÓm tra sè hiÖu
            sql = "SELECT SoHieu AS F1 FROM " + TenBang + " WHERE MaSo = " + CStr(.CapTren)
            shieu_ctren = CStr(SelectSQL(sql))
            If shieu_ctren <> "0" Then
                If Not Left(.sohieu, Len(shieu_ctren)) = shieu_ctren Then s = ChrW(83) & ChrW(7889) & ChrW(32) & ChrW(104) & ChrW(105) & ChrW(7879) & ChrW(117) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(273) & ChrW(250) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(113) & ChrW(117) & ChrW(121) & ChrW(32) & ChrW(273) & ChrW(7883) & ChrW(110) & ChrW(104): GoTo Err_InValidate
            End If
        End If
    End With
    HopLe = 0
    Exit Function
Err_InValidate:
    Beep
    'MsgBox thong_bao, vbCritical, App.ProductName
    MessageBoxW Me.hWnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

    HopLe = -1
End Function
'======================================================================================
' SUB ChiDinh
'======================================================================================
Private Sub ChiDinh(ma_pl As Long)
Dim rs_phanloai As Recordset
      sql = "SELECT * FROM " + TenBang + " WHERE MaSo = " + CStr(ma_pl)
      Set rs_phanloai = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
            PhanLoai.MaSo = rs_phanloai!MaSo
            PhanLoai.Ten = rs_phanloai!Ten
            PhanLoai.sohieu = rs_phanloai!sohieu
            PhanLoai.CapTren = rs_phanloai!CapTren
            PhanLoai.cap = rs_phanloai!cap
      rs_phanloai.Close
      Set rs_phanloai = Nothing
End Sub
'======================================================================================
' FUNCTION ThemMoi
'======================================================================================
Private Function ThemMoi() As Integer
Dim vi_tri As Integer
      Me.MousePointer = 11
      If ExecuteSQL5("INSERT INTO " + TenBang + " (MaSo,Ten, SoHieu, CapTren, Cap) VALUES(" + CStr(Lng_MaxValue("MaSo", TenBang) + 1) + ",'" _
            + PhanLoai.Ten + "','" + PhanLoai.sohieu + "'," + CStr(PhanLoai.CapTren) + "," + CStr(PhanLoai.cap) + ")") = 0 Then
            PhanLoai.MaSo = Lng_MaxValue("MaSo", TenBang)
            If PhanLoai.cap = 3 Then
                ExecuteSQL5 "UPDATE TaiSan SET MaNhom=" + CStr(PhanLoai.MaSo) + " WHERE MaNhom=0 AND MaLoai=" + CStr(PhanLoai.CapTren)
            End If
            Do                                                                                ' Thªm vµo vÞ trÝ cuèi cïng trong cÊp
                  vi_tri = vi_tri + 1
                  If Outline.ListIndex + vi_tri = Outline.ListCount Then Exit Do
            Loop Until Outline.indent(Outline.ListIndex + vi_tri) < PhanLoai.cap
            Outline.AddItem PhanLoai.sohieu + "  " + PhanLoai.Ten, Outline.ListIndex + vi_tri
            Outline.indent(Outline.ListIndex + vi_tri) = PhanLoai.cap
            Outline.ItemData(Outline.ListIndex + vi_tri) = PhanLoai.MaSo
            If Outline.ListIndex >= 0 Then Outline.Expand(Outline.ListIndex) = True      ' Expand cÊp trªn
            ThemMoi = 0
      Else
            PhanLoai.MaSo = 0
            ThemMoi = -1
      End If
      Me.MousePointer = 0
End Function
'======================================================================================
' FUNCTION SuaDoi
'======================================================================================
Private Function SuaDoi()
      Me.MousePointer = 11
      sql = "UPDATE " + TenBang + " SET Ten = '" + Text(0).Text _
                                                                                      + "', SoHieu = '" + PhanLoai.sohieu _
                                                                         + "' WHERE MaSo = " + CStr(PhanLoai.MaSo)
      If ExecuteSQL5(sql) = 0 Then
            Outline.List(Outline.ListIndex) = PhanLoai.sohieu + "  " + PhanLoai.Ten
            SuaDoi = 0
      Else
            SuaDoi = -1
      End If
      Me.MousePointer = 0
End Function
'======================================================================================
' FUNCTION Xoa
'======================================================================================
Private Function xoa() As Integer
      Me.MousePointer = 11
      If pPhanLoai = 1 And SelectSQL("SELECT DISTINCTROW Count(MaSo) AS F1 FROM TaiSan WHERE MaNhom = " + CStr(Outline.ItemData(Outline.ListIndex))) > 0 Then
            MsgBox "Ph©n lo¹i ®· ®¨ng ký tµi s¶n, kh«ng xo¸!", vbInformation, App.ProductName
            xoa = -1
      Else
            sql = "DELETE * FROM " + TenBang + " WHERE MaSo = " + CStr(Outline.ItemData(Outline.ListIndex))
            If ExecuteSQL5(sql) = 0 Then
                  If pPhanLoai = 1 Then ExecuteSQL5 "UPDATE TaiSan SET MaNhom = 0 WHERE MaNhom = " + CStr(Outline.ItemData(Outline.ListIndex))
                  Outline.RemoveItem Outline.ListIndex
                  xoa = 0
            Else
                  xoa = -1
            End If
      End If
      Me.MousePointer = 0
End Function
'======================================================================================
' SUB CapNhatSoHieu
'======================================================================================
Private Sub CapNhatSoHieu(cap_ct As Integer, maso_ct As Long, sohieu_moi As String)
Dim do_dai As Integer
      Me.MousePointer = 11
      do_dai = Len(tmpSoHieu)
      If cap_ct = 2 Then
            ExecuteSQL5 "UPDATE LoaiTaiSan SET SoHieu = '" + sohieu_moi _
                  + "' + Right(SoHieu,Len(SoHieu) - " + CStr(do_dai) + ") WHERE CapTren = " + CStr(maso_ct)
            ExecuteSQL5 "UPDATE TaiSan SET SoHieu = '" + sohieu_moi _
                  + "' + Right(SoHieu,Len(SoHieu) - " + CStr(do_dai) + ") WHERE MaLoai = " + CStr(maso_ct)
      Else
            ExecuteSQL5 "UPDATE TaiSan SET SoHieu = '" + sohieu_moi _
                  + "' + Right(SoHieu,Len(SoHieu) - " + CStr(do_dai) + ") WHERE MaNhom = " + CStr(maso_ct)
      End If
      tmpSoHieu = sohieu_moi
      Me.MousePointer = 1
End Sub
'======================================================================================
' SUB LayDanhSachPhanLoai
'======================================================================================
Private Sub LayDanhSachPhanLoai()
Dim rs_danhsach As Recordset
Dim chi_so As Integer
      Me.MousePointer = 11
      If Outline.ListCount > 0 Then Outline.Clear
      sql = "SELECT * FROM " + TenBang + " WHERE Cap > 0 ORDER BY SoHieu"
      Set rs_danhsach = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
      Do Until rs_danhsach.EOF
            Outline.AddItem rs_danhsach!sohieu + Chr(9) + rs_danhsach!Ten
            On Error Resume Next
            Outline.indent(chi_so) = rs_danhsach!cap
            On Error GoTo 0
            Outline.ItemData(chi_so) = rs_danhsach!MaSo
            chi_so = chi_so + 1
            rs_danhsach.MoveNext
      Loop
      rs_danhsach.Close
      Set rs_danhsach = Nothing
      For chi_so = 0 To (Outline.ListCount - 1)
            If Outline.HasSubItems(chi_so) Then Outline.Expand(chi_so) = True
      Next
      Me.MousePointer = 0
End Sub
'======================================================================================
' SUB KhoiTao
'======================================================================================
Private Sub KhoiTao(tiep_tuc As Boolean)
      PhanLoai.MaSo = 0
      PhanLoai.Ten = ""
      PhanLoai.sohieu = ""
      PhanLoai.CapTren = 0
      PhanLoai.cap = 0
      Text(0).Text = ""
      If tiep_tuc = True Then
            RFocus Text(1)
            Text(1).Text = tmpSoHieu
            SendKeys "{END}"
      Else
            Text(1).Text = ""
      End If
End Sub

