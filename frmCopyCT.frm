VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSCCT 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   LinkTopic       =   "Form4"
   ScaleHeight     =   825
   ScaleWidth      =   5565
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
         Left            =   5040
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
      Left            =   4320
      TabIndex        =   4
      Top             =   350
      Width           =   1095
   End
   Begin VB.ComboBox CboThang 
      Height          =   315
      Index           =   1
      ItemData        =   "frmCopyCT.frx":BDDA
      Left            =   3120
      List            =   "frmCopyCT.frx":BDE1
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1095
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
      Width           =   1095
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
      Left            =   2280
      TabIndex        =   2
      Tag             =   "to"
      Top             =   405
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
Public Function ThayThang(ByVal sChuoi As String, ByVal iThangMoi As Integer) As String
    Dim s As String
    Dim soMoi As String
    Dim pos As Long
    Dim i As Long
    
    s = Trim(sChuoi)
    If s = "" Then
        ThayThang = ""
        Exit Function
    End If
    
    soMoi = CStr(iThangMoi)
    
    ' === Uu tiên 1: D?ng tháng 1, thang 1, Tháng 1... ===
    Dim arrThang As Variant
    arrThang = Array("tháng", "Tháng", "THÁNG", "thang", "Thang", "THANG")
    
    Dim m As Integer
    For m = LBound(arrThang) To UBound(arrThang)
        pos = InStr(1, s, arrThang(m), vbTextCompare)
        If pos > 0 Then
            Dim startNum As Long
            startNum = pos + Len(arrThang(m))
            
            Do While startNum <= Len(s) And Mid(s, startNum, 1) = " "
                startNum = startNum + 1
            Loop
            
            Dim n As Long
            n = startNum
            Do While n <= Len(s) And Mid(s, n, 1) Like "[0-9]"
                n = n + 1
            Loop
            
            If n > startNum Then
                s = Left(s, startNum - 1) & soMoi & Mid(s, n)
                ThayThang = s
                Exit Function
            End If
        End If
    Next m
    
    ' === Uu tiên 2: D?ng tháng/ngày (VD: 12/25, 1/15) ===
    ' Tìm d?u / và xác d?nh s? nào là tháng
    pos = InStr(1, s, "/", vbTextCompare)
    If pos > 0 Then
        ' Tìm s? tru?c d?u /
        Dim beforeSlash As String
        Dim beforeStart As Long
        Dim beforeEnd As Long
        
        ' Tìm v? trí b?t d?u s? tru?c /
        beforeEnd = pos - 1
        beforeStart = beforeEnd
        Do While beforeStart > 1 And Mid(s, beforeStart - 1, 1) Like "[0-9]"
            beforeStart = beforeStart - 1
        Loop
        
        ' Ki?m tra n?u có s? tru?c d?u /
        If beforeStart <= beforeEnd And Mid(s, beforeStart, 1) Like "[0-9]" Then
            ' L?y s? tru?c /
            Dim beforeNumber As String
            beforeNumber = Mid(s, beforeStart, beforeEnd - beforeStart + 1)
            
            ' Ki?m tra n?u s? dó n?m trong kho?ng 1-12 (có th? là tháng)
            Dim beforeValue As Integer
            beforeValue = val(beforeNumber)
            
            ' Ki?m tra n?u s? tru?c / là tháng h?p lý (1-12) HO?C n?u s? sau / l?n hon 12
            ' thì uu tiên thay s? tru?c /
            Dim afterNumber As String
            Dim afterStart As Long
            Dim afterEnd As Long
            
            afterStart = pos + 1
            afterEnd = afterStart
            Do While afterEnd <= Len(s) And Mid(s, afterEnd, 1) Like "[0-9]"
                afterEnd = afterEnd + 1
            Loop
            afterEnd = afterEnd - 1
            
            If afterStart <= afterEnd Then
                afterNumber = Mid(s, afterStart, afterEnd - afterStart + 1)
                Dim afterValue As Integer
                afterValue = val(afterNumber)
                
                ' Logic thông minh:
                ' - N?u s? tru?c / t? 1-12 (có th? là tháng) và s? sau / > 12 (ch?c ch?n là ngày) -> thay s? tru?c
                ' - N?u s? tru?c / > 12 (ch?c ch?n là ngày) và s? sau / t? 1-12 (có th? là tháng) -> thay s? sau
                ' - N?u c? hai d?u <= 12 -> uu tiên thay s? tru?c / (coi là tháng/ngày)
                ' - N?u c? hai d?u > 12 -> thay s? sau (d?ng mã/ngày)
                
                If beforeValue >= 1 And beforeValue <= 12 And afterValue > 12 Then
                    ' Tru?ng h?p: tháng/ngày (VD: 12/25)
                    s = Left(s, beforeStart - 1) & soMoi & Mid(s, beforeEnd + 1)
                    ThayThang = s
                    Exit Function
                ElseIf beforeValue > 12 And afterValue >= 1 And afterValue <= 12 Then
                    ' Tru?ng h?p: mã/tháng (VD: cl25/1)
                    s = Left(s, afterStart - 1) & soMoi & Mid(s, afterEnd + 1)
                    ThayThang = s
                    Exit Function
                ElseIf beforeValue >= 1 And beforeValue <= 12 And afterValue >= 1 And afterValue <= 12 Then
                    ' C? hai d?u <= 12, uu tiên thay s? tru?c (coi là tháng/ngày)
                    s = Left(s, beforeStart - 1) & soMoi & Mid(s, beforeEnd + 1)
                    ThayThang = s
                    Exit Function
                Else
                    ' C? hai d?u > 12, thay s? sau (d?ng mã/ngày)
                    s = Left(s, afterStart - 1) & soMoi & Mid(s, afterEnd + 1)
                    ThayThang = s
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' === Uu tiên 3: D?ng T3, T12, t3 ===
    pos = InStr(1, s, "T", vbTextCompare)
    If pos > 0 Then
        For i = pos + 1 To Len(s)
            If Mid(s, i, 1) Like "[0-9]" Then
                Dim startT As Long
                startT = i
                Do While i <= Len(s) And Mid(s, i, 1) Like "[0-9]"
                    i = i + 1
                Loop
                s = Left(s, startT - 1) & soMoi & Mid(s, i)
                ThayThang = s
                Exit Function
            End If
        Next i
    End If
    
    ' === Uu tiên 4: D?ng LVP1, ABC2, XYZ12 (ch? + s?) ===
    For i = Len(s) To 1 Step -1
        If Mid(s, i, 1) Like "[0-9]" Then
            Dim startNumPos As Long
            startNumPos = i
            Do While startNumPos > 1 And Mid(s, startNumPos - 1, 1) Like "[0-9]"
                startNumPos = startNumPos - 1
            Loop
            
            If startNumPos > 1 Then
                Dim prevChar As String
                prevChar = Mid(s, startNumPos - 1, 1)
                If prevChar Like "[A-Za-z]" Then
                    s = Left(s, startNumPos - 1) & soMoi & Mid(s, i + 1)
                    ThayThang = s
                    Exit Function
                End If
            End If
            Exit For
        End If
    Next i
    
    ' Không tìm th?y m?u nào ? tr? v? nguyên chu?i
    ThayThang = sChuoi
    
End Function
Public Function ReplaceMonthNumber(ByVal inputString As String, ByVal newMonth As Integer) As String
    Dim result As String
    Dim i As Integer
    Dim tempNumber As String
    Dim startPos As Integer
    Dim j As Integer
    Dim found As Boolean

    result = inputString
    found = False

    i = 1
    Do While i <= Len(result)
        ' Ki?m tra n?u ký t? là s?
        If Mid(result, i, 1) >= "0" And Mid(result, i, 1) <= "9" Then
            ' L?y toàn b? s? b?t d?u t? v? trí i
            tempNumber = ""
            startPos = i
            j = i

            Do While j <= Len(result)
                If Mid(result, j, 1) >= "0" And Mid(result, j, 1) <= "9" Then
                    tempNumber = tempNumber & Mid(result, j, 1)
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop

            ' Ki?m tra t? phía tru?c s?
            If startPos > 1 Then
                ' Tru?ng h?p 1: "T" + s? (ch? thay n?u T d?ng m?t mình và sau T là s?)
                If Mid(result, startPos - 1, 1) = "T" Then
                    ' Ki?m tra n?u ký t? tru?c T không ph?i là ch? cái (ho?c là d?u chu?i)
                    If startPos - 1 = 1 Then
                        ' T ? d?u chu?i
                        result = "T" & CStr(newMonth) & Mid(result, j)
                        found = True
                        Exit Do
                    Else
                        Dim prevChar As String
                        prevChar = Mid(result, startPos - 2, 1)
                        ' N?u ký t? tru?c T không ph?i ch? cái (a-z, A-Z)
                        If Not ((prevChar >= "A" And prevChar <= "Z") Or (prevChar >= "a" And prevChar <= "z")) Then
                            result = Left(result, startPos - 2) & "T" & CStr(newMonth) & Mid(result, j)
                            found = True
                            Exit Do
                        End If
                    End If
                End If

                ' Tru?ng h?p 2: "tháng" + s? (có ho?c không có kho?ng tr?ng)
                Dim checkLen As Integer
                Dim checkText As String

                ' Ki?m tra "tháng " (có kho?ng tr?ng)
                If startPos >= 6 Then
                    checkText = LCase(Mid(result, startPos - 5, 6))
                    If checkText = "tháng " Then
                        result = Left(result, startPos - 5) & "tháng " & CStr(newMonth) & Mid(result, j)
                        found = True
                        Exit Do
                    End If
                End If

                ' Ki?m tra "tháng" (không kho?ng tr?ng)
                If startPos >= 5 Then
                    checkText = LCase(Mid(result, startPos - 5, 5))
                    If checkText = "tháng" Then
                        result = Left(result, startPos - 5) & "tháng" & CStr(newMonth) & Mid(result, j)
                        found = True
                        Exit Do
                    End If
                End If

                ' Ki?m tra "thang " (có kho?ng tr?ng)
                If startPos >= 6 Then
                    checkText = LCase(Mid(result, startPos - 5, 6))
                    If checkText = "thang " Then
                        result = Left(result, startPos - 5) & "thang " & CStr(newMonth) & Mid(result, j)
                        found = True
                        Exit Do
                    End If
                End If

                ' Ki?m tra "thang" (không kho?ng tr?ng)
                If startPos >= 5 Then
                    checkText = LCase(Mid(result, startPos - 5, 5))
                    If checkText = "thang" Then
                        result = Left(result, startPos - 5) & "thang" & CStr(newMonth) & Mid(result, j)
                        found = True
                        Exit Do
                    End If
                End If
            End If

            i = j
        Else
            i = i + 1
        End If
    Loop

    ReplaceMonthNumber = result
End Function
Private Sub Command1_Click()

    Tuthang = val(Replace(LCase$(CboThang(0).Text), "tháng", ""))
    denthang = val(Replace(LCase$(CboThang(1).Text), "tháng", ""))
    'Kiem tra xem loai ngay cua chung tu la cuoi thang ko
    Dim Query64 As String
    Query64 = "SELECT * FROM ChungTu WHERE MaCT=" & MaCT & ""
    Set rs_ktra164 = DBKetoan.OpenRecordset(Query64, dbOpenSnapshot)
    If Not rs_ktra164.EOF Then
        'Neu la tong hop
        If rs_ktra164!maloai = 0 Or rs_ktra164!maloai = 4 Or rs_ktra164!maloai = 3 Then
            Dim lastDay As Date
            lastDay = DateSerial(Year(rs_ktra164!NgayCT), month(rs_ktra164!NgayCT) + 1, 0)
            Dim GetDate As Date
            Dim getsh As String
            Dim getDG As String
            If lastDay = rs_ktra164!NgayCT Then
                GetDate = DateSerial(Year(rs_ktra164!NgayCT), Tuthang, 0)
            Else
                GetDate = DateSerial(Year(rs_ktra164!NgayCT), Tuthang - 1, Day(rs_ktra164!NgayCT))
            End If
            FrmChungtu.bakNgay = GetDate
            'Tang so hieu hien tai5 len den gatedate
            Dim currentThang As Integer
            currentThang = month(rs_ktra164!NgayCT)
            Dim i As Integer
            getsh = rs_ktra164!sohieu
            TextBox1.Text = VniToUnicode(rs_ktra164!diengiai)
            getDG = TextBox1.Text

            FrmChungtu.bakSoHieu = ThayThang(getsh, Tuthang)
            FrmChungtu.bakNoidung = ThayThang(getDG, Tuthang)
        End If
    End If
    If Not getsh Like "*[0-9]*" Then
        getsh = getsh & "_" & Tuthang
        FrmChungtu.bakSoHieu = getsh
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
Private Sub lblTitle_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
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

