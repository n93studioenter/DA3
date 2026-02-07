VERSION 5.00
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmTaiLieu 
   BorderStyle     =   0  'None
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9555
   Icon            =   "frmTaiLieu.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFakeTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   13575
      TabIndex        =   2
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
         Left            =   9000
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   0
         Width           =   4455
      End
      Begin VB.Image picIcon 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   120
         Picture         =   "frmTaiLieu.frx":57E2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   8550
         Index           =   0
         Left            =   840
         Picture         =   "frmTaiLieu.frx":5A9F
         Stretch         =   -1  'True
         Top             =   240
         Width           =   7890
      End
   End
   Begin MSGrid.Grid Grid 
      Height          =   4365
      Left            =   9900
      TabIndex        =   0
      Top             =   180
      Width           =   600
      _Version        =   65536
      _ExtentX        =   1058
      _ExtentY        =   7699
      _StockProps     =   77
      BackColor       =   16777215
      FixedRows       =   0
      HighLight       =   0   'False
   End
   Begin MSOutl.Outline Outline 
      Height          =   8775
      Left            =   -120
      TabIndex        =   1
      Top             =   -120
      Width           =   9495
      _Version        =   65536
      _ExtentX        =   16748
      _ExtentY        =   15478
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
      MousePointer    =   1
      MouseIcon       =   "frmTaiLieu.frx":115BC
      PicturePlus     =   "frmTaiLieu.frx":115D8
      PictureMinus    =   "frmTaiLieu.frx":116D2
      PictureLeaf     =   "frmTaiLieu.frx":117CC
      PictureOpen     =   "frmTaiLieu.frx":118C6
      PictureClosed   =   "frmTaiLieu.frx":119C0
   End
End
Attribute VB_Name = "frmTaiLieu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mang(100) As String


Sub LayFile(ByVal ThuMuc As String, ByVal i As Integer)
    Dim f As String
    Dim thoat
    thoat = 0


    If Right(ThuMuc, 1) <> "\" Then
        ThuMuc = ThuMuc & "\"
    ElseIf Right(ThuMuc, 1) = "\" Then

    End If

    f = Dir$(ThuMuc & "*.*")
    'List1.Clear
    While Len(f)
        Outline.AddItem f
        Outline.indent(i) = 2
        Outline.ItemData(i) = 0

        ' Grid.AddItem Str(i) + Chr(9) + f, 0
        ' List1.AddItem F
        f = Dir$
        i = i + 1
    Wend

End Sub
Public Sub ColumnSetUp(Grid_control As Grid, col_index As Integer, col_Width As Integer, col_alignment As Integer)
      Grid_control.Row = 0
      Grid_control.col = col_index
      Grid_control.ColWidth(col_index) = col_Width
      Grid_control.FixedAlignment(col_index) = col_alignment
      If col_index >= Grid_control.FixedCols Then Grid_control.ColAlignment(col_index) = col_alignment
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
Public Sub AnControl(frm As Form)
    Dim ctl As Control
Dim TITLE_HEIGHT As Long
TITLE_HEIGHT = 450
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
    Caption = "Tµi liÖu l­u tr÷"
    lblTitle(11).Caption = Caption
    lblTitle(11).AutoSize = True
    Me.Height = Me.Height + 390 + 10
    picFakeTitle.Width = Me.ScaleWidth
    picFakeTitle.Height = 325
    picIcon(1).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2
    lblTitle(11).Left = picIcon(1).Left + picIcon(1).Width + 90
    lblTitle(11).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2 + 15
    lblClose.Top = 55
    AnControl Me

    Dim fso As New FileSystemObject
    Dim fil As folder    ' File
    Dim fil1 As file
    Dim file
    Dim j As Integer
    j = 0
    Dim ThuMuc As String
    ThuMuc = pCurDir + "tailieu\"
    'If Right(ThuMuc, 1) <> "\" Then
    'ThuMuc = ThuMuc & "\"
    'ElseIf Right(ThuMuc, 1) = "\" Then
    'End If

    f = Dir$(ThuMuc & "*.*")
    While Len(f)
        Outline.AddItem f
        Outline.indent(j) = 1
        Outline.ItemData(j) = 0
        f = Dir$
        mang(j) = ThuMuc + f
        j = j + 1
    Wend

    For Each fil In fso.GetFolder(pCurDir + "tailieu\").SubFolders
        Outline.AddItem fil.Name
        Outline.indent(j) = 1
        Outline.ItemData(j) = 0
        mang(j) = fil
        For Each fil1 In fso.GetFolder(fil).Files
            j = j + 1
            Outline.AddItem fil1.Name
            Outline.indent(j) = 2
            Outline.ItemData(j) = 0
            mang(j) = fil1
        Next
        j = j + 1

    Next


End Sub

Private Sub Grid_Click()
'Shell "EXPLORER.EXE " & "D:\hinhanh\bill\" + Grid.Text
End Sub



Private Sub Outline_Click()
'MsgBox mang(Outline.ListIndex)
Shell "Explorer.exe " & mang(Outline.ListIndex), vbMaximizedFocus

End Sub
