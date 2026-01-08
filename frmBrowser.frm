VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmBrowser 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Xem HD"
   ClientHeight    =   9570
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   12630
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9495
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      ExtentX         =   21193
      ExtentY         =   16748
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim plct As Integer
Private Sub Form_Load()
    Dim namct As Integer
    namct = SelectSQL("SELECT NamTC as F1 FROM License ")
    Dim mypath As String
    Dim folderName As String
    folderName = "HD" & namct
    mypath = App.path & "\Hoadon\" & folderName
    Dim LoaiHD As String
    If FrmChungtu.txtPhanloaichungtu.Text = 1 Or FrmChungtu.txtPhanloaichungtu.Text = 0 Then
        LoaiHD = "\HDVao"
    Else
        LoaiHD = "\HDRa"
    End If
    Dim kyhhd As String
    kyhhd = FrmChungtu.txtVT(1).Text
    If Left(kyhhd, 1) = "1" Then
        kyhhd = Mid(kyhhd, 2)    ' B? di ký t? d?u tiên
    End If

    Dim mst As String
    Dim sohd As String
    sohd = FrmChungtu.txt(0).Text
    Do While Left(sohd, 1) = "0" And sohd <> ""
        sohd = Mid(sohd, 2)
    Loop

    'mypath = mypath & LoaiHD & "\" & month(CDate(FrmChungtu.CboThang.Text)) & "\" & sohd & "_" & kyhhd & ".html"
    If LoaiHD = "\HDVao" Then
        mst = FrmChungtu.txtVT(9).Text
    Else
        mst = SelectSQL("select MaSoThue AS f1 from  License")
    End If
    mypath = mypath & LoaiHD & "\" & month(CDate(FrmChungtu.CboThang.Text)) & "\" & mst & "_" & sohd & "_" & kyhhd & ".html"

    'MsgBox FrmChungtu.


    ' Ki?m tra file có t?n t?i không
    If Dir(mypath) = "" Then

        Dim txtPath As String
        Dim FileNum As Integer

        txtPath = App.path & "\Hoadon\hdlink.txt"
        FileNum = FreeFile    ' L?y s? file còn tr?ng

        ' M? t?p d? ghi dè n?i dung
        Open txtPath For Output As #FileNum


        Print #FileNum, FrmChungtu.txtVT(9).Text & "_" & FrmChungtu.txtPhanloaichungtu & "_" & sohd & "_" & kyhhd & "_" & month(CDate(FrmChungtu.CboThang.Text))
        Close #FileNum
        Dim exePath As String
        exePath = App.path & "\\Tools\\Debug\\SaovietTax.exe"

        ' Shell d? m? ?ng d?ng
        Shell exePath, vbMinimizedNoFocus

    End If
    While Dir(mypath) = ""
        ' Có th? thêm m?t th?i gian ch? nh? d? tránh vi?c l?p quá nhanh
        DoEvents
    Wend
    Dim FilePath As String
    FilePath = mypath
    WebBrowser1.Navigate FilePath
End Sub
Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    SetZoomTo50Percent
End Sub

Private Sub SetZoomTo50Percent()
    On Error Resume Next

    ' Cách 1: Dùng ExecWB (n?u h? tr?)
    Const OLECMDID_OPTICAL_ZOOM = 63
    Const OLECMDEXECOPT_DONTPROMPTUSER = 2

    WebBrowser1.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 50

    ' Cách 2: Dùng CSS zoom (fallback)
    If Err.number <> 0 Then
        Err.Clear
        Dim doc As Object
        Set doc = WebBrowser1.Document
        doc.body.Style.Zoom = "0.9"
        ' Ho?c
        doc.body.Style.Transform = "scale(0.9)"
        doc.body.Style.transformOrigin = "0 0"
    End If
End Sub

