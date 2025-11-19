VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmBrowser 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Xem HD"
   ClientHeight    =   7305
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   13440
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      ExtentX         =   23098
      ExtentY         =   11456
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

    Dim mypath As String
    mypath = App.path & "\Hoadon"
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

    Dim sohd As String
    sohd = FrmChungtu.txt(0).Text
    Do While Left(sohd, 1) = "0" And sohd <> ""
        sohd = Mid(sohd, 2)
    Loop

    mypath = mypath & LoaiHD & "\" & month(CDate(FrmChungtu.CboThang.Text)) & "\" & sohd & "_" & kyhhd & ".html"
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
