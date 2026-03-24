VERSION 5.00
Begin VB.Form FrmMatkhau 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   2385
   ClientLeft      =   4620
   ClientTop       =   4875
   ClientWidth     =   4230
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
   Icon            =   "Frmmatkh.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Security Check"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Frmmatkh.frx":57E2
   ScaleHeight     =   2385
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.PictureBox picTitle 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   4170
      TabIndex        =   6
      Top             =   0
      Width           =   4230
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
         Height          =   400
         Left            =   3720
         TabIndex        =   8
         Top             =   75
         Width           =   600
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   250
         Left            =   100
         Picture         =   "Frmmatkh.frx":62CC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   250
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   500
         TabIndex        =   7
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   1440
      Picture         =   "Frmmatkh.frx":6589
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Return"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2880
      Picture         =   "Frmmatkh.frx":79AB
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "&Ok"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox CboUser 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   765
      Width           =   2775
   End
   Begin VB.TextBox txtPsw 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1260
      Width           =   2775
   End
   Begin VB.Label Label 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nh©n viªn"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Tag             =   "User Name"
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MËt khÈu "
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Tag             =   "Password"
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "FrmMatkhau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 



Private Type TrackMouseEvent
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Declare Function TrackMouseEvent Lib "user32" _
                                         (lpEventTrack As TrackMouseEvent) As Long

Private Const TME_LEAVE = &H2

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                                     (ByVal hwnd As Long, ByVal wMsg As Long, _
                                      ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2


Private Declare Function SetWindowTextW Lib "user32" _
                                        (ByVal hwnd As Long, ByVal lpString As Long) As Long
Private Declare Function MultiByteToWideChar Lib "Kernel32" _
                                             (ByVal CodePage As Long, ByVal dwFlags As Long, _
                                              lpMultiByteStr As Any, ByVal cchMultiByte As Long, _
                                              ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Private Declare Function GetAdaptersInfo Lib "iphlpapi" (lpAdapterInfo As Any, lpSize As Long) As Long

Dim Counter As Integer
Dim pass As Integer
Dim psw As String
Dim ok As Boolean
Dim scecretpws As String
Private m_Title As String   ' gi? chu?i s?ng


'====================================================================================================
' KiÓm tra mËt khÈu
'====================================================================================================

Public Function GetMacAddress() As String
    Const OFFSET_LENGTH As Long = 400
    Dim lSize As Long
    Dim baBuffer() As Byte
    Dim lIdx As Long
    Dim sRetVal As String

    Call GetAdaptersInfo(ByVal 0, lSize)
    If lSize <> 0 Then
        ReDim baBuffer(0 To lSize - 1) As Byte
        Call GetAdaptersInfo(baBuffer(0), lSize)
        Call CopyMemory(lSize, baBuffer(OFFSET_LENGTH), 4)
        For lIdx = OFFSET_LENGTH + 4 To OFFSET_LENGTH + 4 + lSize - 1
            sRetVal = IIf(LenB(sRetVal) <> 0, sRetVal & ":", vbNullString) & Right$("0" & Hex$(baBuffer(lIdx)), 2)
        Next
    End If
    GetMacAddress = sRetVal
End Function
Public Sub CheckAndCreateTableThongTinToKhai()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbThongTinToKhai"    ' Thay d?i tên b?ng c?a b?n ? dây
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        Set fld = tdf.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField    ' Thi?t l?p thu?c tính t? d?ng tang
        tdf.Fields.Append fld

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("Quy1", dbDouble)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Quy2", dbDouble)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Quy3", dbDouble)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Quy4", dbDouble)
        tdf.Fields.Append fld
        Dim i As Integer

        For i = 1 To 12
            Set fld = tdf.CreateField("T" & i, dbDouble)    ' dbInteger cho ki?u s?
            ' Ho?c có th? dùng dbDouble, dbSingle, dbCurrency tùy nhu c?u
            tdf.Fields.Append fld
        Next i
        Set fld = tdf.CreateField("NguoiKy", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Nam", dbInteger, 255)
        tdf.Fields.Append fld
        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf

    End If
End Sub
Public Sub CheckAndCreateTableDinhDanh()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbDinhdanhtaikhoan"    ' Thay d?i tên b?ng c?a b?n ? dây
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        Set fld = tdf.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField    ' Thi?t l?p thu?c tính t? d?ng tang
        tdf.Fields.Append fld

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("Type", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("KeyValue", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TKNo", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TKCo", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TKThue", dbText, 255)
        tdf.Fields.Append fld
        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf

    End If
End Sub
Public Sub CheckAndCreateTableFontSize()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String
    tableName = "tbFontsize"    ' Thay d?i tên b?ng c?a b?n ? dây
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        Set fld = tdf.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField    ' Thi?t l?p thu?c tính t? d?ng tang
        tdf.Fields.Append fld

        ' Menu
        Set fld = tdf.CreateField("MemnuFont", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("MenuSize", dbDouble)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("MenuBold", dbDouble)
        tdf.Fields.Append fld
        'Font control
        Set fld = tdf.CreateField("ControlvniFont", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("ControlvniSize", dbDouble)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("ControluniFont", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("ControluniSize", dbDouble)
        tdf.Fields.Append fld
        DBKetoan.TableDefs.Append tdf

    End If
End Sub
Public Sub CheckAndCreateTablePL1()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbPL1"    ' Thay d?i tên b?ng c?a b?n ? dây
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        Set fld = tdf.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField    ' Thi?t l?p thu?c tính t? d?ng tang
        tdf.Fields.Append fld

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("TenHH", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("GT1", dbDouble)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("GT2", dbDouble)
        tdf.Fields.Append fld
        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf

    End If
End Sub
Public Sub CheckAndCreateTablePL2()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbPL2"    ' Thay d?i tên b?ng c?a b?n ? dây
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        Set fld = tdf.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField    ' Thi?t l?p thu?c tính t? d?ng tang
        tdf.Fields.Append fld

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("TenHH", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("GT1", dbDouble)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("GT2", dbDouble)
        tdf.Fields.Append fld
        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf

    End If
End Sub
Public Sub CheckAndCreateTableImport()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbimport"    ' Thay d?i tên b?ng c?a b?n ? dây
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        Set fld = tdf.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField    ' Thi?t l?p thu?c tính t? d?ng tang
        tdf.Fields.Append fld
        ' Thi?t l?p tru?ng ID là khóa chính
         

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("SHDon", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("KHHDon", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("NLap", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Ten", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Noidung", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TKCo", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TKNo", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TkThue", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Mst", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Status", dbDouble)  ' Ho?c dbInteger n?u b?n mu?n ki?u s? nguyên
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Ngaytao", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TongTien", dbDouble)  ' Ho?c dbInteger n?u b?n mu?n ki?u s? nguyên
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Vat", dbDouble)  ' Ho?c dbInteger n?u b?n mu?n ki?u s? nguyên
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("SohieuTP", dbText, 255)
        tdf.Fields.Append fld
        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf

    End If
End Sub
Public Sub CheckAndCreateTableImportDetail()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbimportdetail"    ' Thay d?i tên b?ng c?a b?n ? dây
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        Set fld = tdf.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField    ' Thi?t l?p thu?c tính t? d?ng tang
        tdf.Fields.Append fld

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("ParentId", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("SoHieu", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("SoLuong", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("DonGia", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("DVT", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Ten", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("MaCT", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TKNo", dbText, 255)
        tdf.Fields.Append fld
        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf

    End If
End Sub
Public Sub CreateLicense()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbLicensekey"
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf


    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("Type", dbText, 255)
        tdf.Fields.Append fld
        ' T?o tru?ng hoadonpath
        Set fld = tdf.CreateField("Year", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Totals", dbText, 255)
        tdf.Fields.Append fld

        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf

    End If
End Sub
Public Sub CheckAndCreateToKhaiThueTable()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "ToKhaiThue"    ' Tên b?ng
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        ' T?o các tru?ng s? t? N11 d?n N43
        Dim i As Integer
        For i = 11 To 43
            Set fld = tdf.CreateField("N" & i, dbDouble)
            tdf.Fields.Append fld
        Next i

        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf
        
        Dim sqlInsert As String
        ' Chèn m?t b?n ghi m?i v?i t?t c? các tru?ng = 0
        sqlInsert = "INSERT INTO ToKhaiThue ("
        For i = 11 To 43
            sqlInsert = sqlInsert & "N" & i & ", "
        Next i
        sqlInsert = Left(sqlInsert, Len(sqlInsert) - 2) ' Lo?i b? d?u ph?y cu?i
        sqlInsert = sqlInsert & ") VALUES ("
        For i = 11 To 43
            sqlInsert = sqlInsert & "0, "
        Next i
        sqlInsert = Left(sqlInsert, Len(sqlInsert) - 2) ' Lo?i b? d?u ph?y cu?i
        sqlInsert = sqlInsert & ");"
        
        DBKetoan.Execute sqlInsert
    Else
        'MsgBox "B?ng ToKhaiThue dã t?n t?i!"
    End If
End Sub
Public Sub CheckAndCreateTBLogin()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbLogin"
    tableExists = False


    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        ' Username
        Set fld = tdf.CreateField("Username", dbText, 255)
        fld.Required = False
        fld.AllowZeroLength = True
        tdf.Fields.Append fld

        ' Password
        Set fld = tdf.CreateField("Password", dbText, 255)
        fld.Required = False
        fld.AllowZeroLength = True   ' ?? QUAN TR?NG (fix l?i c?a b?n)
        tdf.Fields.Append fld

        ' Thêm b?ng vào CSDL
        DBKetoan.TableDefs.Append tdf
    End If
End Sub
Public Sub CheckAndCreateTable()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbRegister"    ' Thay d?i tên b?ng c?a b?n ? dây
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("Name", dbText, 255)
        tdf.Fields.Append fld
        ' T?o tru?ng hoadonpath
        Set fld = tdf.CreateField("Hoadonpath", dbText, 255)
        tdf.Fields.Append fld

        ' T?o tru?ng dbpath
        Set fld = tdf.CreateField("Dbpath", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Username", dbText, 255)
        tdf.Fields.Append fld
         Set fld = tdf.CreateField("Password", dbText, 255)
        tdf.Fields.Append fld
        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf
        ' Chèn d?a ch? MAC vào dòng d?u tiên
        Dim mac As String
        mac = GetMacAddress()
        Dim sql As String

        sql = "INSERT INTO tbRegister(Name) VALUES ('" & mac & "');"
        DBKetoan.Execute sql
    End If
End Sub
Private Sub importRegister()

    Dim FilePath As String
    Dim fileNumber As Integer
    fileNumber = FreeFile    ' L?y s? file t? d?ng

    Dim pathHoadon As String
    pathHoadon = App.path & "\Hoadon"    ' S?a d?u "\" d? d?m b?o du?ng d?n dúng
    FilePath = App.path & "\Hoadon\dpPath.txt"

    ' M? file d? ghi (n?u file dã t?n t?i, nó s? b? ghi dè)
    Open FilePath For Output As #fileNumber

    ' Ghi n?i dung vào file
    Print #fileNumber, pDataPath

    ' Ðóng file
    Close #fileNumber
    Dim rs As Recordset
    Dim sql As String
    Dim hoadonPathValue As String
    hoadonPathValue = App.path & "\Hoadon"    ' Ðu?ng d?n m?i cho hoadonpath

    ' Truy v?n d? l?y b?n ghi
    sql = "SELECT * FROM tbRegister"    ' Gi? d?nh b?ng ch? có 1 dòng
    Set rs = DBKetoan.OpenRecordset(sql)

    If Not rs.EOF Then
        ' C?p nh?t giá tr? cho hoadonpath
        rs.Edit
        rs!Hoadonpath = hoadonPathValue
        rs!Dbpath = pDataPath    ' C?p nh?t giá tr? cho pathDB

        rs.Update
    Else
        MsgBox "Không tìm th?y b?n ghi."
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Command_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If

    'Lay dia chi mac
    Select Case FrmMatkhau.tag
    Case 0:
        'Kiem tra tblogin da co1 data chua
        Dim mac As String
        mac = GetMacAddress()

        Dim rs_countlogin As Recordset
        Dim rsCount As DAO.Recordset

        Set rsCount = DBKetoan.OpenRecordset( _
                      "SELECT COUNT(*) AS Tong FROM tbLogin ", dbOpenSnapshot)
        If rsCount!tong = 0 Then
            'Neu chua thi kiem tra bt
            If KiemTraMatKhau(txtPsw.Text) Then
                HienThongBao VString(CboUser.Text), 3
                ok = True
                ExecuteSQL5 "UPDATE Users SET WS='" + GetComputerName1 + "' WHERE MaSo=" + CStr(UserID), False


                Dim sql As String

                sql = "update tbRegister SET Name= ('" & mac & "');"
                DBKetoan.Execute sql
                'Tao user moi tu user cu
                Dim pw As String
                pw = 64 + pNamTC + Day(Date)
                If Trim(txtPsw.Text) = pw Or txtPsw.Text = "" Then
                    ExecuteSQL5 "INSERT INTO tbLogin(Username, Password) VALUES('" & _
                                Replace(mac, "'", "''") & "','" & pNamTC & "')"
                Else
                    ExecuteSQL5 "INSERT INTO tbLogin(Username, Password) VALUES('" & _
                                Replace(mac, "'", "''") & "','" & Replace(txtPsw.Text, "'", "''") & "')"
                End If

                Unload Me
            Else

                'MsgBox "Sai mËt khÈu !", vbExclamation, App.ProductName
                Dim s As String
                s = ChrW(83) & ChrW(97) & ChrW(105) & ChrW(32) & ChrW(109) & ChrW(7853) & ChrW(116) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(7849) & ChrW(117)
                MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

                Counter = Counter + 1
                If Counter > 3 Then
                    Unload Me
                Else
                    RFocus txtPsw
                End If
            End If
        Else
            'Truong hop da co tk login
            'Kiem tra user dang nhap
            Dim rs As Recordset
            If txtPsw.Text = "" Then
                txtPsw.Text = pNamTC
            End If
            sql = "SELECT * FROM tbLogin WHERE Username='" & Replace(mac, "'", "''") & _
                  "' AND Password='" & Replace(Trim(txtPsw.Text), "'", "''") & "'"
            Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
            If Not rs.EOF Then
                HienThongBao VString(CboUser.Text), 3
                ok = True
                Unload Me
            Else
                'Neu dang nhap   that bai , kiem tra tiep co dang dung mat khau 64 khong

                pass = 64 + pNamTC + Day(Date)
                If Trim(txtPsw.Text) = pass Then
                    HienThongBao VString(CboUser.Text), 3
                    ok = True

                    'Kiem tra xem user nay moi hay cu
                    'Neu moi tao moi
                    Dim rskt As DAO.Recordset

                    sql = "SELECT * FROM tbLogin WHERE Username='" & Replace(mac, "'", "''") & "'"
                    Set rskt = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)

                    'Neu chua co tao moi user
                    If rskt.EOF Then
                        'neu count user trc do =2 thi xoa user dau tien
                        ' If rsCount!tong = 2 Then
                        'ExecuteSQL5 "DELETE FROM tbLogin WHERE Username = (SELECT TOP 1 Username FROM tbLogin ORDER BY Username)"
                        'End If
                        ExecuteSQL5 "INSERT INTO tbLogin(Username, Password) VALUES('" & _
                                    Replace(mac, "'", "''") & "','" & pNamTC & "')"
                    Else
                        'neu co toi thi cap nhat lai mk=''
                        ExecuteSQL5 "UPDATE tbLogin SET Password='' WHERE Username='" & Replace(mac, "'", "''") & "'"
                    End If
                    'Neu cu thi set mk=''
                    Unload Me
                Else
                    txtPsw.Text = ""
                    s = ChrW(83) & ChrW(97) & ChrW(105) & ChrW(32) & ChrW(109) & ChrW(7853) & ChrW(116) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(7849) & ChrW(117)
                    MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

                    Counter = Counter + 1
                    If Counter > 3 Then
                        Unload Me
                    Else
                        RFocus txtPsw
                    End If
                End If


            End If
            rs.Close
            Set rs = Nothing
        End If


    Case 1:
        Select Case pass
        Case 0:
            'Kiem tra mat khau hien tai
            mac = GetMacAddress()
            If txtPsw.Text = "" Then
                txtPsw.Text = pNamTC
            End If
            sql = "SELECT * FROM tbLogin WHERE Username='" & Replace(mac, "'", "''") & _
                  "' AND Password='" & Replace(Trim(txtPsw.Text), "'", "''") & "'"
            Set rskt = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
            If Not rskt.EOF Then
                pass = 1
                Label(0).Caption = "MËt khÈu míi"
                txtPsw.Text = ""
                RFocus txtPsw
            Else
                'MsgBox "Sai mËt khÈu !", vbExclamation, App.ProductName
                s = ChrW(83) & ChrW(97) & ChrW(105) & ChrW(32) & ChrW(109) & ChrW(7853) & ChrW(116) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(7849) & ChrW(117)
                MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly
                Unload FrmMatkhau
            End If
        Case 10:
            If KiemTraMatKhau(txtPsw.Text) Then
                pass = 1
                Label(0).Caption = "MËt khÈu míi"
                'txtPsw.Text = ""
                RFocus txtPsw
            Else
                'MsgBox "Sai mËt khÈu !", vbExclamation, App.ProductName
                s = ChrW(83) & ChrW(97) & ChrW(105) & ChrW(32) & ChrW(109) & ChrW(7853) & ChrW(116) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(7849) & ChrW(117)
                MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly
                Unload FrmMatkhau
            End If
        Case 1:
            psw = txtPsw.Text
            pass = 2
            txtPsw.Text = ""
            RFocus txtPsw
        Case 2:
            If txtPsw.Text = psw Then
                mac = GetMacAddress()
                If txtPsw.Text = "" Then
                    txtPsw.Text = pNamTC
                End If
                ExecuteSQL5 "UPDATE tbLogin SET Password='" & Replace(Trim(txtPsw.Text), "'", "''") & "' WHERE Username='" & Replace(mac, "'", "''") & "'"
                'ExecuteSQL5 "UPDATE Users SET Psw = " + CStr(Int_StrToCode(psw) + pNamTC) + " WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex))
                Unload FrmMatkhau
            Else
                MsgBox "B¹n ch­a nhí ®óng mËt khÈu !", vbExclamation, App.ProductName
                RFocus txtPsw
            End If
        End Select
    End Select

    'On Error Resume Next
    'Unload frmToast          ' Gi?i phóng instance cu n?u có
    'On Error GoTo 0
    'Dim Duration As Double
    'Duration = 200
    'Load frmToast


    ' C?u hình Duration n?u du?c ch? d?nh


    'frmToast.Show vbModeless   ' Hi?n th? không block form chính
    'frmToast.Refresh
    'frmToast.ShowToast "Thành công!", "Ðã d?ng b? hóa don d?u vào thành công.", 6000

End Sub

Private Sub AddDataLCTT()

'Chinh sua so tk cac cot
'ExecuteSQL5 "Update LCTT set TKCo= '7113' where MaSo=22"
'Xoa  Luu chuy?n ti?n t? ho?t d?ng d?u tu xong add lai
    ExecuteSQL5 "Delete from LCTT where MaSoCha=30"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (11,21,'1. TiÒn chi ®Ó mua s¾m, x©y dùng TSC§, B§S§T vµ c¸c tµi s¶n dµi h¹n kh¸c','211','11',0,0,0,0,30,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (12,22,'2. TiÒn thu tõ thanh lý, nh­îng b¸n TSC§, B§S§T vµ c¸c tµi s¶n dµi h¹n kh¸c','11','7113',0,0,0,0,30,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (13,23,'3. TiÒn chi cho vay, ®Çu t­ gãp vèn vµo ®¬n vÞ kh¸c','128','11',0,0,0,0,30,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (14,24,'4. TiÒn thu håi cho vay, ®Çu t­ gãp vèn vµo ®¬n vÞ kh¸c','11','128',0,0,0,0,30,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (15,25,'5. TiÒn thu l·i cho vay, cæ tøc vµ lîi nhuËn ®­îc chia','11','515',0,0,0,0,30,1,'new')"

    '07
    ExecuteSQL5 "Delete from LCTT where MaSo=6"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (7,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','711',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "Delete from LCTT where MaSo=7"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (37,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','811','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (38,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','138','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (39,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','3331','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (40,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','3332','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (41,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','3333','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (42,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','3335','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (43,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','3336','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (44,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','3337','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (45,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','3338','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (46,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','3339','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (47,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','338','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (48,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','352','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (49,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','353','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (50,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','356','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (51,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','131','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (52,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','515','11',0,0,0,0,20,-1,'new')"
    '06

    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (53,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','133',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (54,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','141',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (55,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','138',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (56,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','338',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (57,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','331',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (58,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','152',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (59,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','153',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (60,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','156',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (61,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','642',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (62,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','334',0,0,0,0,20,1,'new')"
    '02
    ExecuteSQL5 "Delete from LCTT where MaSo=2"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (63,2,'2. TiÒn chi tr¶ cho ng­êi cung cÊp hµng ho¸ vµ dÞch vô','331','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (64,2,'2. TiÒn chi tr¶ cho ng­êi cung cÊp hµng ho¸ vµ dÞch vô','152','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (65,2,'2. TiÒn chi tr¶ cho ng­êi cung cÊp hµng ho¸ vµ dÞch vô','153','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (66,2,'2. TiÒn chi tr¶ cho ng­êi cung cÊp hµng ho¸ vµ dÞch vô','154','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (67,2,'2. TiÒn chi tr¶ cho ng­êi cung cÊp hµng ho¸ vµ dÞch vô','156','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (68,2,'2. TiÒn chi tr¶ cho ng­êi cung cÊp hµng ho¸ vµ dÞch vô','642','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (69,2,'2. TiÒn chi tr¶ cho ng­êi cung cÊp hµng ho¸ vµ dÞch vô','242','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (70,2,'2. TiÒn chi tr¶ cho ng­êi cung cÊp hµng ho¸ vµ dÞch vô','1331','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (71,2,'2. TiÒn chi tr¶ cho ng­êi cung cÊp hµng ho¸ vµ dÞch vô','632','11',0,0,0,0,20,-1,'new')"

    'Chi tieu 40
    ExecuteSQL5 "Delete from LCTT where MaSoCha=40"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (72,31,'1. TiÒn thu tõ ph¸t hµnh cæ phiÕu, nhËn vèn gãp cña CSH','11','411',0,0,0,0,40,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (73,32,'2. TiÒn tr¶ l¹i vèn gãp cho c¸c chñ së h÷u, mua l¹i cæ phiÕu cña doanh nghiÖp ®· ph¸t hµnh','411','11',0,0,0,0,40,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (74,33,'3. TiÒn thu tõ ®i vay','11','341',0,0,0,0,40,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (75,34,'4. TiÒn tr¶ nî gèc vay vµ nî thuª tµi chÝnh','341','11',0,0,0,0,40,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (76,35,'5. Cæ tøc, lîi nhuËn ®· tr¶ cho chñ së h÷u','421','11',0,0,0,0,40,-1,'new')"

    'bo sung chi tieu 6
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (77,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','351',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (78,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','352',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (79,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','353',0,0,0,0,20,1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (80,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','356',0,0,0,0,20,1,'new')"
    'bo sung chi tieu 7
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (81,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','351','11',0,0,0,0,20,-1,'new')"
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (82,7,'7. TiÒn chi kh¸c cho ho¹t ®éng kinh doanh','141','11',0,0,0,0,20,-1,'new')"

    'bo sung hoan tien tk loai 333
    ExecuteSQL5 "INSERT INTO LCTT (Ma,MaSo,Ten,TKNo,TKCo,DauNam,KyTruoc,KyNay,TongHop,MaSoCha,Dau,TenE) VALUES (83,6,'6. TiÒn thu kh¸c tõ ho¹t ®éng kinh doanh','11','3330',0,0,0,0,20,1,'new')"

    ExecuteSQL5 "Update LCTT set Ten='TiÒn vµ t­¬ng ®­¬ng tiÒn cuèi kú (70 = 50+60+61)' where MaSo=70 "
    ExecuteSQL5 "Update LCTT set Ten='L­u chuyÓn tiÒn thuÇn trong kú (50 = 20+30+40)' where MaSo=50 "
End Sub
Private Sub Form_Activate()
    AddDataLCTT
    Left = frmMain.ScaleWidth * 30 / 100
    Top = frmMain.ScaleHeight * 40 / 100
    CheckAndCreateTBLogin
    CheckAndCreateTable
    CheckAndCreateToKhaiThueTable
    CreateLicense
    'Kiem tra neu chua co dong nao thi insert dong mac dinh
    Dim countrow As Integer

    countrow = SelectSQL("select count(*) AS f1 from  tbLicensekey")
    If countrow = 0 Then
        ExecuteSQL5 ("insert into tbLicensekey(Type,Year,Totals) values(0,0,0)")
    End If
    CheckAndCreateTableFontSize
    CheckAndCreateTableThongTinToKhai
    CheckAndCreateTablePL1
    CheckAndCreateTablePL2
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD k1 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD k2 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD k3 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD k4 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD k5 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD k6 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD k7 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD k8 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD k9 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD k10 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD k11 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD k12 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD xml1 MEMO")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD xml2 MEMO")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD xml3 MEMO")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD xml4 MEMO")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD xml5 MEMO")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD xml6 MEMO")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD xml7 MEMO")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD xml8 MEMO")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD xml9 MEMO")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD xml10 MEMO")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD xml11 MEMO")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbThongTinToKhai ADD xml12 MEMO")
    'CheckAndCreateTableDinhDanh
    'CheckAndCreateTableImport
    'CheckAndCreateTableImportDetail
    importRegister
    scecretpws = ""
    If Counter < 0 Then
        Counter = 0
        If Me.tag = 1 Then
            Dim i As Integer

            Me.Caption = "Thay ®æi mËt khÈu"
            Label(0).Caption = "MËt khÈu cò"
            SetListIndex CboUser, UserID
            ok = True
        Else
            ok = False
        End If
    End If
    Dim rs As DAO.Recordset
    Set rs = DBKetoan.OpenRecordset("SELECT TOP 1 Name FROM tbRegister ")
    If Not rs.EOF Then
        Dim mac As String
        mac = GetMacAddress()
        If rs!Name <> mac Then
            Dim newpsw As Integer
            newpsw = 64 + Day(Date) + pNamTC
            scecretpws = Int_StrToCode(CStr(newpsw))
            ExecuteSQL5 "UPDATE Users SET Psw = " + scecretpws + " WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex))
            'Dang xai o may khac
            'Cap nhat mat khau theo tohng so he thong

        End If
    End If

End Sub
'====================================================================================================
' Thu tuc kiem tra mat khau
'====================================================================================================
Private Function KiemTraMatKhau(pstr_psw As String) As Boolean

    Dim newpsw As Integer
    newpsw = 64 + Day(Date) + pNamTC
    If pstr_psw <> "" Then
        If pstr_psw = newpsw Then
            scecretpws = Int_StrToCode(CStr(newpsw))
            ExecuteSQL5 "UPDATE Users SET Psw = " + scecretpws + " WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex))
        End If
    End If

    Dim rs_mk As Recordset

    Set rs_mk = DBKetoan.OpenRecordset("SELECT Users.* FROM Users WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex)), dbOpenSnapshot, dbForwardOnly)
    If (Int_StrToCode(pstr_psw) = rs_mk!psw - pNamTC Or Int_StrToCode(pstr_psw) = rs_mk!psw) Then
        KiemTraMatKhau = True
        If Int_StrToCode(pstr_psw) = rs_mk!psw Then
            ExecuteSQL5 "UPDATE Users SET Psw =  '" & pNamTC & "' WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex))
        End If

    Else
        KiemTraMatKhau = False
        On Error GoTo SaiMK
        KiemTraMatKhau = (CInt5(pstr_psw) = Day(Date) + month(Date) + pNamTC)
        'Thu voi tbloign
        Dim mac As String
        mac = GetMacAddress()
        Dim rs As DAO.Recordset
        Dim sql As String

        sql = "SELECT * FROM tbLogin WHERE Username='" & Replace(mac, "'", "''") & _
              "' AND (Password='" & Replace(Trim(txtPsw.Text), "'", "''") & _
              "' OR Password IS NULL OR Password='')"

        Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)

        If Not rs.EOF Then
            KiemTraMatKhau = True
        Else
            KiemTraMatKhau = False
        End If

        rs.Close
        Set rs = Nothing
        On Error GoTo 0
    End If

    User_Right = rs_mk!UserRight
    UserID = rs_mk!MaSo
    UserName = rs_mk!TenNSD
    frmMain.tag = CStr(rs_mk!vt)
    frmMain.SetUserRight
    frmMain.sbStatusBar.Panels(3).ToolTipText = "Log On Time: " + Format(Time, "hh:mm:ss")
SaiMK:
    rs_mk.Close
    Set rs_mk = Nothing
End Function
Private Function KiemTraMatKhau2(pstr_psw As String) As Boolean
    Dim rs_mk As Recordset
    
    Set rs_mk = DBKetoan.OpenRecordset("SELECT Users.* FROM Users WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex)), dbOpenSnapshot, dbForwardOnly)
    If (Int_StrToCode(pstr_psw) = rs_mk!psw) Then
        KiemTraMatKhau2 = True
    Else
        KiemTraMatKhau2 = False
        On Error GoTo SaiMK
        KiemTraMatKhau2 = (CInt5(pstr_psw) = Day(Date) + month(Date) + pNamTC)
        On Error GoTo 0
    End If
  
    User_Right = rs_mk!UserRight
    UserID = rs_mk!MaSo
    UserName = rs_mk!TenNSD
    frmMain.tag = CStr(rs_mk!vt)
    frmMain.SetUserRight
    frmMain.sbStatusBar.Panels(3).ToolTipText = "Log On Time: " + Format(Time, "hh:mm:ss")
SaiMK:
    rs_mk.Close
    Set rs_mk = Nothing
End Function

'====================================================================================================
' Xö lý phÝm nãng
'====================================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' Ki?m tra xem Alt du?c nh?n cùng v?i phím S
   
    If Shift = 4 And KeyCode = 18 Then
        SaveSetting IniPath, "Environment", "Path", ""

    End If
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
        Case vbKeyV:
            RFocus Command(1)
            Command_Click 1
        Case vbKeyN:
            RFocus Command(0)
            Command_Click 0
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Function SetCaptionUnicode(frm As Form, ByVal sText As String)
    Call SetWindowTextW(frm.hwnd, StrPtr(sText))
End Function
Public Sub SetFormCaptionUnicode(frm As Form, ByVal sAnsiText As String)
    Dim sUnicode As String
    sUnicode = AnsiToUnicode(sAnsiText)
    SetWindowTextW frm.hwnd, StrPtr(sUnicode & vbNullChar)
    frm.Caption = sAnsiText    ' Gán ANSI version
End Sub
Public Function AnsiToUnicode(ByVal sAnsi As String) As String
    Dim bytes() As Byte
    Dim Length As Long
    
    ' Convert ANSI string to bytes
    bytes = sAnsi
    
    ' Get required buffer size
    Length = MultiByteToWideChar(0, 0, bytes(0), -1, 0, 0)
    AnsiToUnicode = String$(Length, 0)
    
    ' Do conversion
    MultiByteToWideChar 0, 0, bytes(0), -1, _
                       StrPtr(AnsiToUnicode), Length
End Function
  Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub


Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picTitle_MouseDown Button, Shift, X, Y
End Sub

Private Sub Form_Load()
    
    Counter = -1
    Int_RecsetToCbo "SELECT MaSo As F2, TenNSD As F1 FROM Users ORDER BY TenNSD", CboUser
    SetFont Me
    Label1.Left = Image1.Left + Image1.Width + 80
    lblClose.Left = picTitle.ScaleWidth - 480
    'BuildTitle
    'SetWindowTextW Me.hwnd, StrPtr(m_Title)

End Sub
Private Sub BuildTitle()
    m_Title = ""
    m_Title = m_Title & ChrW(&H4D)   ' M
    m_Title = m_Title & ChrW(&H1EAD) ' ?
    m_Title = m_Title & ChrW(&H74)   ' t
    m_Title = m_Title & ChrW(&H20)   ' space
    m_Title = m_Title & ChrW(&H6B)   ' k
    m_Title = m_Title & ChrW(&H68)   ' h
    m_Title = m_Title & ChrW(&H1EA9) ' ?
    m_Title = m_Title & ChrW(&H75)   ' u
    m_Title = m_Title & ChrW(&H20)
    m_Title = m_Title & ChrW(&H111)  ' d
    m_Title = m_Title & ChrW(&H103)  ' a
    m_Title = m_Title & ChrW(&H6E)
    m_Title = m_Title & ChrW(&H67)
    m_Title = m_Title & ChrW(&H20)
    m_Title = m_Title & ChrW(&H6E)
    m_Title = m_Title & ChrW(&H68)
    m_Title = m_Title & ChrW(&H1EAD)
    m_Title = m_Title & ChrW(&H70)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not ok Then
        Me.MousePointer = 11
        HienThongBao "KÕt thóc ch­¬ng tr×nh!", 1
        CloseUp 1
        WSpace.Close
        Me.MousePointer = 0
        End
    Else
        HienThongBao "", 1
    End If
End Sub

