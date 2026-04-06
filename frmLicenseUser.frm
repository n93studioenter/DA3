VERSION 5.00
Begin VB.Form frmLicenseUser 
   BorderStyle     =   0  'None
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   Picture         =   "frmLicenseUser.frx":0000
   ScaleHeight     =   2550
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFakeTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   13575
      TabIndex        =   12
      Top             =   0
      Width           =   13575
      Begin VB.Label lblTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   " License"
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
         TabIndex        =   14
         Top             =   0
         Width           =   3000
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
         Left            =   7320
         TabIndex        =   13
         Top             =   0
         Width           =   480
      End
      Begin VB.Image picIcon 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   120
         Picture         =   "frmLicenseUser.frx":D51B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   8550
         Index           =   0
         Left            =   840
         Picture         =   "frmLicenseUser.frx":D7D8
         Stretch         =   -1  'True
         Top             =   240
         Width           =   7890
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7350
      TabIndex        =   11
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KÝch ho¹t"
      Height          =   285
      Index           =   1
      Left            =   6480
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   3
      Left            =   3360
      MaxLength       =   6
      TabIndex        =   9
      Top             =   1560
      Width           =   835
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   2
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1560
      Width           =   1000
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   13
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   0
      Left            =   360
      MaxLength       =   100
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   4
      Left            =   4200
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1560
      Width           =   940
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   5
      Left            =   5160
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1560
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sao chÐp"
      Height          =   300
      Index           =   0
      Left            =   6480
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "NhÊn sao chÐp m· ®Ó göi l¹i cho chóng t«i"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "lbl1"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Vui lßng ®¨ng ký ®Ó ®­îc sö dông phÇn mÒm"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmLicenseUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SECRET_KEY_MST As Long = &H7B4D8E2F
Private isNewActive As Boolean
Private Const RANDOM_MIN As Long = 10
Private Const RANDOM_MAX As Long = 99
Private Const SECRET_KEY2 As Long = 13579
Private Const SECRET_KEYMac As Long = &H5A3C9F2E
Private Const SECRET_KEY As Byte = &H5A    ' Khóa bí m?t
Private Const BASE36 As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Private Const MAGIC_NUMBER As Long = 7919
Private Const CHARSET As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Dim bakStr As String

'Mã hóa cho so chung tu
Public Function EncodeLicense6(ByVal num As Long, ByRef randomNum As Long) As String
'Dim randomNum As Long
    Dim combined As Variant   ' <--- Dùng Variant
    Dim encrypted As Variant  ' <--- Dùng Variant
    Dim result As String

    On Error GoTo ErrorHandler

    ' Gi?i h?n s? d?u vào du?i 2 tri?u
    If num > 1999999 Then
        EncodeLicense6 = "ERROR"
        Exit Function
    End If

    ' T?o s? random 3 ch? s?
    'Randomize
    'randomNum = Int((RANDOM_MAX - RANDOM_MIN + 1) * Rnd + RANDOM_MIN)

    ' K?t h?p: random (3 ch? s?) + num (t?i da 1,999,999)
    combined = randomNum * 2000000
    combined = combined + num

    ' Mã hóa
    encrypted = combined Xor SECRET_KEY
    encrypted = encrypted + 12345

    ' Gi?i h?n trong 36^6 = 2,176,782,336
    encrypted = encrypted - Int(encrypted / 2176782336#) * 2176782336#

    ' Chuy?n sang base36 (6 ký t?)
    result = ""
    Do While encrypted > 0
        result = Mid(CHARSET, (encrypted Mod 36) + 1, 1) & result
        encrypted = Int(encrypted / 36)
    Loop

    EncodeLicense6 = Right(String(6, "0") & result, 6)
    Exit Function

ErrorHandler:
    EncodeLicense6 = "ERROR"
End Function

' ==============================================
' DECODE: Mã 6 ký t? -> s? g?c và random
' ==============================================
Public Function DecodeLicense6(ByVal txt As String, ByRef randomNum As Long) As Long
    Dim i As Integer
    Dim result As Variant   ' <--- Dùng Variant
    Dim combined As Variant
    
    On Error GoTo ErrorHandler
    
    ' Gi?i mã base36
    result = 0
    For i = 1 To Len(txt)
        result = result * 36 + (InStr(CHARSET, Mid(txt, i, 1)) - 1)
    Next
    
    ' Gi?i mã
    result = result - 12345
    combined = result Xor SECRET_KEY
    
    ' Tách random và s? g?c
    randomNum = Int(combined / 2000000)
    DecodeLicense6 = combined - (randomNum * 2000000)
    Exit Function
    
ErrorHandler:
    DecodeLicense6 = -1
End Function

' ==============================================
' ENCODE: Có th? truy?n random vào (dùng cho test)
' ==============================================

' ==============================================
' ENCODE: Có th? truy?n random vào (dùng cho test)
' ==============================================
Public Function EncodeLicense6Ex(ByVal num As Long, ByVal randomNum As Long) As String
    Dim combined As Currency   ' <--- Ð?i thành Currency
    Dim encrypted As Long
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    If num > 1999999 Or randomNum < 100 Or randomNum > 999 Then
        EncodeLicense6Ex = "ERROR"
        Exit Function
    End If
    
    ' Dùng Currency d? tránh tràn
    combined = CDec(randomNum) * 2000000 + num
    
    encrypted = (CLng(combined) Xor SECRET_KEY) + 12345
    encrypted = encrypted Mod 2176782336#
    
    result = ""
    Do While encrypted > 0
        result = Mid(CHARSET, (encrypted Mod 36) + 1, 1) & result
        encrypted = encrypted \ 36
    Loop
    
    EncodeLicense6Ex = Right(String(6, "0") & result, 6)
    Exit Function
    
ErrorHandler:
    EncodeLicense6Ex = "ERROR"
End Function

' ==============================================
' GI? L?I HÀM CU (4 ký t?) d? tuong thích
' ==============================================
Public Function EncodeLicense4(ByVal num As Long) As String
    Dim encrypted As Long
    Dim result As String
    
    If num > 999999 Then
        EncodeLicense4 = "ERROR"
        Exit Function
    End If
    
    encrypted = ((num Xor SECRET_KEY) + 12345) Mod 1679616
    
    result = ""
    Do While encrypted > 0
        result = Mid(CHARSET, (encrypted Mod 36) + 1, 1) & result
        encrypted = encrypted \ 36
    Loop
    
    EncodeLicense4 = Right("0000" & result, 4)
End Function

Public Function DecodeLicense4(ByVal txt As String) As Long
    Dim i As Integer
    Dim result As Long
    
    result = 0
    For i = 1 To Len(txt)
        result = result * 36 + (InStr(CHARSET, Mid(txt, i, 1)) - 1)
    Next
    
    result = result - 12345
    result = result Xor SECRET_KEY
    
    DecodeLicense4 = result
End Function

Public Function EncodeLicense2(ByVal num As Long) As String
    Dim encrypted As Long
    Dim result As String
    
    encrypted = (num Xor SECRET_KEY) + 12345
    
    result = ""
    Do While encrypted > 0
        result = Mid(CHARSET, (encrypted Mod 36) + 1, 1) & result
        encrypted = encrypted \ 36
    Loop
    
    EncodeLicense2 = result
End Function

Public Function DecodeLicense2(ByVal txt As String) As Long
    Dim i As Integer
    Dim result As Long
    
    result = 0
    For i = 1 To Len(txt)
        result = result * 36 + (InStr(CHARSET, Mid(txt, i, 1)) - 1)
    Next
    
    result = result - 12345
    result = result Xor SECRET_KEY
    
    DecodeLicense2 = result
End Function

'Cho mac
 Public Function EncodeMAC12(ByVal macAddress As String) As String
    Dim cleanMAC As String
    Dim p1 As Long, p2 As Long
    Dim code1 As String, code2 As String
    
    ' Xóa d?u :
    cleanMAC = UCase(Replace(macAddress, ":", ""))
    
    If Len(cleanMAC) <> 12 Then
        EncodeMAC12 = "ERROR"
        Exit Function
    End If
    
    ' Tách 12 hex thành 2 ph?n, m?i ph?n 6 hex (3 byte)
    p1 = CLng("&H" & Mid(cleanMAC, 1, 6))   ' 6 hex d?u
    p2 = CLng("&H" & Mid(cleanMAC, 7, 6))   ' 6 hex cu?i
    
    ' Encode m?i ph?n thành 6 ký t?
    code1 = Encode6(p1)
    code2 = Encode6(p2)
    
    ' N?i l?i v?i d?u -
    EncodeMAC12 = code1 & "-" & code2
End Function

' ==============================================
' DECODE MAC t? 12 ký t? (2 block x 6 ký t?)
' ==============================================
Public Function DecodeMAC12(ByVal code As String) As String
    Dim parts() As String
    Dim p1 As Long, p2 As Long
    Dim hexMAC As String
    
    ' Tách 2 ph?n
    parts = Split(code, "-")
    
    If UBound(parts) <> 1 Then
        DecodeMAC12 = "ERROR"
        Exit Function
    End If
    
    ' Gi?i mã t?ng ph?n
    p1 = Decode6(parts(0))
    p2 = Decode6(parts(1))
    
    ' Ghép l?i thành MAC 12 hex
    hexMAC = Right("000000" & Hex(p1), 6) & _
             Right("000000" & Hex(p2), 6)
    
    DecodeMAC12 = UCase(hexMAC)
End Function

' ==============================================
' ENCODE: 6 hex (3 byte) -> 6 ký t? base36 (có random)
' ==============================================
Public Function Encode6(ByVal num As Long) As String
    Dim randomNum As Long
    Dim combined As Variant
    Dim encrypted As Variant
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    ' Gi?i h?n num trong 0-16,777,215 (6 hex = 24 bit)
    num = num And &HFFFFFF
    
    ' T?o random 3 ch? s?
    Randomize
    randomNum = Int((RANDOM_MAX - RANDOM_MIN + 1) * Rnd + RANDOM_MIN)
    
    ' K?t h?p: random (3 ch? s?) * 16,777,216 + num
    ' 16,777,216 = 2^24
    combined = CDec(randomNum) * 16777216
    combined = combined + num
    
    ' Mã hóa
    encrypted = combined Xor SECRET_KEY2
    encrypted = encrypted + 12345
    
    ' Gi?i h?n trong 36^6 = 2,176,782,336
    encrypted = encrypted - Int(encrypted / 2176782336#) * 2176782336#
    
    ' Chuy?n sang base36 (6 ký t?)
    result = ""
    Do While encrypted > 0
        result = Mid(CHARSET, (encrypted Mod 36) + 1, 1) & result
        encrypted = Int(encrypted / 36)
    Loop
    
    Encode6 = Right(String(6, "0") & result, 6)
    Exit Function
    
ErrorHandler:
    Encode6 = "ERROR"
End Function

' ==============================================
' DECODE: 6 ký t? base36 -> 6 hex (3 byte)
' ==============================================
Public Function Decode6(ByVal txt As String) As Long
    Dim i As Integer
    Dim result As Variant
    Dim combined As Variant
    
    On Error GoTo ErrorHandler
    
    ' Gi?i mã base36
    result = 0
    For i = 1 To Len(txt)
        result = result * 36 + (InStr(CHARSET, Mid(txt, i, 1)) - 1)
    Next
    
    ' Gi?i mã
    result = result - 12345
    combined = result Xor SECRET_KEY2
    
    ' L?y num (b? random)
    Decode6 = combined Mod 16777216
    Exit Function
    
ErrorHandler:
    Decode6 = -1
End Function

' ==============================================
' Ð?NH D?NG MAC CÓ D?U ":"
' ==============================================
Public Function FormatMAC(ByVal mac As String) As String
    Dim result As String
    Dim i As Integer
    
    If Len(mac) <> 12 Then
        FormatMAC = mac
        Exit Function
    End If
    
    result = ""
    For i = 1 To Len(mac) Step 2
        If result <> "" Then result = result & ":"
        result = result & Mid(mac, i, 2)
    Next i
    
    FormatMAC = result
End Function
 

'random 6 cho 1_2
Public Function EncodeWithRandom6(ByVal strValue As String, ByRef randomNum As Long) As String
    Dim parts() As String
    Dim num1 As Long   ' <--- S?A: num1, không ph?i numl
    Dim num2 As Long
    Dim combined As Long
    
    On Error GoTo ErrorHandler
    
    ' Tách chu?i d?ng "1_2"
    parts = Split(strValue, "_")
    
    If UBound(parts) <> 1 Then
        EncodeWithRandom6 = "ERROR"
        Exit Function
    End If
    
    num1 = CLng(parts(0))
    num2 = CLng(parts(1))
    
    ' Gi?i h?n m?i s? trong 0-999
    If num1 > 999 Or num2 > 999 Then   ' <--- S?A: num1
        EncodeWithRandom6 = "ERROR"
        Exit Function
    End If
    
    ' T?o s? random 4 ch? s?
    'Randomize
    'randomNum = Int((RANDOM_MAX - RANDOM_MIN + 1) * Rnd + RANDOM_MIN)
    
    ' K?t h?p: randomNum (4 ch? s?) + num1 (3 ch? s?) + num2 (3 ch? s?)
    combined = randomNum * 1000000 + num1 * 1000 + num2   ' <--- S?A: num1
    
    ' Mã hóa thành 6 ký t?
    EncodeWithRandom6 = ToBase36_6(combined)
    Exit Function
    
ErrorHandler:
    EncodeWithRandom6 = "ERROR"
End Function

' ==============================================
' DECODE: 6 ký t? -> chu?i "x_y" và random
' ==============================================
Public Function DecodeWithRandom6(ByVal code As String, ByRef randomNum As Long) As String
    Dim combined As Long
    Dim num1 As Long
    Dim num2 As Long
    
    On Error GoTo ErrorHandler
    
    ' Gi?i mã base36
    combined = FromBase36_6(code)
    
    ' Tách thành random, num1, num2
   ' randomNum = combined \ 1000000
    num1 = (combined Mod 1000000) \ 1000
    num2 = combined Mod 1000
    
    DecodeWithRandom6 = num1 & "_" & num2
    Exit Function
    
ErrorHandler:
    DecodeWithRandom6 = "ERROR"
End Function

' ==============================================
' ENCODE: Có th? truy?n random vào (dùng cho test)
' ==============================================
Public Function EncodeWithRandom6Ex(ByVal strValue As String, ByVal randomNum As Long) As String
    Dim parts() As String
    Dim num1 As Long
    Dim num2 As Long
    Dim combined As Long
    
    On Error GoTo ErrorHandler
    
    ' Tách chu?i d?ng "1_2"
    parts = Split(strValue, "_")
    
    If UBound(parts) <> 1 Then
        EncodeWithRandom6Ex = "ERROR"
        Exit Function
    End If
    
    num1 = CLng(parts(0))
    num2 = CLng(parts(1))
    
    ' Gi?i h?n
    If num1 > 999 Or num2 > 999 Or randomNum < 1000 Or randomNum > 9999 Then
        EncodeWithRandom6Ex = "ERROR"
        Exit Function
    End If
    
    ' K?t h?p: randomNum (4 ch? s?) + num1 (3 ch? s?) + num2 (3 ch? s?)
    combined = randomNum * 1000000 + num1 * 1000 + num2
    
    ' Mã hóa thành 6 ký t?
    EncodeWithRandom6Ex = ToBase36_6(combined)
    Exit Function
    
ErrorHandler:
    EncodeWithRandom6Ex = "ERROR"
End Function

' ==============================================
' HÀM H? TR?: Chuy?n s? sang base36 (6 ký t?)
' ==============================================
Private Function ToBase36_6(ByVal num As Long) As String
    Dim result As String
    
    If num = 0 Then
        ToBase36_6 = "000000"
        Exit Function
    End If
    
    result = ""
    Do While num > 0
        result = Mid(CHARSET, (num Mod 36) + 1, 1) & result
        num = num \ 36
    Loop
    
    ToBase36_6 = Right(String(6, "0") & result, 6)
End Function

' ==============================================
' HÀM H? TR?: Gi?i mã base36
' ==============================================
Private Function FromBase36_6(ByVal txt As String) As Long
    Dim i As Integer

    FromBase36_6 = 0
    For i = 1 To Len(txt)
        FromBase36_6 = FromBase36_6 * 36 + (InStr(CHARSET, Mid(txt, i, 1)) - 1)
    Next i
End Function
'cho mst
 Public Function EncodeMST8(ByVal mst As String, ByRef randomNum As Long) As String
    Dim cleanMST As String
    Dim totalLen As Integer
    Dim p1 As String, p2 As String
    Dim len1 As Integer, len2 As Integer
    Dim mixed1 As String, mixed2 As String
    
    On Error GoTo ErrorHandler
    
    cleanMST = Replace(mst, "-", "")
    
    If Not IsNumeric(cleanMST) Then GoTo ErrorHandler
    
    totalLen = Len(cleanMST)
    
    Select Case totalLen
        Case 10
            len1 = 5: len2 = 5
        Case 12
            len1 = 6: len2 = 6
        Case 13
            len1 = 6: len2 = 7
        Case Else
            GoTo ErrorHandler
    End Select
    
    p1 = Left(cleanMST, len1)
    p2 = Right(cleanMST, len2)
    
    ' ? KHÔNG Randomize ? dây
    randomNum = Int((RANDOM_MAX - RANDOM_MIN + 1) * Rnd + RANDOM_MIN)
    
    mixed1 = MixString(p1, randomNum)
    mixed2 = MixString(p2, randomNum)
    
    EncodeMST8 = ToBase36_2_Fixed(randomNum) & _
                 ToBase36_1(len1) & ToBase36_1(len2) & _
                 EncodeString(mixed1) & EncodeString(mixed2)
    Exit Function

ErrorHandler:
    EncodeMST8 = "ERROR"
End Function

' ==============================================
' DECODE (ÐÃ FIX M?T S? 0)
' ==============================================
Public Function DecodeMST8(ByVal code As String, ByRef randomNum As Long) As String
    Dim len1 As Integer, len2 As Integer
    Dim mixed1 As String, mixed2 As String
    Dim p1 As String, p2 As String
    Dim rdn As Long
    
    On Error GoTo ErrorHandler

    If Len(code) <> 12 Then GoTo ErrorHandler

    rdn = FromBase36_2_Fixed(Left(code, 2))
    randomNum = rdn
    
    len1 = FromBase36_1(Mid(code, 3, 1))
    len2 = FromBase36_1(Mid(code, 4, 1))

    mixed1 = DecodeString(Mid(code, 5, 4))
    mixed2 = DecodeString(Mid(code, 9, 4))

    ' ?? FIX QUAN TR?NG: gi? l?i s? 0 d?u
    mixed1 = Right(String(len1, "0") & mixed1, len1)
    mixed2 = Right(String(len2, "0") & mixed2, len2)

    p1 = UnmixString(mixed1, rdn)
    p2 = UnmixString(mixed2, rdn)

    DecodeMST8 = p1 & p2
    Exit Function

ErrorHandler:
    DecodeMST8 = "ERROR"
End Function

' ==============================================
' MIX
' ==============================================
Private Function MixString(ByVal s As String, ByVal key As Integer) As String
    Dim result As String, i As Integer
    Dim digit As Integer, keyDigit As Integer
    
    For i = 1 To Len(s)
        digit = CInt(Mid(s, i, 1))
        keyDigit = (key \ (10 ^ ((i - 1) Mod 2))) Mod 10
        digit = (digit + keyDigit) Mod 10
        result = result & digit
    Next i
    
    MixString = result
End Function

' ==============================================
' UNMIX
' ==============================================
Private Function UnmixString(ByVal s As String, ByVal key As Integer) As String
    Dim result As String, i As Integer
    Dim digit As Integer, keyDigit As Integer
    
    For i = 1 To Len(s)
        digit = CInt(Mid(s, i, 1))
        keyDigit = (key \ (10 ^ ((i - 1) Mod 2))) Mod 10
        digit = (digit - keyDigit) Mod 10
        If digit < 0 Then digit = digit + 10
        result = result & digit
    Next i
    
    UnmixString = result
End Function

' ==============================================
' BASE36 1 CHAR
' ==============================================
Private Function ToBase36_1(ByVal num As Integer) As String
    If num < 0 Or num > 35 Then num = 0
    ToBase36_1 = Mid(CHARSET, num + 1, 1)
End Function

Private Function FromBase36_1(ByVal c As String) As Integer
    Dim pos As Integer
    pos = InStr(CHARSET, c)
    If pos > 0 Then
        FromBase36_1 = pos - 1
    Else
        FromBase36_1 = 0
    End If
End Function

' ==============================================
' BASE36 2 CHAR
' ==============================================
Private Function ToBase36_2_Fixed(ByVal num As Integer) As String
    Dim high As Integer, low As Integer
    high = num \ 36
    low = num Mod 36
    ToBase36_2_Fixed = Mid(CHARSET, high + 1, 1) & Mid(CHARSET, low + 1, 1)
End Function

Private Function FromBase36_2_Fixed(ByVal code As String) As Integer
    Dim high As Integer, low As Integer
    
    If Len(code) <> 2 Then Exit Function
    
    high = InStr(CHARSET, Mid(code, 1, 1)) - 1
    low = InStr(CHARSET, Mid(code, 2, 1)) - 1
    
    If high < 0 Then high = 0
    If low < 0 Then low = 0
    
    FromBase36_2_Fixed = high * 36 + low
End Function

' ==============================================
' ENCODE STRING ? 4 CHAR
' ==============================================
Private Function EncodeString(ByVal s As String) As String
    Dim num As Double
    Dim i As Integer
    Dim result As String

    num = 0

    For i = 1 To Len(s)
        num = num * 10 + CInt(Mid(s, i, 1))
    Next i

    If num = 0 Then
        EncodeString = "0000"
        Exit Function
    End If

    Do While num > 0
        result = Mid(CHARSET, (num Mod 36) + 1, 1) & result
        num = Int(num / 36)
    Loop

    EncodeString = Right(String(4, "0") & result, 4)
End Function

' ==============================================
' DECODE 4 CHAR ? STRING
' ==============================================
Private Function DecodeString(ByVal code As String) As String
    Dim i As Integer
    Dim num As Double

    num = 0
    
    For i = 1 To Len(code)
        num = num * 36 + (InStr(CHARSET, Mid(code, i, 1)) - 1)
    Next i

    DecodeString = CStr(num)
End Function



Private Sub Command2_Click()
'Quen data
    FrmMatkhau.Command(0).Enabled = True
    FrmMatkhau.Refresh
    Unload Me
    FrmMatkhau.isreload = True
    frmMain.mnHT_Click 0

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13:
        Command1_Click 1
    End Select
End Sub
Private Sub Command1_Click(Index As Integer)
    Clipboard.Clear
    Select Case Index
    Case 0
        Clipboard.SetText Trim(Label1(1).Caption)
    Case 1
        If FrmOptions.KiemTraKey(bakStr) = True Then
            Dim mac As String
            mac = GetMacAddress()
            Dim rsCount As DAO.Recordset

            Set rsCount = DBKetoan.OpenRecordset( _
                          "SELECT COUNT(*) AS Tong FROM Users ", dbOpenSnapshot)

            If rsCount!tong > 1 Then
                ExecuteSQL5 "INSERT INTO Users (TenNSD,Psw,UserRight,VT,TS,HDV,WS,MacAddress,IsReister) VALUES ('Administrator" & (rsCount!tong + 1) & "','" & pNamTC & "',0,1111111111,1,1,'...','" & mac & "',1)"
            Else
                'Kiem tra xem day co phai la chinh chu
                Dim rs_checkus As Recordset
                Set rs_checkus = DBKetoan.OpenRecordset( _
                                 "SELECT * FROM Users WHERE IsReister=1", _
                                 dbOpenSnapshot, dbForwardOnly)
                If rs_checkus.EOF Then
                    ExecuteSQL5 "UPDATE Users SET IsReister=1, Psw='" & pNamTC & "', MacAddress='" & mac & "'"
                Else
                    ExecuteSQL5 "INSERT INTO Users (TenNSD,Psw,UserRight,VT,TS,HDV,WS,MacAddress,IsReister) VALUES ('Administrator" & (rsCount!tong + 1) & "','" & pNamTC & "',0,1111111111,1,1,'...','" & mac & "',1)"
                End If

            End If
            'MsgBox "Kich hoat thanh cong"
            Dim s As String
            s = ChrW(75) & ChrW(237) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(104) & ChrW(111) & ChrW(7841) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(224) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(99) & ChrW(244) & ChrW(110) & ChrW(103)
            MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

            FrmMatkhau.Command(0).Enabled = True
            FrmMatkhau.Refresh
            Unload Me
        Else
            s = ChrW(75) & ChrW(237) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(104) & ChrW(111) & ChrW(7841) & ChrW(116) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(224) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(99) & ChrW(244) & ChrW(110) & ChrW(103)
            MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

        End If
    Case Else
    End Select
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

Private Sub Form_Load()
    lblTitle(11).AutoSize = True
    Me.Height = Me.Height + 350 + 10
    picFakeTitle.Width = Me.ScaleWidth
    picFakeTitle.Height = 325
    picIcon(1).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2
    lblTitle(11).Left = picIcon(1).Left + picIcon(1).Width + 90
    lblTitle(11).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2 + 15
    lblClose.Top = 55
    AnControl Me

     lblTitle(11).Caption = "License"
    Me.KeyPreview = True
    Dim strValue As String
    Dim rs As Recordset
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW License.* FROM License", dbOpenSnapshot)
    Dim rls As Recordset
    Set rls = DBKetoan.OpenRecordset("SELECT DISTINCTROW tbLicensekey.* FROM tbLicensekey", dbOpenSnapshot)

    Dim mst10 As String
    Dim mst13 As String
    Dim encoded10 As String
    Dim encoded13 As String
    Dim decoded10 As String
    Dim decoded13 As String

    Debug.Print "=== ENCODE MST THÀNH 8 KÝ T? ==="
    Debug.Print ""
    Dim randomNum As Long
    ' Test MST 10 s?
    mst10 = rs!masothue
    Dim encoded1 As String, encoded2 As String
    Dim decoded1 As String, decoded2 As String
    Dim random1 As Integer, random2 As Integer

    encoded1 = EncodeMST14(mst10, randomNum)
    decoded1 = DecodeMST14(encoded1, randomNum)
    If decoded1 <> mst10 Then
        MsgBox "Chua can"
        encoded1 = EncodeMST8(mst10, randomNum)
        decoded1 = DecodeMST8(encoded1, randomNum)
    End If
    '---------------------


    Dim code As String
    Dim decoded As String

    Dim randomNum2 As Long

    Debug.Print "=== ENCODE/DECODE V?I S? RANDOM 4 CH? S? ==="
    Debug.Print ""

    ' Test 1: T? d?ng random
    'strValue = "1_2"
    Dim tp As Integer
    If rls!Type = 1 Then
        tp = 2
    Else
        tp = 1
    End If
    strValue = tp & "_" & rls!Year
    code = EncodeWithRandom6(strValue, randomNum)
    decoded = DecodeWithRandom6(code, randomNum)

    'mac
    Dim mac As String
    Dim encoded As String
    Dim formatted As String

    Debug.Print "=== ENCODE MAC THÀNH 12 KÝ T? (2 BLOCK) ==="
    Debug.Print ""

    ' Test MAC 1
    mac = GetMacAddress()
    encoded = EncodeMAC12(mac)
    decoded = DecodeMAC12(encoded)
    formatted = FormatMAC(decoded)
    'So chung tu phai nho hon 1 tr 6
    Dim license As String
    license = EncodeLicense6(CLng(rls!totals), randomNum)
    Debug.Print "1 -> " & license
    Debug.Print DecodeLicense6(license, randomNum)    ' 1

    Label1(1).Caption = randomNum & "*" & encoded1 & "*" & code & "*" & license & "*" & encoded
End Sub

Private Sub Text_Change(Index As Integer)
 Dim a() As String
    If bakStr = "" Then
        bakStr = Text(0).Text
    End If
    On Error GoTo Error_Handler
    If (Len(Text(0).Text) > 13) Then
        a = Split(Text(0).Text, "*")
        Text(0).Text = a(0)
        Text(1).Text = a(1)
        Text(2).Text = a(2)
        Text(3).Text = a(3)
        Dim A2() As String
        A2 = Split(a(4), "-")
        Text(4).Text = A2(0)
        Text(5).Text = A2(1)
    End If
Error_Handler:
End Sub
'end code mst moi
Public Function EncodeMST14(ByVal mst As String, ByRef randomNum As Long) As String
    Dim cleanMST As String
    Dim totalLen As Integer
    Dim p1 As String, p2 As String
    Dim len1 As Integer, len2 As Integer
    
    On Error GoTo ErrorHandler
    
    cleanMST = Replace(mst, "-", "")
    
    If Not IsNumeric(cleanMST) Then GoTo ErrorHandler
    
    totalLen = Len(cleanMST)
    
    Select Case totalLen
        Case 10
            len1 = 5: len2 = 5
        Case 12
            len1 = 6: len2 = 6
        Case 13
            len1 = 6: len2 = 7
        Case Else
            GoTo ErrorHandler
    End Select
    
    p1 = Left(cleanMST, len1)
    p2 = Right(cleanMST, len2)
    
    Randomize
    randomNum = Int((RANDOM_MAX - RANDOM_MIN + 1) * Rnd + RANDOM_MIN)
    
    p1 = MixString(p1, randomNum)
    p2 = MixString(p2, randomNum)
    
    EncodeMST14 = ToBase36_2(randomNum) & _
                  ToBase36_1(len1) & ToBase36_1(len2) & _
                  ToBase36_N(CLng(p1), 5) & _
                  ToBase36_N(CLng(p2), 5)
    Exit Function

ErrorHandler:
    EncodeMST14 = "ERROR"
End Function

' ==============================================
' DECODE
' ==============================================
Public Function DecodeMST14(ByVal code As String, ByRef randomNum As Long) As String
    Dim len1 As Integer, len2 As Integer
    Dim rdn As Long
    Dim p1 As String, p2 As String
    Dim result As String

    On Error GoTo ErrorHandler

    If Len(code) <> 14 Then GoTo ErrorHandler

    rdn = FromBase36_2(Left(code, 2))
    randomNum = rdn

    len1 = FromBase36_1(Mid(code, 3, 1))
    len2 = FromBase36_1(Mid(code, 4, 1))

    p1 = FromBase36_N(Mid(code, 5, 5))
    p2 = FromBase36_N(Mid(code, 10, 5))

    p1 = Right(String(len1, "0") & p1, len1)
    p2 = Right(String(len2, "0") & p2, len2)

    p1 = UnmixString(p1, rdn)
    p2 = UnmixString(p2, rdn)

    result = p1 & p2

    ' auto format MST
    If Len(result) = 13 Then
        DecodeMST14 = Left(result, 10) & "-" & Right(result, 3)
    Else
        DecodeMST14 = result
    End If

    Exit Function

ErrorHandler:
    DecodeMST14 = "ERROR"
End Function
 Private Function ToBase36_2(ByVal num As Long) As String
    ToBase36_2 = Mid(CHARSET, num \ 36 + 1, 1) & _
                 Mid(CHARSET, num Mod 36 + 1, 1)
End Function

Private Function FromBase36_2(ByVal code As String) As Long
    FromBase36_2 = (InStr(CHARSET, Mid(code, 1, 1)) - 1) * 36 + _
                   (InStr(CHARSET, Mid(code, 2, 1)) - 1)
End Function
Private Function ToBase36_N(ByVal num As Long, ByVal length As Integer) As String
    Dim result As String
    
    Do While num > 0
        result = Mid(CHARSET, (num Mod 36) + 1, 1) & result
        num = num \ 36
    Loop
    
    ToBase36_N = Right(String(length, "0") & result, length)
End Function

Private Function FromBase36_N(ByVal code As String) As String
    Dim i As Integer
    Dim num As Double
    
    num = 0
    
    For i = 1 To Len(code)
        num = num * 36 + (InStr(CHARSET, Mid(code, i, 1)) - 1)
    Next i
    
    FromBase36_N = Format$(num, "0")
End Function
