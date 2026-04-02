VERSION 5.00
Begin VB.Form FrmGetStrNew 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   LinkTopic       =   "Form4"
   ScaleHeight     =   1290
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
         Left            =   6360
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   0
         Width           =   4455
      End
      Begin VB.Image picIcon 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   120
         Picture         =   "FrmGetStrNew.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   8550
         Index           =   0
         Left            =   840
         Picture         =   "FrmGetStrNew.frx":02BD
         Stretch         =   -1  'True
         Top             =   240
         Width           =   7890
      End
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   0
      Left            =   0
      MaxLength       =   100
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   2
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   3
      Left            =   4440
      MaxLength       =   4
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      Height          =   255
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Liªn hÖ nhµ cung cÊp ®Ó cã m· sè cµi ®Æt"
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
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3495
   End
End
Attribute VB_Name = "FrmGetStrNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ALPHABET As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ*"  ' Base 37
Private Const BASE As Long = 37
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
    SetFont Me


    Dim inputData As String
    inputData = "1,2,1000,5e:3e:70:b5:1c:1c"

    Dim encoded As String
    encoded = EncodeCustom(inputData)
    Debug.Print "Encoded: " & encoded

    Debug.Print "Decoded: " & DecodeCustom(encoded)

End Sub
 ' ====================== ENCODE ======================
Public Function EncodeCustom(ByVal InputStr As String) As String
    Dim parts() As String
    Dim i As Long
    Dim result As String
    
    InputStr = Replace(Replace(InputStr, ":", ","), ";", ",")
    parts = Split(InputStr, ",")
    
    For i = LBound(parts) To UBound(parts)
        Dim s As String
        s = Trim$(parts(i))
        If s = "" Then GoTo NextItem
        
        If IsMAC(s) Then
            result = result & EncodeMACGroup(s) & " "
        Else
            result = result & EncodeSmallNumber(val(s)) & " "
        End If
NextItem:
    Next i
    
    EncodeCustom = Trim$(result)
End Function

Private Function IsMAC(ByVal s As String) As Boolean
    Dim t As String
    t = Replace$(Replace$(s, ":", ""), "-", "")
    IsMAC = (Len(t) = 12) And IsHexString(t)
End Function

Private Function IsHexString(ByVal txt As String) As Boolean
    Dim i As Long, c As String
    For i = 1 To Len(txt)
        c = UCase$(Mid$(txt, i, 1))
        If Not ((c >= "0" And c <= "9") Or (c >= "A" And c <= "F")) Then Exit Function
    Next i
    IsHexString = True
End Function

Private Function EncodeSmallNumber(ByVal num As Variant) As String
    EncodeSmallNumber = EncodeBigNum(num)
End Function

Private Function EncodeMACGroup(ByVal mac As String) As String
    Dim hexClean As String
    Dim num As Variant
    Dim i As Long
    
    hexClean = Replace$(Replace$(mac, ":", ""), "-", "")
    
    ' Chuy?n toàn b? MAC hex thành 1 s? l?n (48-bit)
    num = 0
    For i = 1 To Len(hexClean) Step 2
        num = num * 256 + CByte("&H" & Mid$(hexClean, i, 2))
    Next i
    
    ' Encode và thêm d?nh d?ng gi?ng m?u
    EncodeMACGroup = "00*" & AddGroups(EncodeBigNum(num), 7)
End Function

Private Function EncodeBigNum(ByVal num As Variant) As String
    If num = 0 Then
        EncodeBigNum = "0"
        Exit Function
    End If
    
    Dim s As String
    Do While num > 0
        s = Mid$(ALPHABET, (num Mod BASE) + 1, 1) & s
        num = num \ BASE
    Loop
    EncodeBigNum = s
End Function

Private Function AddGroups(ByVal s As String, Optional ByVal GroupLen As Integer = 7) As String
    Dim result As String
    Dim i As Long, cnt As Long
    
    result = ""
    cnt = 0
    For i = Len(s) To 1 Step -1
        result = Mid$(s, i, 1) & result
        cnt = cnt + 1
        If cnt Mod GroupLen = 0 And i > 1 Then
            result = "*" & result
        End If
    Next i
    AddGroups = result
End Function

' ====================== DECODE ======================
Public Function DecodeCustom(ByVal EncodedStr As String) As String
    Dim groups() As String
    Dim i As Long
    Dim result As String
    Dim clean As String
    Dim val As Variant
    
    groups = Split(Trim$(EncodedStr), " ")
    
    For i = LBound(groups) To UBound(groups)
        If Trim$(groups(i)) = "" Then GoTo NextG
        
        clean = Replace$(groups(i), "*", "")
        val = DecodeBigNum(clean)
        
        If Left$(groups(i), 3) = "00*" And Len(clean) >= 8 Then
            result = result & BigNumToMAC(val) & ","
        Else
            result = result & CStr(val) & ","
        End If
NextG:
    Next i
    
    If Right$(result, 1) = "," Then result = Left$(result, Len(result) - 1)
    DecodeCustom = result
End Function

Private Function DecodeBigNum(ByVal s As String) As Variant
    Dim v As Variant: v = 0
    Dim i As Long, pos As Long
    For i = 1 To Len(s)
        pos = InStr(1, ALPHABET, Mid$(s, i, 1), vbTextCompare) - 1
        If pos >= 0 Then v = v * BASE + pos
    Next i
    DecodeBigNum = v
End Function

Private Function BigNumToMAC(ByVal num As Variant) As String
    Dim bytes(5) As Byte
    Dim i As Long
    Dim s As String
    
    For i = 5 To 0 Step -1
        bytes(i) = num Mod 256
        num = num \ 256
    Next i
    
    For i = 0 To 5
        s = s & Right$("0" & Hex$(bytes(i)), 2)
        If i < 5 Then s = s & ":"
    Next i
    BigNumToMAC = LCase$(s)
End Function
