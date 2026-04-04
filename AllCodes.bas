Attribute VB_Name = "AllCodes"
'What you see on the computer screen isn't what you will get when you print,
'the computer screen doesn't have the same resolution as a printer, therefore
'lines might appear to "merge" on the screen.
'The values in varBar1 are the available text in a given Barcode language to be printed
'The values in varBar2 are the Barcode equivalent of the text in varBar1
'sBar is the accumulated Barcode equivalents of the text to be printed
'The Barcode() Function will print one character of sBar at a time in a loop
'To add more Barcode types, just continue to build functions that make the appropriate sBar String
Option Explicit





Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Private Declare Function GetAdaptersInfo Lib "iphlpapi" (lpAdapterInfo As Any, lpSize As Long) As Long

Public Const CLR_MENU_NORMAL As Long = &HE0E0E0
Public Const CLR_MENU_HOVER As Long = &HC0C0FF

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2


Public Const TITLE_HEIGHT As Long = 350   ' twips
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function InsertMenuItemW Lib "user32" ( _
                                        ByVal hMenu As Long, _
                                        ByVal uItem As Long, _
                                        ByVal fByPosition As Long, _
                                        lpmii As Any) As Long
' ByRef ? dây, và lpmii As MENUITEMINFOW

Public Declare Function TrackPopupMenu Lib "user32" ( _
                                       ByVal hMenu As Long, _
                                       ByVal wFlags As Long, _
                                       ByVal X As Long, _
                                       ByVal Y As Long, _
                                       ByVal nReserved As Long, _
                                       ByVal hwnd As Long, _
                                       ByVal prcRect As Long) As Long   ' Ð? hi?n th? popup
Public Declare Function SetForegroundWindow Lib "user32" ( _
                                            ByVal hwnd As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Const TPM_LEFTALIGN = &H0
Public Const TPM_TOPALIGN = &H0
Public Const TPM_RIGHTBUTTON = &H2
Public Const TPM_RETURNCMD = &H100
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function SetWindowTextW Lib "user32" _
                                       (ByVal hwnd As Long, ByVal lpString As Long) As Long

Private Declare Function LoadStringW Lib "user32" ( _
                                     ByVal hInstance As Long, ByVal uID As Long, _
                                     ByVal lpBuffer As Long, ByVal nBufferMax As Long) As Long
Declare Function MessageBoxW Lib "user32" ( _
                             ByVal hwnd As Long, _
                             ByVal lpText As Long, _
                             ByVal lpCaption As Long, _
                             ByVal wType As Long) As Long


Public Const MIIM_STRING = &H40
Public Const MIIM_FTYPE = &H100
Public Const MFT_STRING = &H0
Public Type MENUITEMINFOW
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long  ' Ðây là LPWSTR (pointer to wide string)
    cch As Long
    hbmpItem As Long    ' Cho Vista+ n?u c?n icon
End Type

Public Declare Function SetMenuItemInfoW Lib "user32" _
                                         (ByVal hMenu As Long, ByVal uItem As Long, _
                                          ByVal fByPosition As Long, _
                                          ByRef lpmii As MENUITEMINFOW) As Long

Public Declare Function GetMenu Lib "user32" _
                                (ByVal hwnd As Long) As Long

Public Declare Function GetSubMenu Lib "user32" _
                                   (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Public Const pBCode = 39


Public LO_XXXX As String
Public SL_XXXX As Double
Dim sBar As String, i0 As Integer, i1 As Integer
Attribute i0.VB_VarUserMemId = 1073741826
Attribute i1.VB_VarUserMemId = 1073741826
Public Function VniToUnicode2(ByVal sVNI As String) As String
    Dim VNI As String, UNI As String
    Dim arrVNI() As String, arrUNI() As String
    Dim i As Long, j As Long, result As String

    ' B?ng mapping VNI Windows ? Unicode (hex)
    VNI = "aù,aø,aû,aõ,aï,aâ,aê,aá,aà,aå,aã,aä,aé,aè,aú,aü,aë," & _
          "AÙ,AØ,AÛ,AÕ,AÏ,AÂ,AÊ,AÁ,AÀ,AÅ,AÃ,AÄ,AÉ,AÈ,AÚ,AÜ,AË," & _
          "eù,eø,eû,eõ,eï,eâ,eá,eà,eå,eã,eä," & _
          "EÙ,EØ,EÛ,EÕ,EÏ,EÂ,EÁ,EÀ,EÅ,EÃ,EÄ," & _
          "où,oø,oû,oõ,oï,oâ,ô,oá,oà,oå,oã,oä,ôù,ôø,ôû,ôõ,ôï," & _
          "OÙ,OØ,OÛ,OÕ,OÏ,OÂ,Ô,OÁ,OÀ,OÅ,OÃ,OÄ,ÔÙ,ÔØ,ÔÛ,ÔÕ,ÔÏ," & _
          "uù,uø,uû,uõ,uï,ö,öù,öø,öû,öõ,öï," & _
          "UÙ,UØ,UÛ,UÕ,UÏ,Ö,ÖÙ,ÖØ,ÖÛ,ÖÕ,ÖÏ," & _
          "yù,yø,yû,yõ,yï," & _
          "YÙ,YØ,YÛ,YÕ,YÏ," & _
          "ñ,Ñ,ñ,Ñ,î,Î,d,Ð"

    UNI = "E1,E0,1EA3,E3,1EA1,E2,103,1EA5,1EA7,1EA9,1EAB,1EAD,1EAF,1EB1,1EB3,1EB5,1EB7," & _
          "C1,C0,1EA2,C3,1EA0,C2,102,1EA4,1EA6,1EA8,1EAA,1EAC,1EAE,1EB0,1EB2,1EB4,1EB6," & _
          "E9,E8,1EBB,1EBD,1EB9,EA,1EBF,1EC1,1EC3,1EC5,1EC7," & _
          "C9,C8,1EBA,1EBC,1EB8,CA,1EBE,1EC0,1EC2,1EC4,1EC6," & _
          "F3,F2,1ECF,F5,1ECD,F4,1A1,1ED1,1ED3,1ED5,1ED7,1ED9," & _
          "D3,D2,1ECE,D5,1ECC,D4,1A0,1ED0,1ED2,1ED4,1ED6,1ED8," & _
          "FA,F9,1EE7,169,1EE5,1B0,1EE9,1EEB,1EED,1EEF,1EF1," & _
          "DA,D9,1EE6,168,1EE4,1AF,1EE8,1EEA,1EEC,1EEE,1EF0," & _
          "FD,1EF3,1EF7,1EF9,1EF5," & _
          "DD,1EF2,1EF6,1EF8,1EF4," & _
          "111,110,ñ,Ñ,î,Î,111,110"

    arrVNI = Split(VNI, ",")
    arrUNI = Split(UNI, ",")

    result = sVNI   ' B?t d?u t? nguyên b?n

    For i = 0 To UBound(arrVNI)
        If Len(arrVNI(i)) > 0 Then
            result = Replace(result, arrVNI(i), ChrW(CLng("&H" & arrUNI(i))))
        End If
    Next i

    VniToUnicode2 = result
End Function
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

Public Function GetResString(id As Long) As String
    On Error Resume Next
    Dim B() As Byte
    B = LoadResData(id, "CUSTOM")   ' type là CUSTOM (vi?t hoa)
    If UBound(B) >= 0 Then
        GetResString = StrConv(B, vbUnicode)   ' convert byte ? Unicode, gi? d?u d?p
        ' Xóa null terminator n?u có
        If Right(GetResString, 1) = Chr(0) Then GetResString = Left(GetResString, Len(GetResString) - 1)
    End If
End Function
Private Function StrToUnicode(ByVal str As String) As Long
    Dim byteArray() As Byte
    byteArray = StrConv(str, vbUnicode)
    StrToUnicode = VarPtr(byteArray(0))
End Function
Public Function Code39(strCode As String)
Dim varBar1, varBar2
    varBar1 = Split("0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,-,., ,$,/,+,%,*", ",")
    varBar2 = Split("111221211,211211112,112211112,212211111,111221112,211221111,112221111,111211212,211211211,112211211,211112112,112112112,212112111,111122112,211122111,112122111,111112212,211112211,112112211,111122211,211111122,112111122,212111121,111121122,211121121,112121121,111111222,211111221,112111221,111121221,221111112,122111112,222111111,121121112,221121111,122121111,121111212,221111211,122111211,121212111,121211121,121112121,111212121,121121211", ",")
sBar = "121121211" & "1"
For i0 = 1 To Len(strCode)
    For i1 = 0 To UBound(varBar1)
        If Mid(strCode, i0, 1) = varBar1(i1) Then sBar = sBar & varBar2(i1) & "1"
    Next
Next
sBar = sBar & "121121211"
End Function

Public Function Code128(strCode As String)
Dim varBar1, varBar2
    varBar1 = Split(" <>!<>" & Chr(34) & "<>#<>$<>%<>&<>'<>(<>)<>*<>+<>,<>-<>.<>/<>0<>1<>2<>3<>4<>5<>6<>7<>8<>9<>:<>;<><<>=<>><>?<>@<>A<>B<>C<>D<>E<>F<>G<>H<>I<>J<>K<>L<>M<>N<>O<>P<>Q<>R<>S<>T<>U<>V<>W<>X<>Y<>Z<>[<>\<>]<>^<>_<>`<>a<>b<>c<>d<>e<>f<>g<>h<>i<>j<>k<>I<>m<>n<>o<>p<>q<>r<>s<>t<>u<>v<>w<>x<>y<>z<>{<>|<>}<>~<>DEL<>FNC 3<>FNC 2<>SHIFT<>CODE C<>FNC 4<>CODE A<>FNC 1<>Start A<>Start B<>Start C<>Stop", "<>")
    varBar2 = Split("212222,222122,222221,121223,121322,131222,122213,122312,132212,221213,221312,231212,112232,122132,122231,113222,123122,123221,223211,221132,221231,213212,223112,312131,311222,321122,321221,312212,322112,322211,212123,212321,232121,111323,131123,131321,112313,132113,132311,211313,231113,231311,112133,112331,132131,113123,113321,133121,313121,211331,231131,213113,213311,213131,311123,311321,331121,312113,312311,332111,314111,221411,431111,111224,111422,121124,121421,141122,141221,112214,112412,122114,122411,142112,142211,241211,221114,413111,241112,134111,111242,121142,121241,114212,124112,124211,411212,421112,421211,212141,214121,412121,111143,111341,131141,114113,114311,411113,411311,113141,114131,311141,411131,211412,211214,211232,2331112", ",")
Dim chksum As Single: chksum = 104
sBar = "211214"
For i0 = 1 To Len(strCode)
    For i1 = 0 To UBound(varBar1)
        If Mid(strCode, i0, 1) = varBar1(i1) Then
            sBar = sBar & varBar2(i1)
            chksum = chksum + (i1 * i0)
            Exit For
        End If
    Next
Next
sBar = sBar & varBar2(chksum - (Int(chksum / 103) * 103)) & "2331112"
End Function

Public Function BarCode(strCode As String, Pic As Object, barscale As Integer, barHeight As Single, StartX As Single, startY As Single)
Dim barWidth As Single, i0 As Integer, barStart As Single

'Select Case pBCode
'    Case 39:    strCode = UCase(strCode): Code39 strCode
'    Case 128:   Code128 strCode
'End Select

Code128 strCode

barStart = StartX
For i0 = 1 To Len(sBar)
    barWidth = Mid(sBar, i0, 1) * barscale
    If i0 Mod 2 > 0 Then Pic.Line (barStart, startY)-Step(barWidth, barHeight), vbBlack, BF
    barStart = barStart + IIf(i0 Mod 2 > 0, barWidth, barWidth * 1.3)
Next

End Function

Public Function PrintBarCode(vt As ClsVattu, sl As Integer) As Double
    Dim i As Integer, j As Integer
    
    i = sl \ 3 + IIf(sl Mod 3 <> 0, 1, 0)
    ExecuteSQL5 "DELETE * FROM BarCode"
    For j = 1 To i
        ExecuteSQL5 "INSERT INTO BarCode (MaSo,BarCode, Ten, GiaBan) VALUES (" + CStr(Lng_MaxValue("MaSo", "BarCode") + 1) + ",'" + vt.sohieu + "','" + vt.TenVattu + "'," + DoiDau(vt.GiaBan1) + ")"
    Next
    
    SetRptInfo
    frmMain.Rpt.ReportFileName = "BARCODE.RPT"
    frmMain.Rpt.PrinterName = "DATAMAX DMX I-4208"
    frmMain.Rpt.Destination = crptToWindow
    InBaoCaoRPT
    
End Function

