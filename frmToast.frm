VERSION 5.00
Begin VB.Form frmToast 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   LinkTopic       =   "Form4"
   ScaleHeight     =   3210
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmToast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.Left = Screen.Width - Me.Width - 200   ' cách mép ph?i 200 twips (tùy ch?nh)
    Me.Top = Screen.Height - Me.Height - taskbarHeight - 200    ' cách taskbar lên chút
End Sub
Public Sub ShowToast(Title As String, Message As String, Optional Duration As Long = 5000)

    Dim screenWidth As Long
    Dim screenHeight As Long
    Dim taskbarHeight As Long

    ' L?y kích thu?c màn hình
    screenWidth = Screen.Width \ Screen.TwipsPerPixelX
    screenHeight = Screen.Height \ Screen.TwipsPerPixelY

    ' U?c lu?ng chi?u cao taskbar (có th? di?u ch?nh)
    taskbarHeight = 40 * Screen.TwipsPerPixelY

    ' Ð?t v? trí form
    Left = Screen.Width - frmToast.Width - 10 * Screen.TwipsPerPixelX
    Top = Screen.Height - frmToast.Height - 10 * Screen.TwipsPerPixelY

    ' Ð?m b?o form hi?n th? trên cùng
    Me.ZOrder 0
End Sub

