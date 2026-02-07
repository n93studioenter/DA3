VERSION 5.00
Begin VB.Form frmSubMenu 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmSubMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public menuIndex As Integer

Private Sub Form_Load()

    Me.BorderStyle = 0
    Me.BackColor = vbWhite
    Me.Width = 2400

    Select Case menuIndex
        Case 0
            AddItem "Ngu?i dùng"
            AddItem "Ð?i m?t kh?u"
            AddItem "Thoát"
        Case 1
            AddItem "Khách hàng"
            AddItem "Nhà cung c?p"
        Case 2
            AddItem "X? lý d? li?u"
        Case 3
            AddItem "Báo cáo t?ng h?p"
    End Select

End Sub
Sub AddItem(ByVal sText As String)

    Dim lbl As Label
    Set lbl = Controls.Add("VB.Label", "lbl" & Controls.count)

    With lbl
        .Caption = sText
        .Left = 120
        .Top = (Controls.count - 1) * 360
        .Width = Me.Width - 240
        .Height = 360
        .BackColor = vbWhite
        .Visible = True
    End With

    Me.Height = lbl.Top + lbl.Height + 120

End Sub


