VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formLength 
   Caption         =   "Length"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2670
   OleObjectBlob   =   "formLength.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formLength"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
ActiveCell.Value = TextBox1.text
If TextBox1.text <> vbNullString Then length = TextBox1.text
Unload Me
End Sub

Private Sub bynCANCEL_Click()
    Unload Me
End Sub



Private Sub TextBox1_Change()

    If TextBox1 = vbNullString Then Exit Sub

        If Not IsNumeric(TextBox1) Then

            MsgBox "Only Numbers can be entered!"

            TextBox1 = vbNullString

        End If

End Sub

