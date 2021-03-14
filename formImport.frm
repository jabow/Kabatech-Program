VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formImport 
   Caption         =   "Import"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3525
   OleObjectBlob   =   "formImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
    ProgramNum = TextBox1.text
    Me.Tag = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    End
    'Unload Me
End Sub

Private Sub CommandButton1_Click()
Me.Tag = True
Me.Hide
End Sub

Private Sub TextBox1_Change()
    If TextBox1 = vbNullString Then Exit Sub
End Sub

Private Sub TextBox2_Change()
    If TextBox1 = vbNullString Then Exit Sub
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    End
End Sub
