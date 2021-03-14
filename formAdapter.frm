VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formAdapter 
   Caption         =   "Adapter"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3480
   OleObjectBlob   =   "formAdapter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
        
    If ComboBox1.Value = "Up" Then
        Standard.adapter (True)
    Else
        Standard.adapter (False)
    End If
    Unload Me
End Sub

Private Sub UserForm_Activate()

ComboBox1.AddItem ("Up")
ComboBox1.AddItem ("Down")
ComboBox1.Value = "Up"

End Sub
