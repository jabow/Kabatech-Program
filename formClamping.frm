VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formClamping 
   Caption         =   "Clamping"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3525
   OleObjectBlob   =   "formClamping.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formClamping"
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
        Standard.clamping_device (True)
    Else
        Standard.clamping_device (False)
    End If
        
Unload Me
End Sub

Private Sub UserForm_Activate()

ComboBox1.AddItem ("Up")
ComboBox1.AddItem ("Down")
ComboBox1.Value = "Up"

End Sub
