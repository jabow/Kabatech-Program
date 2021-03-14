VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formWWF 
   Caption         =   "Wind With Feed"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   OleObjectBlob   =   "formWWF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formWWF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
    If Standard.wind_with_feed(ComboBox1.Value, ComboBox2.Value, TextBox1.Value) = False Then GoTo error
error:
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub UserForm_Activate()

ComboBox1.AddItem ("400")
ComboBox1.Value = "400"
ComboBox2.AddItem ("18")
ComboBox2.AddItem ("40")
ComboBox2.Value = "40"

End Sub
