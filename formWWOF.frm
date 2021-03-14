VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formWWOF 
   Caption         =   "Wind Without Feed"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   OleObjectBlob   =   "formWWOF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formWWOF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Standard.wind_without_feed(ComboBox1.Value, ComboBox2.Value) = False Then GoTo error
error:
    Unload Me
End Sub

Private Sub UserForm_Activate()
    ComboBox1.AddItem ("400")
    ComboBox1.Value = "400"
    ComboBox2.AddItem ("2")
    ComboBox2.AddItem ("3")
    ComboBox2.Value = 2
End Sub
