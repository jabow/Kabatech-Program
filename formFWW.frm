VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formFWW 
   Caption         =   "Feed Without WInd"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3480
   OleObjectBlob   =   "formFWW.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formFWW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    Standard.feed_without_wind (TextBox1.Value)
    Unload Me
End Sub
