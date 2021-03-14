VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formConsecLengths 
   Caption         =   "Consecutive Taped Lengths"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4365
   OleObjectBlob   =   "formConsecLengths.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formConsecLengths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()

Dim Lengths(1 To 3) As String
Dim Tapes(1 To 3) As Long
    Lengths(1) = TextBox1.text
    Lengths(2) = TextBox2.text
    Lengths(3) = TextBox3.text
        
    If ComboBox1.Value = "Space Taped" Then
        Tapes(1) = SpaceTaped
    Else: Tapes(1) = fullyTaped
    End If
    If ComboBox2.Value = "Space Taped" Then
        Tapes(2) = SpaceTaped
    Else: Tapes(2) = fullyTaped
    End If
    If ComboBox3.Value = "Space Taped" Then
        Tapes(3) = SpaceTaped
    Else: Tapes(3) = fullyTaped
    End If
    
    
If TextBox1.text <> vbNullString Or TextBox2.text <> vbNullString Or TextBox3.text <> vbNullString Then
    
    If Standard.rollers(True) = False Then GoTo error
    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error
    
    If Cells(ActiveCell.Row - 4, 2) = "clamping device" Then
        If Standard.clamping_device(True) = False Then GoTo error
        If Standard.rollers(False) = False Then GoTo error
    End If
    
    If Standard.line_off_marker() = False Then GoTo error
    If Standard.wind_without_feed(DefaultSpeed, InitialRotations) = False Then GoTo error
    
    
    Dim Count As Long
    For Count = 1 To 3
        If Lengths(Count) <> vbNullString Then
            If Standard.wind_with_feed(DefaultSpeed, Tapes(Count), CInt(Lengths(Count))) = False Then GoTo error
            TallyLength = TallyLength + Lengths(Count)
        End If
    Next Count
    
    If Standard.wind_without_feed(DefaultSpeed, FinalRotations) = False Then GoTo error
    
    With Range("A" & ActiveCell.Row, "E" & ActiveCell.Row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 15
    End With
End If

error:
Unload Me
End Sub

Private Sub UserForm_Activate()

ComboBox1.AddItem ("Space Taped")
ComboBox1.AddItem ("Fully Taped")
ComboBox2.AddItem ("Space Taped")
ComboBox2.AddItem ("Fully Taped")
ComboBox3.AddItem ("Space Taped")
ComboBox3.AddItem ("Fully Taped")

ComboBox1.Value = "Space Taped"
ComboBox2.Value = "Fully Taped"
ComboBox3.Value = "Space Taped"

End Sub
