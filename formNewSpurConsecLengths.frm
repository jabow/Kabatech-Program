VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formNewSpurConsecLengths 
   Caption         =   "New Spur with Consecutive Lengths"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "formNewSpurConsecLengths.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formNewSpurConsecLengths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdOK_Click()

Dim PreviousRow As String
Dim tape As Long


    Cells(ActiveCell.Row - 1, 2).Select
    PreviousRow = ActiveCell.Value
    Cells(ActiveCell.Row + 1, 2).Select
    TotalLength = TallyLength
        
    If Not PreviousRow = "clamping device" Then
        If Not (TotalLength < 300) And (AutoAdapter = True) Then
            If Standard.cutting(1) = False Then GoTo error
        End If
    End If

    If Standard.rollers(True) = False Then GoTo error

    If (TotalLength < 300) And (AutoAdapter = True) Then
        If Standard.adapter(True) = False Then GoTo error
    End If

    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error

    If PreviousRow = "clamping device" Then
        If Standard.rollers(False) = False Then GoTo error
        If Standard.clamping_device(True) = False Then GoTo error
    End If

    If Standard.line_off_marker() = False Then GoTo error
    If Standard.position(10) = False Then GoTo error
    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error
    If Standard.wind_without_feed(DefaultSpeed, InitialRotations) = False Then GoTo error
    
    Cells(ActiveCell.Row, 5).Select

        ActiveCell.Value = TextBox1.Value
        If ActiveCell.Value >= NewLengthAdjustment Then
            TallyLength = ActiveCell.Value
            ActiveCell.Value = ActiveCell.Value - NewLengthAdjustment
            x = 2
        Else
           x = MsgBox("Value too low for machine to reproduce.  Must be at least " & NewLengthAdjustment & "mm!", vbRetryCancel)
           Exit Sub
        End If

    
    If ComboBox1.Value = "Space Taped" Then
        tape = SpaceTaped
    Else: tape = fullyTaped
    End If
    
    If Standard.wind_with_feed(DefaultSpeed, tape, ActiveCell.Value) = False Then GoTo error

    Cells(ActiveCell.Row, 5).Select
    ActiveCell.Value = TextBox2.Value
    TallyLength = ActiveCell.Value
    ActiveCell.Value = ActiveCell.Value
    
    If ComboBox2.Value = "Space Taped" Then
        tape = SpaceTaped
    Else: tape = fullyTaped
    End If
    
    If Standard.wind_with_feed(DefaultSpeed, tape, ActiveCell.Value) = False Then GoTo error

    If Standard.wind_without_feed(DefaultSpeed, FinalRotations) = False Then GoTo error
    
    With Range("A" & ActiveCell.Row, "E" & ActiveCell.Row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 15
    End With
error:
Unload Me
End Sub

Private Sub UserForm_Activate()

ComboBox1.AddItem ("Space Taped")
ComboBox1.AddItem ("Fully Taped")
ComboBox2.AddItem ("Space Taped")
ComboBox2.AddItem ("Fully Taped")

ComboBox1.Value = "Space Taped"
ComboBox2.Value = "Fully Taped"


End Sub
