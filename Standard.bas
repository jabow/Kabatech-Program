Attribute VB_Name = "Standard"
Option Explicit
Public Function cutting(number As Long) As Boolean
    If GetKeyState(45) = 1 Then ActiveCell.EntireRow.Insert Shift:=xlDown
    Cells(ActiveCell.Row, ActiveCell.Column).EntireRow.Clear
    Cells(ActiveCell.Row, 2) = "cutting"
    Cells(ActiveCell.Row, 3) = number
    ActiveCell.Offset(1, 0).Select
    cutting = True
End Function

Public Function adapter(up As Boolean) As Boolean
    If GetKeyState(45) = 1 Then ActiveCell.EntireRow.Insert Shift:=xlDown
    Cells(ActiveCell.Row, ActiveCell.Column).EntireRow.Clear
    Cells(ActiveCell.Row, 2) = "adapter"
    If up = True Then Cells(ActiveCell.Row, 3) = "up"
    If up = False Then Cells(ActiveCell.Row, 3) = "dow"
    adapter = True
    ActiveCell.Offset(1, 0).Select
End Function

Public Function wind_with_feed(speed As Long, tape As Long, length As Long) As Boolean
    If GetKeyState(45) = 1 Then ActiveCell.EntireRow.Insert Shift:=xlDown
    Cells(ActiveCell.Row, ActiveCell.Column).EntireRow.Clear
    Cells(ActiveCell.Row, 2) = "wind with feed"
    Cells(ActiveCell.Row, 3) = speed
    Cells(ActiveCell.Row, 4) = tape
    Cells(ActiveCell.Row, 5) = length
    
    With Range("E" & ActiveCell.Row).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Range("E" & ActiveCell.Row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Range("E" & ActiveCell.Row).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Range("E" & ActiveCell.Row).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    
    ActiveCell.Offset(1, 0).Select
    wind_with_feed = True
End Function

Public Function feed_without_wind(length As Long) As Boolean
    If GetKeyState(45) = 1 Then ActiveCell.EntireRow.Insert Shift:=xlDown
    Cells(ActiveCell.Row, ActiveCell.Column).EntireRow.Clear
    Cells(ActiveCell.Row, 2) = "feed w.o. wind."
    Cells(ActiveCell.Row, 3) = length
    
    With Range("C" & ActiveCell.Row).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Range("C" & ActiveCell.Row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Range("C" & ActiveCell.Row).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Range("C" & ActiveCell.Row).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    
    ActiveCell.Offset(1, 0).Select
    feed_without_wind = True
End Function

Public Function wind_without_feed(speed As Long, rotations As Long) As Boolean
    If GetKeyState(45) = 1 Then ActiveCell.EntireRow.Insert Shift:=xlDown
    Cells(ActiveCell.Row, ActiveCell.Column).EntireRow.Clear
    Cells(ActiveCell.Row, 2) = "wind w.o. feed"
    Cells(ActiveCell.Row, 3) = speed
    Cells(ActiveCell.Row, 4) = rotations
    ActiveCell.Offset(1, 0).Select
    wind_without_feed = True
End Function

Public Function rollers(up As Boolean) As Boolean
    If GetKeyState(45) = 1 Then ActiveCell.EntireRow.Insert Shift:=xlDown
    Cells(ActiveCell.Row, ActiveCell.Column).EntireRow.Clear
    Cells(ActiveCell.Row, 2) = "rollers"
    If up = True Then Cells(ActiveCell.Row, 3) = "up"
    If up = False Then Cells(ActiveCell.Row, 3) = "dow"
    ActiveCell.Offset(1, 0).Select
    rollers = True
End Function

Public Function start() As Boolean
    If GetKeyState(45) = 1 Then ActiveCell.EntireRow.Insert Shift:=xlDown
    Cells(ActiveCell.Row, ActiveCell.Column).EntireRow.Clear
    Cells(ActiveCell.Row, 2) = "start"
    ActiveCell.Offset(1, 0).Select
    start = True
End Function

Public Function hood_open() As Boolean
    If GetKeyState(45) = 1 Then ActiveCell.EntireRow.Insert Shift:=xlDown
    Cells(ActiveCell.Row, ActiveCell.Column).EntireRow.Clear
    Cells(ActiveCell.Row, 2) = "hood open"
    ActiveCell.Offset(1, 0).Select
    hood_open = True
End Function

Public Function position(distance As Long) As Boolean
    If GetKeyState(45) = 1 Then ActiveCell.EntireRow.Insert Shift:=xlDown
    Cells(ActiveCell.Row, ActiveCell.Column).EntireRow.Clear
    Cells(ActiveCell.Row, 2) = "position"
    Cells(ActiveCell.Row, 3) = distance
    ActiveCell.Offset(1, 0).Select
    position = True
End Function

Public Function line_off_marker() As Boolean
    If GetKeyState(45) = 1 Then ActiveCell.EntireRow.Insert Shift:=xlDown
    Cells(ActiveCell.Row, ActiveCell.Column).EntireRow.Clear
    Cells(ActiveCell.Row, 2) = "line off marker"
    ActiveCell.Offset(1, 0).Select
    line_off_marker = True
End Function

Public Function clamping_device(up As Boolean) As Boolean
    If GetKeyState(45) = 1 Then ActiveCell.EntireRow.Insert Shift:=xlDown
    Cells(ActiveCell.Row, ActiveCell.Column).EntireRow.Clear
    Cells(ActiveCell.Row, 2) = "clamping device"
    If up = True Then Cells(ActiveCell.Row, 3) = "up"
    If up = False Then Cells(ActiveCell.Row, 3) = "dow"
    ActiveCell.Offset(1, 0).Select
    clamping_device = True
End Function
