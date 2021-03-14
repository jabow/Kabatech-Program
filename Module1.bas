Attribute VB_Name = "Module1"
Option Explicit
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Default Settings

Public Const DefaultSpeed As Long = 400              'Tape winder speed
Public Const SpaceTaped As Long = 40                 'Tape spacings
Public Const fullyTaped As Long = 18                 'Tape spacings
Public Const NewLengthAdjustment As Long = 110        'Amount taken off each new length
Public Const MinKabatecLength As Long = NewLengthAdjustment + 10 'shortest length possible due to machine limitations
Public Const MinCuttingLength = 100 + NewLengthAdjustment 'the machine will only automatically cut a new spur at 100
Public Const pos As Long = 10                        'standard position
Public Const InitialRotations As Long = 5            'Number of tape winds before starting feed
Public Const FinalRotations As Long = 6              'Number of extra tape winds when finishing feed
Public Const InitPos As String = "A7"                   'Intitial Cell
Public Const changecolor As Long = 16776961             'colour of changed text
Public Const NumberOfProperties As Long = 7          'Number of properties at the top of input sheet
Public Const StartRow As Long = NumberOfProperties + 3                   'Start point of input sheet
Public Const NumberOFBranchDetails As Long = 6       'Number details about a branch
Public Const inputSheet As String = "Input"             'Name of input sheet
Public Const ProgramSheet As String = "Program"            'Name of program sheet
Public Const WordCyan As Long = 3                    'Cyan color for word
'Public Const InputSheetFilePath = "C:\Users\" & Environ("username") & "\Dropbox (BCA_TECH)\Kabatech Program\Input Sheets\" '****WORKPATH
Public Const BuildsheetFilePath = "T:\Technical\Kabatech\Build Sheets\"  'WORKPATH
Public Const ProgramFilePath = "T:\Technical\Kabatech\Programs\" 'WORKPATH
'Public Const InputSheetFilePath = "C:\Users\James\Dropbox (BCA_TECH)\James\Excel Programs\C16 Program James\Example File Structure\Input Sheet\" '****HOMEPATH
'Public Const BuildsheetFilePath = "C:\Users\James\Dropbox (BCA_TECH)\James\Excel Programs\C16 Program James\Example File Structure\Build sheets\"  'HOMEPATH
'Public Const ProgramFilePath = "C:\Users\James\Dropbox (BCA_TECH)\James\Excel Programs\C16 Program James\Example File Structure\Programs\" 'HOMEPATH
Public Const KabatecLenLimitLow As Long = 300        'Lowest number of kabatech limitation when clamps come down on the arm
Public Const KabatecLenLimitHigh As Long = 390       'Highest number of kabatech limitation when clamps come down on the arm
Public Const MBH As String = "MEASURE BY HAND"          'Measure by hand
Public Const DP As String = "DROPOUT"      'Dropout



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public AutoAdapter As Boolean
Public TotalLength As Long
Public TallyLength As Long
Public length As Long
Public position As Long
Public ProgramNum As String
Public fileName As String
Public importFlag As Boolean 'Used to check if its ok to check the part number (not ok when importing and clearing input sheet)



Sub check_user_spur_length()
    Dim x As Long
    Do
        'If length >= NewLengthAdjustment And is_length_within_kabatech_limtations(length) = False Then
        If length < MinKabatecLength Then
            If MsgBox("Value too low for machine to reproduce.  Must be at least " & MinKabatecLength & "mm!", vbRetryCancel) = vbCancel Then
                'If x = vbCancel Then check_user_spur_length = 0
                ActiveCell.Value = vbNullString
                End
            Else
                formLength.Show
            End If
        ElseIf length > KabatecLenLimitLow And length < KabatecLenLimitHigh Then
            If MsgBox("Length cannot be run on kabatech, cannot be between " & KabatecLenLimitLow & " and " & KabatecLenLimitHigh, vbRetryCancel) = vbCancel Then
                'x = vbCancel Then check_user_spur_length = 0
                ActiveCell.Value = vbNullString
                End
            Else
                formLength.Show
            End If
'            length = 390
'            MsgBox "Length is between " & KabatecLenLimitLow & " and " & KabatecLenLimitHigh & ". Length adjusted to 390"
        Else
            TotalLength = TallyLength
            TallyLength = length
            x = 2
        End If
    Loop Until x = 2
    'check_user_spur_length = length
End Sub

Public Function GetPressed(ByVal control As IRibbonControl, ByRef returnedVal) As Boolean
    GetPressed = False
    AutoAdapter = True
    returnedVal = True
End Function



Public Sub Auto_Adapter(control As IRibbonControl, pressed As Boolean)
If pressed = True Then
    AutoAdapter = True
Else
    AutoAdapter = False
    MsgBox ("The 'Adapter' And 'Cutting' lines of code will no longer be inserted automatically!")
End If
End Sub

Sub clear_adapter_count(control As IRibbonControl)

If MsgBox("This will cancel all current calculations for the automatic adapter setting.  Are you sure you want to continue?", vbQuestion + vbYesNo, "???") = vbYes Then
    TallyLength = 0
    TotalLength = 0
End If

End Sub


Sub Consecutive_Taped_Lengths(control As IRibbonControl)
formConsecLengths.Show
End Sub




Sub Start_New_Spur_Fully_Taped(Optional control As IRibbonControl, Optional btLength As Long)
Attribute Start_New_Spur_Fully_Taped.VB_ProcData.VB_Invoke_Func = "N\n14"
    
'
' Start_New_Spur_Fully_Taped Macro
' Inserts code For a fully taped new spur
'
Dim PreviousRow As String
    
    
    If length = 0 Then 'if the length is set (comes from old buttons)
        formLength.Show
        length = ActiveCell.Value 'if length is not set set it
        check_user_spur_length 'check the length
        position = pos 'set the position
    Else
        TotalLength = TallyLength
        TallyLength = length
    End If

    Cells(ActiveCell.Row - 1, 2).Select
    PreviousRow = ActiveCell.Value
    Cells(ActiveCell.Row + 1, 2).Select
        
    'If Not PreviousRow = "clamping device" Then
        'If Not (TotalLength < 300) And (AutoAdapter = True) Then
            If Standard.cutting(1) = False Then GoTo error
        'End If
    'End If

    If Standard.rollers(True) = False Then GoTo error

    If (TotalLength < 700) And (AutoAdapter = True) Then
        If Standard.adapter(True) = False Then GoTo error
    End If

    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error

    If PreviousRow = "clamping device" Then
        If Standard.rollers(False) = False Then GoTo error
        If Standard.clamping_device(True) = False Then GoTo error
    End If
    If Standard.line_off_marker() = False Then GoTo error
    If Standard.position(position) = False Then GoTo error
    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error
    If Standard.wind_without_feed(DefaultSpeed, InitialRotations) = False Then GoTo error
    
    If Standard.wind_with_feed(DefaultSpeed, fullyTaped, length - NewLengthAdjustment) = False Then GoTo error
    'If there is blue tape add additional wind with feed comman
    If btLength > 0 Then
        If Standard.wind_with_feed(DefaultSpeed, SpaceTaped, btLength) = False Then GoTo error
    End If

    If Standard.wind_without_feed(DefaultSpeed, FinalRotations) = False Then GoTo error
    
    With Range("A" & ActiveCell.Row, "E" & ActiveCell.Row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 15
    End With
    length = 0
error:
End Sub


Sub Start_New_Spur_Space_Taped(Optional control As IRibbonControl, Optional btLength As Long)
Attribute Start_New_Spur_Space_Taped.VB_ProcData.VB_Invoke_Func = "n\n14"
'
' Start_New_Spur_Space_Taped Macro
' Inserts code to start a new spur that is space taped.
'
'If Not check_user_spur_length Then Exit Sub
Dim PreviousRow As String

    If length = 0 Then 'if the length is set (comes from old buttons)
        formLength.Show
        length = ActiveCell.Value 'if length is not set set it
        check_user_spur_length 'check the length
        position = pos 'set the position
    Else
        TotalLength = TallyLength
        TallyLength = length
    End If

    Cells(ActiveCell.Row - 1, 2).Select
    PreviousRow = ActiveCell.Value
    Cells(ActiveCell.Row + 1, 2).Select
        
    'Removed as Kabatecs can cut at any length
    'If Not PreviousRow = "clamping device" Then
       ' If Not (TotalLength < 300) And (AutoAdapter = True) Then
            If Standard.cutting(1) = False Then GoTo error
        'End If
    'End If

    If Standard.rollers(True) = False Then GoTo error

    If (TotalLength < 700) And (AutoAdapter = True) Then
        If Standard.adapter(True) = False Then GoTo error
    End If

    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error

    If PreviousRow = "clamping device" Then
        If Standard.rollers(False) = False Then GoTo error
        If Standard.clamping_device(True) = False Then GoTo error
    End If

    If Standard.line_off_marker() = False Then GoTo error
    If Standard.position(position) = False Then GoTo error
    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error
    If Standard.wind_without_feed(DefaultSpeed, InitialRotations) = False Then GoTo error
    
    If Standard.wind_with_feed(DefaultSpeed, SpaceTaped, length - NewLengthAdjustment) = False Then GoTo error
    'If there is blue tape add additional wind with feed comman
    If btLength > 0 Then
        If Standard.wind_with_feed(DefaultSpeed, fullyTaped, btLength) = False Then GoTo error
    End If

    If Standard.wind_without_feed(DefaultSpeed, FinalRotations) = False Then GoTo error
    
    With Range("A" & ActiveCell.Row, "E" & ActiveCell.Row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 15
    End With
    length = 0
error:
End Sub







Sub Insert_Taped_Length(Optional control As IRibbonControl, Optional btLength As Long)
Attribute Insert_Taped_Length.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' Insert_Taped_Length Macro
' Inserts code for a fully taped length
'
Dim PreviousRow As String

    If length = 0 Then
        formLength.Show
        length = ActiveCell.Value
    End If
    If length = 0 Then Exit Sub
    
    TallyLength = TallyLength + length

    Cells(ActiveCell.Row - 1, 2).Select
    PreviousRow = ActiveCell.Value
    Cells(ActiveCell.Row + 1, 2).Select

    If Standard.rollers(True) = False Then GoTo error
    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error

    If PreviousRow = "clamping device" Then
        If Standard.rollers(False) = False Then GoTo error
        If Standard.clamping_device(True) = False Then GoTo error
    End If

    If Standard.line_off_marker() = False Then GoTo error
    If Standard.wind_without_feed(DefaultSpeed, InitialRotations) = False Then GoTo error

    If Standard.wind_with_feed(DefaultSpeed, fullyTaped, length) = False Then GoTo error
    'If there is blue tape add additional wind with feed comman
    If btLength > 0 Then
        If Standard.wind_with_feed(DefaultSpeed, SpaceTaped, btLength) = False Then GoTo error
    End If
    If Standard.wind_without_feed(DefaultSpeed, FinalRotations) = False Then GoTo error
  
    With Range("A" & ActiveCell.Row, "E" & ActiveCell.Row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 15
    End With
    length = 0
error:
End Sub



Sub Insert_Space_Taped_Length(Optional control As IRibbonControl, Optional btLength As Long)
'
' Insert_Taped_Length Macro
' Inserts code for a fully taped length
'
Dim PreviousRow As String

    
    If length = 0 Then
        formLength.Show
        length = ActiveCell.Value
    End If
    If length = 0 Then Exit Sub
    
    TallyLength = TallyLength + length

    Cells(ActiveCell.Row - 1, 2).Select
    PreviousRow = ActiveCell.Value
    Cells(ActiveCell.Row + 1, 2).Select

    If Standard.rollers(True) = False Then GoTo error
    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error

    If PreviousRow = "clamping device" Then
        If Standard.rollers(False) = False Then GoTo error
        If Standard.clamping_device(True) = False Then GoTo error
    End If

    If Standard.line_off_marker() = False Then GoTo error
    If Standard.wind_without_feed(DefaultSpeed, InitialRotations) = False Then GoTo error

    If Standard.wind_with_feed(DefaultSpeed, SpaceTaped, length) = False Then GoTo error
    'If there is blue tape add additional wind with feed comman
    If btLength > 0 Then
        If Standard.wind_with_feed(DefaultSpeed, fullyTaped, btLength) = False Then GoTo error
    End If
    If Standard.wind_without_feed(DefaultSpeed, FinalRotations) = False Then GoTo error
  
    With Range("A" & ActiveCell.Row, "E" & ActiveCell.Row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 15
    End With
    length = 0
error:
End Sub






Sub Insert_Initial_Spur_Fully_Taped(Optional control As IRibbonControl, Optional btLength As Long)
Attribute Insert_Initial_Spur_Fully_Taped.VB_ProcData.VB_Invoke_Func = "I\n14"
'
' Insert_Initial_Spur_Fully_Taped Macro
' Inserts code for the first spur in the harness
'
'If Not check_user_spur_length Then Exit Sub
    
'    If length = 0 Then
'        formLength.Show
'        length = ActiveCell.Value 'if length is not set set it
'    End If
'    check_user_spur_length
'    If position = 0 Then position = Pos   'if postion is not set set it
    
    If length = 0 Then 'if the length is set (comes from old buttons)
        formLength.Show
        length = ActiveCell.Value 'if length is not set set it
        check_user_spur_length 'check the length
        position = pos 'set the position
    Else
        TotalLength = TallyLength
        TallyLength = length
    End If
    
    
    
    If Standard.position(position) = False Then GoTo error
    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error
    If Standard.wind_without_feed(DefaultSpeed, InitialRotations) = False Then GoTo error
    
    If Standard.wind_with_feed(DefaultSpeed, fullyTaped, length - NewLengthAdjustment) = False Then GoTo error
    'If there is blue tape add additional wind with feed comman
    If btLength > 0 Then
        If Standard.wind_with_feed(DefaultSpeed, SpaceTaped, btLength) = False Then GoTo error
    End If
    
    If Standard.wind_without_feed(DefaultSpeed, FinalRotations) = False Then GoTo error
    With Range("A" & ActiveCell.Row, "E" & ActiveCell.Row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 15
    End With
    length = 0
error:
End Sub





Sub Insert_Initial_Spur_Space_Taped(Optional control As IRibbonControl, Optional btLength As Long)
Attribute Insert_Initial_Spur_Space_Taped.VB_ProcData.VB_Invoke_Func = "i\n14"
'
' Insert_Initial_Spur_Space_Taped Macro
' Inserts code for the first spur of the harness with spaced tape
'
'If Not check_user_spur_length Then Exit Sub
    If length = 0 Then 'if the length is set (comes from old buttons)
        formLength.Show
        length = ActiveCell.Value 'if length is not set set it
        check_user_spur_length 'check the length
        position = pos 'set the position
    Else
        TotalLength = TallyLength
        TallyLength = length
    End If
    
    If Standard.position(position) = False Then GoTo error
    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error
    If Standard.wind_without_feed(DefaultSpeed, InitialRotations) = False Then GoTo error
    
    If Standard.wind_with_feed(DefaultSpeed, SpaceTaped, length - NewLengthAdjustment) = False Then GoTo error
    'If there is blue tape add additional wind with feed comman
    If btLength > 0 Then
        If Standard.wind_with_feed(DefaultSpeed, fullyTaped, btLength) = False Then GoTo error
    End If
    
    If Standard.wind_without_feed(DefaultSpeed, FinalRotations) = False Then GoTo error
    With Range("A" & ActiveCell.Row, "E" & ActiveCell.Row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 15
    End With
    length = 0
error:
End Sub




Sub Insert_Final_Cut(Optional control As IRibbonControl)
Attribute Insert_Final_Cut.VB_ProcData.VB_Invoke_Func = "f\n14"
'
' Insert_Final_Cut Macro
' Inserts the code to finish the harness by cutting and releasing from the C16 machine.  This is required at the end of all programs.
'
' Keyboard Shortcut: Ctrl+f
    If Standard.cutting(1) = False Then GoTo error
    If Standard.rollers(True) = False Then GoTo error
    If Standard.adapter(True) = False Then GoTo error
    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error
    
    With Range("A" & ActiveCell.Row, "E" & ActiveCell.Row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 15
    End With
error:
End Sub








Sub Insert_Feed_Without_Wind(Optional control As IRibbonControl)
Attribute Insert_Feed_Without_Wind.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' Insert_Feed_Without_Wind Macro
' Inserts the code for a feed without wind.
'
' Keyboard Shortcut: Ctrl+Shift+F
'
Dim PreviousRow As String

    If length = 0 Then 'if the length is set (comes from old buttons)
        formLength.Show
        length = ActiveCell.Value 'if length is not set set it
        check_user_spur_length 'check the length
        position = pos 'set the position
    Else
        TotalLength = TallyLength
        TallyLength = length
    End If


    Cells(ActiveCell.Row - 1, 2).Select
    PreviousRow = ActiveCell.Value
    Cells(ActiveCell.Row + 1, 2).Select

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
    If Standard.line_off_marker() = False Then GoTo error
    If Standard.position(position) = False Then GoTo error
    If Standard.hood_open() = False Then GoTo error
    If Standard.start() = False Then GoTo error

    If PreviousRow = "clamping device" Then
        If Standard.clamping_device(True) = False Then GoTo error
    End If

'    Cells(ActiveCell.Row, 3).Select
'    Do
'        formLength.Show
'        If ActiveCell.Value >= NewLengthAdjustment Then
'            TallyLength = ActiveCell.Value
'            ActiveCell.Value = ActiveCell.Value - NewLengthAdjustment
'            X = 2
'        Else
'           X = MsgBox("Value too low for machine to reproduce.  Must be at least " & NewLengthAdjustment & "mm!", vbRetryCancel)
'        End If
'    Loop Until X = 2
    If Standard.feed_without_wind(length - NewLengthAdjustment) = False Then GoTo error

    If Standard.clamping_device(False) = False Then GoTo error
    
    With Range("A" & ActiveCell.Row, "E" & ActiveCell.Row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 15
    End With
error:
End Sub


'Callback for button14 onAction
Sub New_Spur_Tape_Change(control As IRibbonControl)

formNewSpurConsecLengths.Show

End Sub




Sub Insert_Adapter_Line(Optional control As IRibbonControl)
Attribute Insert_Adapter_Line.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' Insert_Adapter_Line Macro
' Inserts Adapter commands to be used if overall length of runs is below 300mm.  Should be inserted after rollers UP command.
'
' Keyboard Shortcut: Ctrl+Shift+A
'
    If Standard.adapter(True) = False Then GoTo error
error:
End Sub






Sub Import_CLC_To_Excel(Optional control As IRibbonControl)
Attribute Import_CLC_To_Excel.VB_ProcData.VB_Invoke_Func = "O\n14"
'
' Import_CLC_To_Excel Macro
' Imports a .CLC file into the excel template for easy changing.
'
' Keyboard Shortcut: Ctrl+Shift+O
'
    Sheets(ProgramSheet).Activate
    Range("a6").Value = 0
    Dim StrFind As String
    Dim LastRow As Long
    Dim fileName As Variant
    Dim iCount As Long
    ChDrive ("U:\\")
    ChDir ("U:\Tech\C16Programs")
    fileName = Application.GetOpenFilename(filefilter:="Text File (*.clc),(*.clc")
    If fileName = False Then
        Exit Sub
    End If

    
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
            
    Range("a6", "f" & LastRow).Clear


    Range("A6").Select
    With Sheets(ProgramSheet).QueryTables.Add( _
        Connection:="TEXT;" & fileName, _
        Destination:=Range("A6"))
        
        .Name = "TEST"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = xlMSDOS
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2, 9, 2, 9, 2, 9, 2, 9, 2)
        .TextFileFixedColumnWidths = Array(3, 4, 15, 4, 3, 4, 2, 4)
        .Refresh BackgroundQuery:=False
    End With
    
    With Sheets(ProgramSheet)
        LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    End With
    
        Range("B4").Value = LastRow
        
        
        Do Until Left$(StrFind, 1) = "\"

            iCount = iCount + 1

            StrFind = Right$(fileName, iCount)

            If iCount = Len(fileName) Then Exit Do

        Loop

        Range("B1").Value = Right$(StrFind, Len(StrFind) - 1)
        Sheets(ProgramSheet).Range("b2").ClearContents ' note that this row expects the worksheet to be named DATA
        Sheets(ProgramSheet).Hyperlinks.Add Range("b2"), fileName

End Sub











Sub Export_As_CLC(ByRef program As clsProgram)
Attribute Export_As_CLC.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' Export_As_C16 Macro
' Exports Spreadsheet as a C16 file
'
' Keyboard Shortcut: Ctrl+a
'

Dim MyStr As String
Dim File As Variant
Dim OldFile As Variant
Dim FirstRow As Long
Dim LastRow As Long
Dim MyRow As Long
Dim Savetofloppy As Boolean

Savetofloppy = False

'CODE FOR FORMATTING SPREADSHEET CORRECTLY




    With Sheets(ProgramSheet)
        FirstRow = .Range("b3").Value ' the range of the table to be exported
        LastRow = .Cells(Rows.Count, "B").End(xlUp).Row
        .Range("b4").Value = LastRow
        LastRow = .Range("b4").Value
        
        .Cells(FirstRow, 1).Value = "01"
        For MyRow = FirstRow + 1 To LastRow
            .Cells(MyRow, 1).Value = Format$(MyRow - FirstRow, "000")
        Next
    End With



'CODE FOR CREATING FILES TO SAVE TO


    File = ProgramFilePath & program.Program_no & ".clc"
    OldFile = ProgramFilePath & program.Part_no & "(" & program.Revision - 1 & ").clc"
    Dim str As String
    str = Dir(OldFile)
    If str <> vbNullString Then
        program.ProgOverwritten = True
    Else: program.ProgOverwritten = False
    End If
    If File = False Then
        MsgBox ("An Error Occurred and the file was not saved.  This could be due to an incorrect filename or path, please retry insuring these inputs are correct.")
        Exit Sub
    End If

'If MsgBox("Would you like to save to the floppy drive A:\\ as well?" & Chr(13) & "Insure a floppy disk is in the drive before clicking Yes!", vbYesNo + vbExclamation, "Save to Floppy?") = vbYes Then
 '   On Error GoTo FloppyProblems
 '   ChDrive ("A:\\")
'    ChDir ("A:\\")
'    File2 = Application.GetSaveAsFilename("Default.clc", "C16 Program (*.clc), *.clc")
'    Savetofloppy = True
'    Open File2 For Output As #2
'End If



'CODE FOR INPUTTING DATE TO FILES
Open File For Output As #1

Worksheets(ProgramSheet).Activate
For MyRow = FirstRow To LastRow ' loop through each row of the table
 MyStr = vbNullString
 If Not Cells(MyRow, 1).Value = vbNullString Then
    If Cells(MyRow, 2).Value = vbNullString Then
        MyStr = Cells(MyRow, 1).Value & String(3 - Len(Cells(MyRow, 1).Value), " ")
    Else
        MyStr = Cells(MyRow, 1).Value & String(3 - Len(Cells(MyRow, 1).Value), " ") & "  ³ "
    End If
    If MyRow = 6 Then
        MyStr = MyStr & Cells(MyRow, 2).Value & String(15 - Len(Cells(MyRow, 2).Value), " ")
    Else
        MyStr = MyStr & Cells(MyRow, 2).Value & String(15 - Len(Cells(MyRow, 2).Value), " ") & "  ³ "
    End If
    If Cells(MyRow, 4).Value = vbNullString Then
        MyStr = MyStr & Cells(MyRow, 3).Value & String(4 - Len(Cells(MyRow, 3).Value), " ")
    Else
        MyStr = MyStr & Cells(MyRow, 3).Value & String(3 - Len(Cells(MyRow, 3).Value), " ") & "  ³ "
    End If
    If Cells(MyRow, 5).Value = vbNullString Then
        MyStr = MyStr & Cells(MyRow, 4).Value & String(2 - Len(Cells(MyRow, 4).Value), " ")
    Else
        MyStr = MyStr & Cells(MyRow, 4).Value & String(2 - Len(Cells(MyRow, 4).Value), " ") & "  ³ "
    End If
    If Cells(MyRow, 6).Value = vbNullString Then
        MyStr = MyStr & Cells(MyRow, 5).Value & String(5 - Len(Cells(MyRow, 5).Value), " ")
    Else
        MyStr = MyStr & Cells(MyRow, 5).Value & String(5 - Len(Cells(MyRow, 5).Value), " ") & "  ³ "
    End If
End If
    
Print #1, MyStr

If Savetofloppy = True Then
    Print #2, MyStr
End If

Next
Close #1

If Savetofloppy = True Then
    Close #2
End If
Dim DateTimeNow As String: DateTimeNow = Format(Now(), "DD-MM-YYYY hh.mm.ss")
program.ProgSaved = DateTimeNow
SetAttr File, vbReadOnly 'set the file as read only
If program.Revision > 0 Then DeleteFile (OldFile)
Sheets(ProgramSheet).Range("b2").ClearContents ' note that this row expects the worksheet to be named DATA
Sheets(ProgramSheet).Hyperlinks.Add Sheets(ProgramSheet).Range("b2"), File
Exit Sub


End Sub

Sub Clear_Cells_of_Formatting_Program(Optional control As IRibbonControl)
Attribute Clear_Cells_of_Formatting_Program.VB_ProcData.VB_Invoke_Func = "X\n14"
    With Sheets(ProgramSheet).Select 'clear the program sheet to start with
        Rows(7 & ":" & Rows.Count).Delete 'clear anything currently there
        Range(InitPos).Activate  'set the initial position
    End With
End Sub

Sub Clear_Cells_of_Formatting_Input(Optional control As IRibbonControl)
    importFlag = True
    With Sheets(inputSheet)
        .Cells(1, 7).ClearContents  'clear wires to mark
        .Range(Cells(1, 2), Cells(NumberOfProperties, 2)).ClearContents 'clear properties
        .Range(Cells(StartRow, 1), Cells(Rows.Count, NumberOFBranchDetails)).ClearContents 'clear branches
        .Range(Cells(StartRow, 1), Cells(Rows.Count, NumberOFBranchDetails)).Interior.ColorIndex = 0
        '.Protect AllowInsertingRows:=True, AllowFormattingCells:=True
    End With
    importFlag = False
End Sub


Sub Clear_Selected_Cells_of_Formatting(Optional control As IRibbonControl)
    importFlag = True
    Dim Ret As Range
    On Error Resume Next
    Set Ret = Application.InputBox("Please select the Cells", "Clear Cells", Type:=8)
    On Error GoTo error
    If Not Ret Is Nothing Then
        Ret.ClearContents
        Ret.Interior.ColorIndex = 0
    End If
    Exit Sub
error:  MsgBox "All or part of the selected area is protected, please just select the cells that need clearing"
    importFlag = False
End Sub

Sub DeleteRow(Optional control As IRibbonControl)
    Dim Ret As Range ', Cl As range
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsActiveSheet As Worksheet
    Set wsActiveSheet = wb.ActiveSheet

    On Error Resume Next
    Set Ret = Application.InputBox("Please select the Cells", "Delete Rows", Type:=8)
    On Error GoTo 0

    wsActiveSheet.Unprotect

    If Not Ret Is Nothing Then Ret.EntireRow.Delete

    wsActiveSheet.Protect AllowInsertingRows:=True, AllowFormattingCells:=True
End Sub

Sub Write_Dropout(Optional IRibbonControl)
    If Selection.Cells.Count = 1 Then
        If ActiveCell.Column = 2 And ActiveCell.Row > 10 Then
            Selection.Value = DP
        Else
            MsgBox "Can only write in the description cell"
        End If
    Else
        MsgBox "Select only one cell"
    End If
End Sub

Sub Write_Measure_By_Hand(Optional IRibbonControl)
    If Selection.Cells.Count = 1 Then
        If ActiveCell.Column = 2 And ActiveCell.Row > 10 Then
            Selection.Value = MBH
        Else
            MsgBox "Can only write in the description cell"
        End If
    Else
        MsgBox "Select only one cell"
    End If
End Sub


Sub change_cell_text_colour(Optional control As IRibbonControl)
    If Selection.Row > StartRow - 1 And Selection.Column < 4 And Selection.Locked = False Then
        If Selection.Interior.Color = changecolor Then
            Selection.Interior.ColorIndex = 0 'Change to no fill
        Else
            Selection.Interior.Color = changecolor 'change cell color
        End If
    Else
        MsgBox "Cant highlight this cell"
    End If
End Sub

Sub Export_Spreadsheet_To_Text(ByRef program As clsProgram)
    Dim filePath As String
    Dim revFPath As String
    Dim CellData As String
    Dim export As Boolean

    Worksheets(inputSheet).Activate

    
    filePath = "C:\Users\" & Environ("username") & "\Dropbox (BCA_TECH)\Kabatech Program\Input Sheets\" & program.Part_no 'new file

    If program.Revision > 0 Then 'what to do with old file
        revFPath = "C:\Users\" & Environ("username") & "\Dropbox (BCA_TECH)\Kabatech Program\Input Sheets\" & "Archive\" & program.Part_no & "\" & program.Part_no & "(" & program.Revision - 1 & ")"
        Name filePath As revFPath 'Move old file to archive folder and rename
        program.InputSheetOverwritten = True
    Else
        program.InputSheetOverwritten = False
    End If
    
'    Dim TestStr As String
'    On Error Resume Next
'    TestStr = Dir(filePath)
'    On Error GoTo 0
'    If TestStr <> vbNullString Then 'if true then file already exists
'        Dim NewFPath As String
'        NewFPath = InputSheetFilePath & "Archive\" & program.Program_no & "\"
'        Dim FSO As Object
'        Set FSO = CreateObject("scripting.filesystemobject")
'        If FSO.FolderExists(NewFPath) = False Then
'            MkDir (NewFPath)
'        End If
'
'        Dim RevFPath As String
'        Dim version As Integer
'        version = 0
'        RevFPath = NewFPath & "rev-" & program.Revision & "(" & version & ")" 'examples - rev0(0) - rev0(1) - rev1(0)
'        TestStr = Dir(RevFPath)
'        While TestStr <> vbNullString 'if true then file already exists
'            version = version + 1
'            RevFPath = NewFPath & "rev-" & program.Revision & "(" & version & ")"
'            TestStr = Dir(RevFPath)
'        Wend
'        Name filePath As RevFPath 'Move old file to archive folder and rename
'        program.InputSheetOverwritten = True
'
'    Else: program.InputSheetOverwritten = False
'    End If
    
    
    
    Open filePath For Output As #2

    CellData = Cells(1, 7) 'wires to mark
    Print #2, CellData
    Dim colourChange As Boolean
    Dim i As Long
    For i = 1 To NumberOfProperties
        CellData = Cells(i, 2)
        Print #2, CellData
    Next
    Dim a As Long
    a = StartRow
    Do While end_case(a) <> True
        Dim x As Long
        For x = 1 To NumberOFBranchDetails
            colourChange = Cells(a, x).Interior.Color = changecolor
            CellData = Cells(a, x)
            Print #2, colourChange
            Print #2, CellData
        Next
        a = a + 1
    Loop
    Close #2
    Dim DateTimeNow As String: DateTimeNow = Format(Now(), "DD-MM-YYYY hh.mm.ss")
    program.InputSheetSaved = DateTimeNow
    SetAttr filePath, vbReadOnly 'set the file as read only
End Sub

Sub Import_Text_To_Spreadsheet(Optional control As IRibbonControl, Optional part_number As String)
    Dim filePath As String
    Dim textline As String
    Dim textline2 As String
    Dim y As Long
    
    Dim old As Boolean
    If part_number = "" Then
        formImport.Show
        old = formImport.Tag = True
    Else
        ProgramNum = part_number
    End If
    
    If old Then
        ChDrive ("C:\\")
        ChDir ("C:\Users\" & Environ("username") & "\Dropbox (BCA_TECH)\Kabatech Program\Input Sheets\Archive")
        filePath = Application.GetOpenFilename
    End If

   
    'If RevisionNum = vbNullString Then RevisionNum = 0

'    If RevisionNum = "#" Then 'Find the newest revision
'        RevisionNum = 0
'        FilePath = InputSheetFilePath & ProgramNum & RevisionNum
'        Do While Test_File_Exist_With_Dir(FilePath)
'            RevisionNum = RevisionNum + 1
'            FilePath = InputSheetFilePath & ProgramNum & RevisionNum
'        Loop
'        RevisionNum = RevisionNum - 1
'    End If
    If ProgramNum <> vbNullString Or filePath <> vbNullString Then
        If filePath = vbNullString Then filePath = "C:\Users\" & Environ("username") & "\Dropbox (BCA_TECH)\Kabatech Program\Input Sheets\" & ProgramNum
        If Test_File_Exist_With_Dir(filePath) Then
        
            Open filePath For Input As #1
            
            Worksheets(inputSheet).Activate
            Clear_Cells_of_Formatting_Input 'clear input sheet
            importFlag = True
            Line Input #1, textline 'read the first line
            Cells(1, 7) = textline
            y = StartRow
            Dim i As Long
            
            Dim createdStamp As Date
            Dim definedStamp As Date
            createdStamp = FileDateTime(filePath)
            definedStamp = "8/12/20 14:20:00 am" 'anything created before then
            
            For i = 1 To NumberOfProperties
                If i = 5 And createdStamp < definedStamp Or i = 7 And createdStamp < definedStamp Then 'some older versions have been exported with revision and program numbers
                    Line Input #1, textline
                End If
                Line Input #1, textline
                Cells(i, 2) = textline
            Next
            Do Until EOF(1)
                Dim x As Long
                For x = 1 To NumberOFBranchDetails
                    Line Input #1, textline
                    If textline <> vbNullString Then
                        If textline = "True" Or textline = "False" Then
                            Line Input #1, textline2
                            If textline = "True" Then
                                Cells(y, x).Interior.Color = changecolor
                            ElseIf textline = "False" Then
                                Cells(y, x).Interior.ColorIndex = 0
                            End If
                            Cells(y, x) = textline2 'Application.WorksheetFunction.Round(x / colChg, 0)
                        Else
                            Cells(y, x) = textline
                        End If
                    End If
                Next
                y = y + 1
            Loop
            Close #1
        Else
            MsgBox ("Doesnt Exist")
        End If
    Else
        MsgBox ("No Program Number Entered")
    End If
    importFlag = False
    Unload formImport
End Sub

Function Test_File_Exist_With_Dir(filePath As String) As Boolean
    Dim TestStr As String
    On Error Resume Next
    TestStr = Dir(filePath)
    On Error GoTo 0
    Test_File_Exist_With_Dir = TestStr <> vbNullString
End Function


Sub Set_Up_Word_Doc(wordapp As Word.Application, objSelection As Selection, objdoc As Document, ByRef program As clsProgram)
'    Dim Customer As String
'    Dim Range As String
'    Dim Model As String
'    Dim Location As String
'    Dim Program_no As String
'    Dim Part_no As String
'    Dim Drawer As String
    Dim answer As String
'    Dim Revision As Long
    Dim regexLetters As Object
    Dim aBord As Variant
    Dim Word_Created_Date As Date
'    Dim program As clsProgram
'    Set program = New clsProgram
    
    
    program.Word_Created_Date = Date 'set drawn date to current date
    Set regexLetters = New RegExp
    regexLetters.Pattern = "[A-Z]"
    
    Worksheets(inputSheet).Activate
    
    ' *** Read all the program info and validate the input ***
    program.Customer = UCase$(Cells(1, 2)) 'read in customer in uppercase
    If program.Customer <> vbNullString Then
        If Len(program.Customer) > 20 Then
            MsgBox "Customer name is too long, should be 20 characters max."
            End
        End If
    Else
        answer = MsgBox("Customer name is empty, would you like to continue anyway?.", vbQuestion + vbYesNo + vbDefaultButton2, "Checking your input")
        If answer = vbNo Then
            End
        End If
    End If
    program.Range = UCase$(Cells(2, 2)) 'read in range in uppercase
    If program.Range <> vbNullString Then
        If Len(program.Range) > 15 Then
            MsgBox "Range is too long, should be 15 characters max."
            End
        End If
    Else
        answer = MsgBox("Range field is empty, would you like to continue anyway?.", vbQuestion + vbYesNo + vbDefaultButton2, "Checking your input")
        If answer = vbNo Then
            End
        End If
    End If
    program.Model = UCase$(Cells(3, 2)) 'read in model in uppercase
    If program.Model <> vbNullString Then
        If Len(program.Model) > 10 Then
            MsgBox "Model is too long, should be 10 characters max."
            End
        End If
    Else
        answer = MsgBox("Model field is empty, would you like to continue anyway?.", vbQuestion + vbYesNo + vbDefaultButton2, "Checking your input")
        If answer = vbNo Then
            End
        End If
    End If
    program.Location = UCase$(Cells(4, 2)) 'read in location in uppercase
    If program.Location <> vbNullString Then
        If Len(program.Location) > 15 Then
            MsgBox "Location is too long, should be 15 characters max."
            End
        End If
    Else
        answer = MsgBox("Location field is empty, would you like to continue anyway?.", vbQuestion + vbYesNo + vbDefaultButton2, "Checking your input")
        If answer = vbNo Then
            End
        End If
    End If
    
'    If program.Program_no <> vbNullString Then
'        If Len(program.Program_no) > 9 Then
'            MsgBox "Program number is too long, should be 9 characters max."
'            End
'        End If
'    Else
'        answer = MsgBox("Program number is empty, would you like to continue anyway?.", vbQuestion + vbYesNo + vbDefaultButton2, "Checking your input")
'        If answer = vbNo Then
'            End
'        End If
'    End If
    program.Part_no = UCase$(Cells(5, 2)) 'read in part_no in uppercase
    If program.Part_no <> vbNullString Then
        If Len(program.Part_no) > 14 Then
            MsgBox "Part number is too long, should be 14 characters max."
            End
        End If
    Else
        answer = MsgBox("Part number is empty, would you like to continue anyway?.", vbQuestion + vbYesNo + vbDefaultButton2, "Checking your input")
        If answer = vbNo Then
            End
        End If
    End If
    program.Drawer = UCase$(Cells(6, 2)) 'read in drawer in uppercase
    If program.Drawer <> vbNullString Then
        If Len(program.Drawer) > 5 Then
            MsgBox "Drawer name is too long, should be 5 characters max."
            End
        End If
    Else
        answer = MsgBox("Drawer field is empty, would you like to continue anyway?.", vbQuestion + vbYesNo + vbDefaultButton2, "Checking your input")
        If answer = vbNo Then
            End
        End If
    End If
'    If IsNumeric(Cells(7, 2)) And Cells(7, 2) <> vbNullString Then
'        program.Revision = Cells(7, 2) 'read in revision
'    Else
'        MsgBox "Revision needs to be a number, program will end"
'        End
'    End If
'    If program.Revision > 999 Then
'        MsgBox "Revision number is too long, should be 3 characters max."
'        End
'    End If

    
    Dim filePath As String
    'work out revision number + program no
    filePath = "C:\Users\" & Environ("username") & "\Dropbox (BCA_TECH)\Kabatech Program\Input Sheets\" & program.Part_no
    Dim TestStr As String
    On Error Resume Next
    TestStr = Dir(filePath)
    On Error GoTo 0
    If TestStr <> vbNullString Then 'if true then file already exists
        Dim NewFPath As String
        NewFPath = "C:\Users\" & Environ("username") & "\Dropbox (BCA_TECH)\Kabatech Program\Input Sheets\" & "Archive\" & program.Part_no & "\"
        Dim fso As Object
        Set fso = CreateObject("scripting.filesystemobject")
        If fso.FolderExists(NewFPath) = False Then
            MkDir (NewFPath)
        End If

        Dim revFPath As String
        revFPath = NewFPath & program.Part_no & "(" & (program.Revision) & ")" 'examples - rev0(0) - rev0(1) - rev1(0)
        TestStr = Dir(revFPath)
        While TestStr <> vbNullString 'if true then file already exists
            program.Revision = program.Revision + 1
            revFPath = NewFPath & program.Part_no & "(" & program.Revision & ")"
            TestStr = Dir(revFPath)
        Wend
        program.Revision = program.Revision + 1
        
    Else 'doesnt already exist
        program.Revision = 0
    End If
        
    program.Program_no = program.Part_no & "(" & program.Revision & ")" 'program_no is now part number + revision number
    
    
    
    '---> setup some word document settings
    With objdoc.PageSetup
        .LeftMargin = wordapp.InchesToPoints(0.5)
        .RightMargin = wordapp.InchesToPoints(0.5)
        .TopMargin = wordapp.InchesToPoints(0.75)
        .BottomMargin = wordapp.InchesToPoints(0.75)
    End With
    objSelection.WholeStory
    objSelection.Font.Name = "Calibri" 'font
    objSelection.Font.Size = 14 'font size
    objSelection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    objSelection.ParagraphFormat.LineSpacing = 20
    With objSelection.Paragraphs.TabStops 'set tabstops for body
        .Add position:=wordapp.InchesToPoints(0.27), Alignment:=wdAlignTabLeft
        .Add position:=wordapp.InchesToPoints(0.7), Alignment:=wdAlignTabLeft
        .Add position:=wordapp.InchesToPoints(3.1), Alignment:=wdAlignTabLeft
        .Add position:=wordapp.InchesToPoints(6.5), Alignment:=wdAlignTabLeft
    End With
    
    With objSelection.Sections(1) 'set up borders
        .Borders.SurroundFooter = True
        .Borders.SurroundHeader = True
        For Each aBord In .Borders
            aBord.LineStyle = wdLineStyleSingle
            aBord.LineWidth = wdLineWidth075pt
        Next aBord
    End With


    'Write to header and footer
    Dim i As Long
    For i = 1 To objSelection.Sections.Count
        With objSelection.Sections(i)
        
            '-----------Header------------
            With .Headers(wdHeaderFooterPrimary).Range
                .text = (program.Customer & " " & program.Range & " " & program.Model & " " & program.Location & Chr$(9) & "Program-" & program.Program_no)
                .Font.Name = "Cambria(Headings)"
                '.Font.Bold = True
                .Font.Size = 16
                With .Borders(wdBorderBottom)
                 .LineStyle = wdLineStyleThickThinSmallGap
                 .LineWidth = wdLineWidth300pt
                 .Color = RGB(98, 36, 25)
                End With
                With .ParagraphFormat
                    .TabStops.ClearAll
                    .TabStops.Add position:=wordapp.InchesToPoints(4.3), Alignment:=wdAlignTabLeft
                End With
            End With
            
            '-----------Footer------------
            With .Footers(wdHeaderFooterPrimary).Range
                 .Font.Name = "Calibri (Body)"
                .text = (Chr$(9) & "DRAWN BY" & Chr$(9) & "DATE" & Chr$(9) & "REVISION" & Chr$(9) & "PART NUMBER" & Chr$(11) & Chr$(9) & program.Drawer & Chr$(9) & program.Word_Created_Date & Chr$(9) & program.Revision & Chr$(9) & program.Part_no)
                .Font.Bold = True
                .Font.Size = 11
                Dim C As Long
                For C = 1 To 36 'highlight the top line only
                    With .Characters(C)
                        .HighlightColorIndex = wdTeal
                        .Font.ColorIndex = wdWhite
                    End With
                Next
                With .ParagraphFormat
                    .TabStops.ClearAll
                    .TabStops.Add position:=wordapp.InchesToPoints(0.4), Alignment:=wdAlignTabCenter
                    .TabStops.Add position:=wordapp.InchesToPoints(2.3), Alignment:=wdAlignTabCenter
                    .TabStops.Add position:=wordapp.InchesToPoints(4.5), Alignment:=wdAlignTabCenter
                    .TabStops.Add position:=wordapp.InchesToPoints(6.5), Alignment:=wdAlignTabCenter
                End With
            End With
        End With
    Next
    'Set Set_Up_Word_Doc = program
End Sub

Function count_number_of_branches() As Long
    Dim i As Long 'count for branch number
    i = StartRow 'set i to start row
    If Cells(i, 1) <> vbNullString Then 'Make sure something is enetered in first cell
        Do While end_case(i + 1) <> True 'Loop until we reach the end case
            i = i + 1
        Loop
        count_number_of_branches = (i - StartRow + 1) 'return the number of branches
    Else
        count_number_of_branches = 0
    End If
    
End Function

Function end_case(i As Long) As Boolean
    If Cells(i, 1) = vbNullString And Cells(i, 2) = vbNullString And Cells(i, 3) = vbNullString Then
        end_case = True
        If Cells(i + 1, 1) <> vbNullString And Cells(i + 1, 2) <> vbNullString And Cells(i + 1, 3) <> vbNullString Then   'if the next line is not empty
            MsgBox "Blank row found not at the end. Line - " & i
            End
        End If
    Else
        end_case = False
    End If
End Function

Sub create_log_file(ByRef program As clsProgram)
    Dim filePath As String
    Dim SecondFP As String 'save in two locations
    Dim StrNm As String: StrNm = Format(Now(), "DD-MM-YYYY hh.mm.ss")
    filePath = "C:\Users\" & Environ("username") & "\Dropbox (BCA_TECH)\Kabatech Program\Log Files\" & "Log File - " & program.Program_no & " rev-" & program.Revision & " " & StrNm & ".txt" 'WORKPATH
    SecondFP = "U:\James.B\C16 Program\Log Files\" & "Log File - " & program.Program_no & " rev-" & program.Revision & " " & StrNm & ".txt" 'WORKPath
    'filePath = "C:\Users\James\Dropbox (BCA_TECH)\James\Excel Programs\C16 Program James\Example File Structure\Log Files\" & "Log File - " & program.Program_no & " rev-" & program.Revision & " " & StrNm & ".txt" 'HOMEPATH
    'SecondFP = "C:\Users\James\Dropbox (BCA_TECH)\James\Excel Programs\C16 Program James\Logs\" & "Log File - " & program.Program_no & " rev-" & program.Revision & " " & StrNm & ".txt" 'HomePath
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim fileStream As TextStream
    Set fileStream = fso.CreateTextFile(filePath)
    fileStream.WriteLine "Log File Created - " & Format(Now(), """Date -"" DD-MM-YYYY  ""Time -"" hh:mm:ss") 'DATE AND TIME
    fileStream.WriteLine "Application User - " & Application.userName '& program.Drawer 'Application user and release intials
    fileStream.WriteLine "Program Number = " & program.Program_no
    fileStream.WriteLine "Revision = " & program.Revision
    fileStream.WriteLine vbNewLine & "Buildsheets Word:" & vbNewLine & "Time and date saved - " & program.WordSaved & vbNewLine & "Overwritten - " & program.WordOverwritten 'word log
    fileStream.WriteLine vbNewLine & "C16 Program:" & vbNewLine & "Time and date saved - " & program.ProgSaved & vbNewLine & "Overwritten - " & program.ProgOverwritten 'c16 program log
    fileStream.WriteLine vbNewLine & "Input Sheet:" & vbNewLine & "Time and date saved - " & program.InputSheetSaved & vbNewLine & "Overwritten - " & program.InputSheetOverwritten 'input sheet log
    fileStream.Close
    SetAttr filePath, vbReadOnly 'set the file as read only
    fso.CopyFile filePath, SecondFP
End Sub

Sub Run(Optional control As IRibbonControl)
    Dim wires_to_mark As String
    Dim bnew As Boolean
    Dim numBranches As Long
    Dim branches() As New clsbranch
    Dim program As New clsProgram
    

    
    AutoAdapter = True 'set the auto adapter to true as we are running an automated program
    Worksheets(inputSheet).Activate
    wires_to_mark = UCase$(Cells(1, 7))
    numBranches = count_number_of_branches 'count the number of lines on the input sheet (number of branches) so we can set up an array of the correct size
    If numBranches <> 0 Then 'make sure the input sheet is not empty
        ReDim branches(numBranches - 1) 'set up an array to store all branches
        branches = Read_Branches_From_Input_Sheet(numBranches - 1, program) 'get all branch details in an array from input sheet
        'Set up a word document
        Dim wordapp As Word.Application
        On Error Resume Next
        Set wordapp = GetObject(, "Word.Application")
        If wordapp Is Nothing Then   'if word is not open then open it
            Set wordapp = CreateObject("Word.Application")
            bnew = True
        End If
        On Error GoTo 0 'reset error warnings
        Dim objdoc As Document
        Set objdoc = wordapp.Documents.Add '(Visible:=True)
        objdoc.ActiveWindow.View.Zoom.Percentage = 100
        Dim objSelection As Selection
        Set objSelection = wordapp.Selection
        On Error GoTo wordError
        'Set program =
        Call Set_Up_Word_Doc(wordapp, objSelection, objdoc, program)    'print header and footer to word doc and check some user input
        
        Clear_Cells_of_Formatting_Program 'clear the program sheet
        Dim first As Boolean
        Dim i As Long
        For i = 0 To (numBranches - 1)
            If branches(i).getDescription <> MBH And branches(i).getDescription <> DP And branches(i).getCurrentPos = False Then
                If branches(i).getPreviousPos > 0 Then
                    position = branches(i - branches(i).getPreviousPos).getLength - NewLengthAdjustment 'minus length adjusment from position length of the previous branch
                Else
                    position = 10
                End If
                    On Error Resume Next
                    'Write to excel program
                    If branches(i).getInstruction = False Then 'if its not an instruction (just buildsheet)
                        If i = 0 And i = numBranches - 1 Then 'its the first and last branch
                            write_branch_to_program branches(i), 4, program
                        ElseIf i > 0 And i < numBranches - 1 And first = True Then 'not first or last branch
                            write_branch_to_program branches(i), 1, program
                        ElseIf i = 0 Or first = False Then 'first branch
                            write_branch_to_program branches(i), 0, program
                            first = True
                        ElseIf i = numBranches - 1 Then 'The last branch
                            'If branches(i - 2).currentPos = True Then
                            If branches(i).getPreviousPos > 0 Then 'if its a position length
                                If branches(i - 1).getDescription = DP Or branches(i - 1).getDescription = MBH Then
                                    write_branch_to_program branches(i), 3, program
                                Else
                                    write_branch_to_program branches(i), 2, program
                                End If
                            Else
                                write_branch_to_program branches(i), 2, program
                            End If
                        End If
                    End If
            End If
            'Write to word document buildsheet
            If i = 0 And i = numBranches - 1 Then 'first and last branch
                wordapp.Selection.Range.ListFormat.ApplyListTemplateWithLevel _
                    ListTemplate:=wordapp.ListGalleries(wdNumberGallery).ListTemplates(2)
                write_branch_to_word branches(i), objSelection
                If wires_to_mark <> "" Then
                    objSelection.TypeText (Chr$(11) + Chr$(11) + Chr$(9) + "MARK " + wires_to_mark)
                End If
            ElseIf i = 0 Then 'first branch
                'On Error GoTo wordError
                wordapp.Selection.Range.ListFormat.ApplyListTemplateWithLevel _
                    ListTemplate:=wordapp.ListGalleries(wdNumberGallery).ListTemplates(2)
                write_branch_to_word branches(i), objSelection
                objSelection.TypeText (Chr$(10))
            ElseIf i = numBranches - 1 Then 'The last branch
                write_branch_to_word branches(i), objSelection
                If wires_to_mark <> "" Then
                    objSelection.TypeText (Chr$(11) + Chr$(11) + Chr$(9) + "MARK " + wires_to_mark)
                End If
            Else 'every other branch
                write_branch_to_word branches(i), objSelection
                objSelection.TypeText (Chr$(10))
            End If
        Next
        Worksheets(inputSheet).Activate 'set it back to input sheet after creating program
        On Error GoTo wordSaveErr
        
        Dim FPath As String
        Dim oldFPath As String
        FPath = BuildsheetFilePath & program.Program_no & ".docx"
        oldFPath = BuildsheetFilePath & program.Part_no & "(" & program.Revision - 1 & ").docx"
        'if buildsheet word file already exists code
        Dim TestStr As String
        On Error Resume Next
        TestStr = Dir(oldFPath)
        On Error GoTo 0
        If TestStr <> vbNullString Then 'if true then file already exists
            program.WordOverwritten = True
        Else: program.WordOverwritten = False
        End If
        objdoc.SaveAs (FPath) 'Save the current word document
        Dim DateTimeNow As String: DateTimeNow = Format(Now(), "DD-MM-YYYY hh.mm.ss")
        program.WordSaved = DateTimeNow
        SetAttr FPath, vbReadOnly 'set the word document as read only
        If program.Revision > 0 Then 'delete the previous version of word now we have replaced it
            DeleteFile (oldFPath)
        End If
        
        objdoc.Close
        Set objdoc = Nothing
        If wordapp Is Nothing Then 'make sure its not already closed
            Exit Sub
        End If
        If bnew = True Then  'if we opened the word application
            On Error GoTo wordError
            wordapp.Quit 'close word
            Set wordapp = Nothing
        End If
        
        Call Export_As_CLC(program) 'export program
        Call Export_Spreadsheet_To_Text(program) ' export input sheet
        Dim Word As String
        Dim prog As String
        Dim inputSheettxt As String
        If program.WordOverwritten Then Word = "Overwritten" Else Word = "First Release"
        If program.ProgOverwritten Then prog = "Overwritten" Else prog = "First Release"
        If program.InputSheetOverwritten Then inputSheettxt = "Overwritten" Else inputSheettxt = "First Release"
        MsgBox "Completed!" & vbNewLine & vbNewLine & "Word - " & Word & vbNewLine & "Program - " & prog & vbNewLine & "InputSheet - " & inputSheettxt
    Else
        MsgBox "Nothing entered in the first cell"
    End If
    Call create_log_file(program)
    Worksheets(inputSheet).Activate 'set it back to input sheet after creating program
    
Exit Sub
wordError: MsgBox "Connection to Microsoft Word lost. Program not completed. VBA ERROR = " & Err.description
Exit Sub
wordSaveErr: MsgBox "Word document not saved. VBA ERROR = " & Err.description

End Sub

Function Read_Branches_From_Input_Sheet(numBranches As Long, ByRef program As clsProgram) As Variant
    Dim count2 As Long
    Dim x As Long
    Dim Count As Long
    Dim prevPos As Long
    Dim i  As Long
    Dim b() As New clsbranch
    Dim endFlag As Boolean
    If IsNumeric(Cells(7, 2)) And Cells(7, 2) <> vbNullString Then
        program.Final_branch_tape_length = Cells(7, 2)
    Else
        MsgBox "Final branch tape length needs to be a number, program will end"
        End
    End If
    
    i = StartRow 'set i as the startrow
    prevPos = 2
    ReDim b(numBranches) 'set up an array to store all branches
    For Count = 0 To numBranches
        If Count = 2 Then
            Count = 2
        End If
        x = 1 'read in wirecodes
        If Cells(i, x).Interior.Color = changecolor Then 'check for blue changes
           b(Count).setWirecodeChange (True)
        Else
            b(Count).setWirecodeChange (False)
        End If
        If IsLetter(Mid$(Cells(i, x), 1, 5)) And Count = numBranches Then  'check the first 5 characters are letters and its the last line testing if it says "all wires" or similar
            Cells(i, x) = "All WIRES"
            
            b(Count).setTapedLength (True)
            If (b(Count).setWirecode(UCase$(Cells(i, x)), i)) Then 'set the wirecode and check the user input
                 End
            End If
        Else
            If Cells(i, x) = vbNullString Then    'if the "cores in branch" cell is empty its likely because its a taped length
                If Cells(i, x + 1) = vbNullString Then    'the description also should be empty
                    b(Count).setTapedLength (True)
                Else 'if the cores is empty but the description is not this is a buildsheet INSTRUCTION
                    'MsgBox "Error, ""Cores in branch"" is empty but there is a description, if its a taped length leave both empty. Line number - " & i, vbInformation, "User input error - exiting program"
                    'endFlag = True
                    b(Count).setInstruction (True)
                    b(Count).setBlueTape (-1)
                    If b(Count).setDescription(UCase$(Cells(i, x + 1)), i) Then 'set the description and check user input
                        End
                    End If
                End If
            Else
                b(Count).setTapedLength (False)
                If (b(Count).setWirecode(UCase$(Cells(i, x)), i)) Then 'set the wirecode and check the user input
                    End
                End If
            End If
        End If
        If b(Count).getInstruction = False Then 'if its an instruction dont read in anything else
            x = 2 'read description
            If Cells(i, x).Interior.Color = changecolor Then 'check if change has been highlighted
               b(Count).setDescChange (True)
            Else
                b(Count).setDescChange (False)
            End If
            If UCase$(Cells(i, x)) <> vbNullString Then 'make sure there is something entered
                If b(Count).setDescription(UCase$(Cells(i, x)), i) Then 'set the description and check user input
                    End
                End If
            End If
            x = 3 'read length
            If Cells(i, x).Interior.Color = changecolor Then 'check if change has been highlighted
               b(Count).setLengthChange (True)
            Else
                b(Count).setLengthChange (False)
            End If
            If b(Count).getInstruction = False Then 'if its a instruction
                If IsNumeric(Cells(i, x)) And Cells(i, x) <> vbNullString Then 'make sure the length is a number and is not empty
                    If b(Count).setLength(Cells(i, x), i, MinKabatecLength, KabatecLenLimitLow, KabatecLenLimitHigh, b(Count).getPreviousPos) Then 'set the length and check user input
                        End
                    End If
                Else
                    MsgBox "Error. Length must be a number.  Line - " & i, vbExclamation, "User input error - exiting program"
                    endFlag = True
                End If
            End If
            x = 4 'space taped/fully taped (default is fully taped)
            If UCase$(Cells(i, x)) = "Y" Then
                b(Count).setFullyTaped (False)
            ElseIf Cells(i, x) = vbNullString Then
                b(Count).setFullyTaped (True)
            Else 'if its anything else
                MsgBox "Unxepected character for ""Space Taped"" on line - " & i
                endFlag = True
            End If
            x = 5 'Blue tape length
            If Cells(i, x) <> vbNullString Then 'if its not nothing
                If IsNumeric(Cells(i, x)) Then 'if its a number
                    b(Count).setBlueTape (Cells(i, x))
                ElseIf UCase$(Cells(i, x)) = "Y" Then 'if this option is selected we want to insert it onto the buildsheet but make no change to the program
                    b(Count).setBlueTape (0) 'setting it at zero will not change the taping
                Else
                    MsgBox "Error. Unxepected input for Blue Tape.  Line - " & i, vbExclamation, "User input error - exiting program"
                    endFlag = True
                End If
            Else
                b(Count).setBlueTape (-1) 'if empty
            End If
            x = 6 'position length
            If UCase$(Cells(i, x)) = "Y" Or b(Count).getLength = MinKabatecLength And b(Count).getDescription <> MBH And b(Count).getDescription <> DP Then 'position also if length is the minimum
                b(Count).setCurrentPos (True)
                If Count = numBranches Then 'if its the last branch
                    'Last branch cannot be a position
                    MsgBox "Last branch cannot be a position", vbExclamation, "User input error - exiting program"
                    endFlag = True
                ElseIf Count > 0 Then 'make sure it is not the first
                    If b(Count - 1).getCurrentPos And b(Count).getLength <> b(Count - 1).getLength Then 'if the previous branch was a position they must both be the same length to be a position
                        MsgBox "Cannot have two positions in a row with different lengths. Line - " & i, vbExclamation, "User input error - exiting program"
                        endFlag = True
                    End If
                End If
            ElseIf Cells(i, x) <> vbNullString Then 'if its not empty but not a "y"
                MsgBox "Unexpected character for ""Pos Length"" on line " & i
                endFlag = True
            Else 'not a position
                If Count > 0 Then  'if its not the first branch
                    If b(Count).getTapedLength = False And b(Count - 1).getCurrentPos And b(Count).getDescription <> DP And b(Count).getDescription <> MBH And Count <> numBranches Then 'if the current branch is a new spur and the last was a position (this branch is not a position)
                        MsgBox "Cannot have a postion then a new spur that is not also a position. Line - " & i, vbExclamation, "User input error - exiting program" 'error
                        endFlag = True
                    ElseIf b(Count - 1).getCurrentPos Then 'if the preivous branch is a position
                        b(Count).setPreviousPos (1)    'set previous position  **************
                        'If count <> numBranches Then
                        b(Count).setTapedLength (False) 'it is now not a taped length
                        'End If
                    ElseIf b(Count - 1).getPreviousPos > 0 Then 'if its more than 0 that means there was a position
                        If b(Count).getDescription = MBH Then  'if there is a measure by hand o precut plug we need to keep track of the last position  removed- b(Count).getDescription = DP Or
                            prevPos = prevPos + 1
                            b(Count).setPreviousPos (prevPos)   '*******************
                         ElseIf b(Count - 1).getDescription = MBH Then 'removed - b(Count - 1).getDescription = DP Or
                            b(Count).setPreviousPos (prevPos)   '******************
                            b(Count).setTapedLength (False) 'it is now not a taped length
                            prevPos = 2
                        End If
                    End If
                End If
                b(Count).setCurrentPos (False) 'set the current position as false
            End If
            
            If b(Count).getCurrentPos And b(Count).getBlueTape >= 0 Then
                MsgBox "Cannot have a postion and blue tape, Line - " & i
                endFlag = True
            End If
            'if the length is less than 210 it will not cut
'            If b(Count).getLength < MinCuttingLength And b(Count).getTapedLength = False And b(Count).getCurrentPos = False And b(Count).getDescription <> MBH And b(Count).getDescription <> DP And b(Count).getLength <> MinKabatecLength And b(Count).getPreviousPos < 1 Then
'                If MsgBox("The machine will not automatically cut if a new spur is below 100 after removing the length adjustment, current length is " & b(Count).getLength & ". Would you like to automatically adjust this length to " & MinCuttingLength & " on Line" & i & "?", vbYesNo) = vbYes Then
'                    If b(Count).setLength(MinCuttingLength, i, MinKabatecLength, KabatecLenLimitLow, KabatecLenLimitHigh, b(Count).getPreviousPos) Then End 'set the length and check user input
'                    Cells(i, 3) = MinCuttingLength 'change the cell so the correct value will be exported
'                End If
'            End If
        End If
        If b(Count).getBlueTape >= 0 Then 'count another build sheet line if blue tape is added
            count2 = count2 + 2 'if blue tape mark is set keep track of the last line number for this branch
        Else
            count2 = count2 + 1 'if not set just add one
        End If
        If Count = numBranches Then
            If program.Final_branch_tape_length > b(Count).getLength Then
                MsgBox "Final Branch tape is more than the actual length ", vbExclamation, "User input error - exiting program"
                endFlag = True
            End If
        End If
        b(Count).setBuildSheetLine (count2)
        'b(count).setBranchNum (count + 1)
        i = i + 1
    Next
    If endFlag = True Then End
    Read_Branches_From_Input_Sheet = b 'return the array
End Function


Sub write_branch_to_program(b As clsbranch, firstLastCheck As Long, ByRef program As clsProgram)
    Dim lengthToRemove As Integer
    Worksheets(ProgramSheet).Activate
    If b.getPreviousPos > 0 And firstLastCheck <> 3 And firstLastCheck <> 2 Then
        length = b.getLength + NewLengthAdjustment
    Else
        length = b.getLength
    End If
    
    If firstLastCheck = 0 Then 'if its the first branch
        If b.getFullyTaped = True Then 'if its fully taped
            Insert_Initial_Spur_Fully_Taped 'Run intial spur code
        Else 'if its spaced taped
            Insert_Initial_Spur_Space_Taped
        End If
    ElseIf firstLastCheck = 2 Then 'if its the last branch
        
'        With Sheets(InputSheet)
'            finalBranchTape = .Cells(9, 2) 'read in final branch length to tape
'            If finalBranchTape > b.getLength Then
'                MsgBox "Final Branch tape is more than the actual length ", vbExclamation, "User input error - exiting program"
'                End
'            End If
'        End With
        lengthToRemove = b.getLength - program.Final_branch_tape_length
        length = b.getLength - lengthToRemove 'take 100 off the last length
        
        If b.getPreviousPos = 0 Then 'if its not a postion
            b.setTapedLength (True) 'set as tape length
        Else
            length = length + NewLengthAdjustment
        End If
    ElseIf firstLastCheck = 3 Then 'last branch but previous position
        lengthToRemove = b.getLength - program.Final_branch_tape_length
        length = b.getLength - 100 + NewLengthAdjustment - lengthToRemove 'take 100 off the last length
    ElseIf firstLastCheck = 4 Then
        lengthToRemove = b.getLength - program.Final_branch_tape_length
        length = b.getLength - lengthToRemove 'take 100 off the last length
        Call b.setTapedLength(True)
    End If
    If firstLastCheck > 0 And firstLastCheck < 5 And length <> 0 Then 'if its a normal branch
        If b.getFullyTaped = True Then 'if its fully taped
            If b.getTapedLength = True Then 'Taped length or branch
                Insert_Taped_Length , b.getBlueTape 'run fully taped length code
            Else
                Start_New_Spur_Fully_Taped , b.getBlueTape 'run fully taped new spur code
            End If
        Else 'if its spaced taped
            If b.getTapedLength = True Then 'Taped length or branch
                Insert_Space_Taped_Length , b.getBlueTape 'run space taped length code
            Else
                Start_New_Spur_Space_Taped , b.getBlueTape 'run space taped new spur code
            End If
        End If
    End If
    
    If firstLastCheck > 1 Then 'if its the last branch finish which the final program code
        Insert_Final_Cut
    End If
End Sub


Sub write_branch_to_word(b As clsbranch, objSelection As Selection)
    If b.getBuildSheetLine < 10 Then 'only use a tab if the buildsheet line is <10
        objSelection.TypeText (Chr$(9)) 'tab
    End If
    If b.getWirecode <> vbNullString Then 'loop through a branches wirecodes and write each char one at a time until the tabstop
        Dim i As Long
        For i = 1 To Len(b.getWirecode)
            If (i Mod 20) = 0 Then 'if its the same pos as the tabstop do this
                Do While Mid$(b.getWirecode, i - 1, 1) <> "." And i <> Len(b.getWirecode) + 1 'print out the remaining wirecode before starting a new line
                    objSelection.TypeText (Mid$(b.getWirecode, i, 1))
                    i = i + 1
                Loop
                If i < Len(b.getWirecode) Then
                    objSelection.TypeText (Chr$(11) + Chr$(9))
                End If
            End If
            If b.getWirecodeChange Then
                objSelection.Font.ColorIndex = WordCyan
            End If
            objSelection.TypeText (Mid$(b.getWirecode, i, 1)) 'just print the character
        Next
    ElseIf b.getInstruction = False Then
        If b.getWirecodeChange Then
            objSelection.Font.ColorIndex = WordCyan
        End If
        objSelection.TypeText ("TAPE LENGTH")
    End If
    If b.getWirecodeChange Then
        objSelection.Font.ColorIndex = wdBlack
    End If
    'objSelection.TypeText (b(a).wirecode)
    'MsgBox "Position = " & objSelection.range.Rows.RelativeHorizontalPosition
    If b.getDescChange Then
        objSelection.Font.ColorIndex = WordCyan
    End If
    If b.getInstruction Then objSelection.Font.Color = vbRed 'if its an instruction
    objSelection.TypeText (Chr$(9) & b.getDescription & Chr$(9))
    If b.getDescChange Or b.getInstruction Then
        objSelection.Font.ColorIndex = wdBlack
    End If
    If b.getInstruction = False Then 'if its not an instruction
        If b.getLengthChange Then
            objSelection.Font.ColorIndex = WordCyan
        End If
    '    If b.length = 100 Then
    '        objSelection.TypeText ("POS10")
    '    Else
        If b.getCurrentPos = True Then
            objSelection.TypeText ("POS " & b.getLength - NewLengthAdjustment)
        Else
            If b.getBlueTape >= 0 Then
                objSelection.TypeText (b.getLength + b.getBlueTape & vbNullString)
            Else
                objSelection.TypeText (b.getLength & vbNullString)
            End If
        End If
        If b.getLengthChange Then
            objSelection.Font.ColorIndex = wdBlack
        End If
        
        If b.getBlueTape >= 0 Then
            objSelection.TypeText (Chr$(10))
            objSelection.Font.ColorIndex = wdBlue
            If b.getBuildSheetLine < 10 Then
                objSelection.TypeText (Chr$(9))
            End If
            objSelection.TypeText ("INSERT BLUE TAPE MARK AS SHOWN ON DRAWING")
        End If
    End If
End Sub

Sub close_word_and_end(objdoc As Document, wordapp As Word.Application, bnew As Boolean)
    objdoc.Close
    Set objdoc = Nothing
    If wordapp Is Nothing Then 'make sure its not already closed
        Exit Sub
    End If
    If bnew = True Then  'if no other documents are open close the word application
        wordapp.Quit 'close word
        Set wordapp = Nothing
    End If
    End
End Sub


 Sub testisletter()
    If IsLetter(Mid$("sk3kkk ", 1, 2)) Then
        MsgBox ("true")
    End If
 End Sub

Function IsLetter(strValue As String) As Boolean 'or space (32)
    Dim intPos As Long
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid$(strValue, intPos, 1))
            Case 65 To 90, 97 To 122, 32
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      ' First remove readonly attribute, if set
      SetAttr FileToDelete, vbNormal
      ' Then delete the file
      Kill FileToDelete
   End If
End Sub

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub test_Wire_to_mark()
    Dim str As String
    Dim length As Long
    length = 0
    str = "X1.X1.D1.D1.D2.D2.D2.X2.X2.C1.C2.X9.C4.X7" '"102.12***.103.f(2222).f(1).f***.134.102**.102.f.f.f.f"
    'count the number of wires
    Dim i As Long
    For i = 1 To Len(str)
        If Mid$(str, i, 1) = "." Then
            length = length + 1
        End If
    Next
    str = (wires_to_mark(str, length))
    MsgBox str
End Sub
'Function to take one big string of all wirecodes + the length of the array, returns a string of wires that need marking
Function wires_to_mark(str As String, length As Long) As String
    Dim wirecode() As String
    Dim i As Long
    Dim Count As Long
    Dim toMark As String
    toMark = "Mark "
    Count = 0
    ReDim wirecode(length)
    'put the string into an array, remove any "*" or "()"
    For i = 1 To Len(str)
        wirecode(Count) = wirecode(Count) + Mid$(str, i, 1)
        Do While Mid$(str, i + 1, 1) = "*"
            i = i + 1
        Loop
        If Mid$(str, i + 1, 1) = "(" Then
            Do While Mid$(str, i, 1) <> ")"
                i = i + 1
            Loop
        End If
        If Mid$(str, i + 1, 1) = "." Then
            Count = Count + 1
            i = i + 1
        End If
    Next
    'loop through the array
    'Dim done As Boolean
    For i = 0 To length
        Dim U As Long
        For U = 0 To length
            'If done = True Then
               ' If wirecode(i) = wirecode(u) Then
                 '   wirecode(i) = "*"
                '    wirecode(u) = "*"
               ' End If
            'End If
            If U = i Or wirecode(i) = "*" Then 'do nothing
            ElseIf wirecode(i) = wirecode(U) Then 'add the wire code to the string
                If toMark = "Mark " Then
                    toMark = toMark + " " + wirecode(i)
                Else
                    toMark = toMark + "." + wirecode(i)
                End If
                Dim x As Long
                For x = i + 1 To length
                    If wirecode(i) = wirecode(x) Then 'remove the already marked wirecode
                        wirecode(x) = "*"
                    End If
                Next
                wirecode(i) = "*"
                'wirecode(u) = "*"
                'done = True
            End If
        Next
        'done = False
    Next
    'MsgBox (toMark)
    wires_to_mark = toMark
    
End Function
Sub test_set_and_get_methods()
    Dim b() As New clsbranch
    ReDim b(3) 'set up an array to store all branches
    b(0).setWirecode "121.123.F.F*"
    b(1).setWirecode "ggf"
    b(2).setWirecode "lalal"
    MsgBox b(0).getWirecode
    MsgBox b(1).getWirecode
    MsgBox b(2).getWirecode
    MsgBox b(0).getWirecode
End Sub

Sub test_regex()
    Dim regexWirecode As Object
    Set regexWirecode = New RegExp
    regexWirecode.Pattern = "^[A-Z0-9\.\*\(\)]+$"
    Dim wirecode As String
    wirecode = "A(1).F (1). F***.123 .123     (1)"
    wirecode = Replace(wirecode, " ", vbNullString) 'remove spaces
    If regexWirecode.test(wirecode) Then    'check if it matches the regular expression
        MsgBox "passes test"
    Else
        MsgBox "failed test"
    End If
End Sub

Sub read_existing_buildsheet()
    Dim FName As String
    Dim wordapp As Word.Application
    Dim objdoc As Document
    Dim objSelection As Selection
    Dim bnew As Boolean
    Dim program As New clsProgram
    
    'FName = "C:\Users\James\Dropbox (BCA_TECH)\James\Excel Programs\C16 Program James\FEXP1803.docx" 'Home?
    FName = "C:\Users\james.b\Dropbox (BCA_TECH)\James\Excel Programs\C16 Program James\FEXP1803.docx" 'Work?
    
    On Error Resume Next
    Set wordapp = GetObject(, "Word.Application")
    If wordapp Is Nothing Then   'if word is not open then open it
        Set wordapp = CreateObject("Word.Application")
        bnew = True
    End If
    On Error GoTo 0 'reset error warnings
    Set objdoc = wordapp.Documents.Open(FName)
    objdoc.ActiveWindow.View.Zoom.Percentage = 100
    Set objSelection = wordapp.Selection



    
    Dim text As String
    
    With objSelection.Sections(1)
        '-----------Header------------
        With .Headers(wdHeaderFooterPrimary).Range
            '.text = (customer & " " & range & " " & model & " " & location & Chr$(9) & "Program-" & program_no & Chr$(10))
            text = .text
            Dim start As Long
            Dim prog As String
            prog = "Program-"
            start = (InStr(text, prog) + Len(prog))
            program.Program_no = Mid(text, start, 8)
        End With
        
        
        '-----------Footer------------
        With .Footers(wdHeaderFooterPrimary).Range
            '.text = (Chr$(9) & "DRAWN BY" & Chr$(9) & "DATE" & Chr$(9) & "REVISION" & Chr$(9) & "PART NUMBER" & Chr$(11) & Chr$(9) & drawer & Chr$(9) & Date & Chr$(9) & revision & Chr$(9) & part_no)
'            text = .text
'            Dim txt As String
'            txt = Chr$(7) & Chr$(13) & Chr$(10) & Chr$(7)
'            start = (InStr(text, txt) + Len(txt))
'            text = Mid(text, start)
        End With
    End With
    
    Dim singleLine As Paragraph
    Dim lineText As String
    Dim length As Long
    Dim previous As Long
    Dim str As Variant
    Dim splitline() As String
    Dim Count As Long
    Count = 0
    Dim regexLineFeeds As Object
    Set regexLineFeeds = New RegExp
    regexLineFeeds.Global = True
    regexLineFeeds.Pattern = "[\r\n]"
    For Each singleLine In objdoc.Paragraphs
        Dim b As New clsbranch
        lineText = singleLine.Range.text
        lineText = regexLineFeeds.Replace(lineText, "")
        'lineText = Replace(lineText, " ", "") 'replace any spaces
        'lineText = Replace(lineText, vbNewLine, "")
        splitline = Split(lineText, vbTab)
        
        Dim AryNoBlanks() As Variant
        Dim Counter As Long, NoBlankSize As Long
        
        'set references and initialize up-front
        ReDim AryNoBlanks(0 To 0)
        NoBlankSize = 0
        
        'loop through the array from the range, adding
        'to the no-blank array as we go
        For Counter = LBound(splitline) To UBound(splitline)
            If splitline(Counter) <> "" Then
                NoBlankSize = NoBlankSize + 1
                AryNoBlanks(UBound(AryNoBlanks)) = splitline(Counter)
                ReDim Preserve AryNoBlanks(0 To UBound(AryNoBlanks) + 1)
            End If
        Next Counter
        
        'remove that pesky empty array field at the end
        If UBound(AryNoBlanks) > 0 Then
            ReDim Preserve AryNoBlanks(0 To UBound(AryNoBlanks) - 1)
        End If
        
        

        'Do something with the array
        'if there are three elements it is likely a new branch
        If UBound(AryNoBlanks) - LBound(AryNoBlanks) + 1 = 3 Then
            b.setWirecode (AryNoBlanks(0)), , True
            b.setDescription (AryNoBlanks(1))
            Dim strlong As String
            strlong = AryNoBlanks(2)
            strlong = Replace(strlong, vbNewLine, "")
            Dim lngValue As Long
            If IsNumeric(strlong) Then
                lngValue = CLng(strlong)
            ElseIf InStr(1, strlong, "pos", 1) Then
                lngValue = removeLettersConvertToLong(strlong)
                b.setCurrentPos (True)
                lngValue = lngValue + NewLengthAdjustment
            Else
                MsgBox "Unxepexted input on line - " & Count
                End
            End If
            b.setLength (lngValue)
        ElseIf UBound(AryNoBlanks) - LBound(AryNoBlanks) + 1 = 2 Then
            b.setTapedLength True
            b.setLength (AryNoBlanks(1))
        ElseIf UBound(AryNoBlanks) - LBound(AryNoBlanks) + 1 = 1 Then
            Count = Count - 1
            b.setBlueTape 0
        Else
            MsgBox "unxepected number of elements in array"
        End If
        'if there are two
'        start = 1 + previous
'        length = InStr(lineText, vtab) - start
'        wirecode = Mid(lineText, start, length)
'        b(i).setWirecode
'        b(i).setDescription
'        b(i).setLength
'        b(i).setBlueTape
'        b(i).setTapedLength
        program.Add (b)
        Count = Count + 1
    Next singleLine
    
    'program.Customer = "James"
    'text = program.Customer
    
    close_word_and_end objdoc, wordapp, bnew
End Sub

Sub test()
    Dim b1 As New clsbranch
    Dim b2 As New clsbranch
    Dim branches As New clsBranches
    Dim program As New clsProgram
    b1.setDescription ("hello")
    b1.branchNr = 1
    b2.setDescription ("2323")
    b2.branchNr = 2
    branches.Add b1
    branches.Add b2
    Dim b3 As New clsbranch
    b3 = branches.item(2)
    b3 = b1
End Sub


Function removeLettersConvertToLong(strlong As String) As Long
    Dim regexLetters As Object
    Set regexLetters = New RegExp
    regexLetters.Global = True
    regexLetters.Pattern = "[a-zA-Z\ \vbnewline]"
    
    strlong = regexLetters.Replace(strlong, "")
    removeLettersConvertToLong = CLng(strlong)
    
End Function

Sub split_header_text(text As String, program As clsProgram)
    For i = 0 To Len(text)
        
    Next
End Sub



