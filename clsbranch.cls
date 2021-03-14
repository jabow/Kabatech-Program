VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsbranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Branch
'@folder("Branch")

    Private m_branchNr As Variant
    Private buildSheetLine As Long
    Private wirecodes As String
    Private description As String
    Private length As Long
    Private tapedLength As Boolean
    Private fullyTaped As Boolean
    Private blueTape As Long
    Private previousPos As Long '>0 if previous branch was a position
    Private currentPos As Boolean 'true if this branch is a position
    Private instruction As Boolean
    'boolean variables to store blue change mark
    Private wirecode_ch As Boolean
    Private desc_ch As Boolean
    Private length_ch As Boolean
    '
    


Property Let branchNr(bn As Variant)
    m_branchNr = bn
End Property

Property Get branchNr() As Variant
    branchNr = m_branchNr
End Property

Public Function setWirecode(wirecode As String, Optional ByVal i As Long, Optional old As Boolean) As Boolean
    If old Then 'if reading in an old buildsheet dont perform any checks
        wirecodes = wirecode
    Else
        '****Check the input matches the following and remove any spaces****
        Dim regexWirecode As Object
        Dim allWires As String
        Dim g As Long
        
        allWires = "ALL WIRES"
        Set regexWirecode = New RegExp
        regexWirecode.Pattern = "^[A-Z0-9\.\*\(\)\-\_\+\\\/]+$" 'pattern the input needs to match
        If wirecode <> allWires Then
            wirecode = Replace(wirecode, " ", vbNullString) 'remove spaces from input
        Else
            
        End If
        If regexWirecode.test(wirecode) Or wirecode = allWires Then     'check if it matches the regular expression
            '****Check each wirecode length****
            Dim single_wirecode As String
            single_wirecode = vbNullString
            For g = 1 To Len(wirecode)
                If Mid$(wirecode, g, 1) <> "." Then
                    single_wirecode = single_wirecode + Mid$(wirecode, g, 1)
                Else
                    single_wirecode = vbNullString
                End If
                If g = Len(wirecode) Or Mid$(wirecode, g, 1) = "." Then
                    If Len(single_wirecode) > 10 And single_wirecode <> allWires Then
                        MsgBox "Error, wirecodes of more than 10 characters are not allowed. Ending program! Line - " & i & " Wirecode - " & single_wirecode, vbExclamation, "User input error - exiting program"
                        setWirecode = True
                        Exit Function
                    End If
                End If
            Next
            wirecodes = wirecode
            setWirecode = False
        Else 'it doesnt match the pattern
            MsgBox "Error, the wirecodes should only contain the following characters - A-Z 0-9 ""."" ""()"" ""*"" ""-"" ""_"" ""+"" ""/"" ""\""  Line - " & i, vbExclamation, "User input error - exiting program"""
            setWirecode = True
            Exit Function
        End If
    End If
End Function

Public Function getWirecode() As String
    getWirecode = wirecodes
End Function

Public Function setDescription(desc As String, Optional i As Long) As Boolean
    If Len(desc) < 25 Then
        description = desc
        setDescription = False
    Else
        MsgBox "Error. Please enter a description less than 26 characters.  Line - " & i, vbExclamation, "User input error - exiting program"""
        setDescription = True 'user input error
    End If
End Function

Public Function getDescription() As String
    getDescription = description
End Function
    
Public Function setLength(l As Long, Optional i As Long, Optional MinKabatecLength As Long, Optional lowLimit As Long, Optional highLimit As Long, Optional prevPos As Long) As Boolean
    If i = Null Then
        length = l 'set the length
        setLength = False 'passed
    Else
        If l > lowLimit And l < highLimit And prevPos = 0 And currentPos = False And description <> MBH And description <> DP And tapedLength = False Then 'if its between these values and the previous or current branch is not a position and is not measured by hand
            'l = highLimit 'automatically set it as the upper limit
            MsgBox "Length on line " & i & " is between " & lowLimit & " and " & highLimit & " Ending program."
            setLength = True 'failed
        ElseIf l < MinKabatecLength And tapedLength = False And prevPos = 0 And description <> MBH And description <> DP Then 'if the length is less than the kabatec limit and it is not a taped length or position length
            MsgBox "Length is too short, a new branch length has to be more than " & MinKabatecLength & " at line - " & i & ". Program will end!", vbExclamation, "User input error - exiting program"
            setLength = True 'failed
        ElseIf l < 1 Then 'make sure length is more than 0 if it is MBH or DP
            MsgBox "Length is less than 1, at line - " & i & ". Program will end!", vbExclamation, "User input error - exiting program"
            setLength = True 'failed
        Else
            length = l 'set the length
            setLength = False 'passed
        End If
    End If
End Function

Public Function getLength() As Long
    getLength = length
End Function

Public Sub setTapedLength(tl As Boolean)
    tapedLength = tl
End Sub

Public Function getTapedLength() As Boolean
    getTapedLength = tapedLength
End Function

Public Sub setFullyTaped(ft As Boolean)
    fullyTaped = ft
End Sub

Public Function getFullyTaped() As Boolean
    getFullyTaped = fullyTaped
End Function

Public Sub setBlueTape(bt As Long)
    blueTape = bt
End Sub

Public Function getBlueTape() As Long
    getBlueTape = blueTape
End Function

Public Sub setPreviousPos(pp As Long)
    previousPos = pp
End Sub

Public Function getPreviousPos() As Long
    getPreviousPos = previousPos
End Function

Public Sub setCurrentPos(cp As Boolean)
    currentPos = cp
End Sub

Public Function getCurrentPos() As Boolean
    getCurrentPos = currentPos
End Function

Public Sub setInstruction(i As Boolean)
    instruction = i
End Sub

Public Function getInstruction() As Boolean
    getInstruction = instruction
End Function

Public Sub setBuildSheetLine(bsl As Long)
    buildSheetLine = bsl
End Sub

Public Function getBuildSheetLine() As Long
    getBuildSheetLine = buildSheetLine
End Function

Public Sub setWirecodeChange(wc As Boolean)
    wirecode_ch = wc
End Sub

Public Function getWirecodeChange() As Boolean
    getWirecodeChange = wirecode_ch
End Function

Public Sub setDescChange(dc As Boolean)
    desc_ch = dc
End Sub

Public Function getDescChange() As Boolean
    getDescChange = desc_ch
End Function

Public Sub setLengthChange(lc As Boolean)
    length_ch = lc
End Sub

Public Function getLengthChange() As Boolean
    getLengthChange = length_ch
End Function



