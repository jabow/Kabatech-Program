Attribute VB_Name = "TestModule1"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.

End Sub

'@TestMethod("Branch")
Private Sub SetWirecodeWithCorrectWirecode()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Branch As New Branch
    Dim wirecode As String
    wirecode = "121.121(2).F*.K***"

    'Act:
    Branch.setWirecode (wirecode)

    'Assert:
    Assert.istrue (Branch.getWirecode = wirecode)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Branch")
Private Sub SetWirecodeWithIncorrectChar()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Branch As New Branch
    Dim wirecode As String
    wirecode = "121.121(2).F*.K***.%"

    'Act:
    Branch.setWirecode (wirecode)

    'Assert:
    Assert.IsFalse (Branch.getWirecode = wirecode)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Branch")
Private Sub SetWirecodeWithIncorrectLetter()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Branch As New Branch
    Dim wirecode As String
    wirecode = "121.121(2).a*.K***"

    'Act:
    Branch.setWirecode (wirecode)

    'Assert:
    Assert.IsFalse (Branch.getWirecode = wirecode)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Branch")
Private Sub SetWirecodeWithIncorrectLength()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Branch As New Branch
    Dim wirecode As String
    wirecode = "121.121(2).a*.K*******"

    'Act:
    Branch.setWirecode (wirecode)

    'Assert:
    Assert.IsFalse (Branch.getWirecode = wirecode)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub
