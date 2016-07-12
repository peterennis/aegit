Option Compare Database
Option Explicit
Option Private Module

'@TestModule - change - change - change
Private Assert As Object
'

'@ModuleInitialize
Public Sub ModuleInitialize()
    On Error GoTo 0
    ' This method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")

End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error GoTo 0
    ' This method runs once per module.
End Sub

'@TestInitialize
Public Sub TestInitialize()
    On Error GoTo 0
    ' This method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    On Error GoTo 0
    ' This method runs after every test in the module.
End Sub

'@TestMethod
Public Sub TestMethodIsPK()

    On Error GoTo TestFail

    Dim blnTestResult As Boolean

    Dim oDbObjects As aegit_expClass
    Set oDbObjects = New aegit_expClass

    blnTestResult = oDbObjects.IsPrimaryKey("_tblChart", "id")
    Debug.Print "blnTestResult=" & blnTestResult
    Assert.IsTrue blnTestResult

TestExit:
    Set oDbObjects = Nothing
    Exit Sub
TestFail:
    Assert.Fail "TestMethodIsPK raised an error: #" & Err.Number & " - " & "Erl=" & Erl & " " & Err.Description
End Sub

'@TestMethod
Public Sub TestMethodWithManyLettersInTheNameButThereCouldBeManyMore()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Inconclusive

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "TestMethodWithManyLettersInTheNameButThereCouldBeManyMore raised an error: #" & Err.Number & " - " & Err.Description
End Sub