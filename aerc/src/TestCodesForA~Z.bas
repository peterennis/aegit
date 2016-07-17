Option Compare Database
Option Explicit
Option Private Module

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Type aeLogger
    blnNoTrace As Boolean
    blnNoEnd As Boolean
    blnNoPrint As Boolean
    blnNoTimer As Boolean
End Type

Private lngIndent As Long
Private mlngStartTime As Long
Private mlngEndTime As Long
Private aeLog As aeLogger

Public Sub aeTestLogging()
    aegitClassLoggingTestA
    aegitClassLoggingTestB
    aegitClassLoggingTestC

End Sub

Private Function PassFail(ByVal blnPassFail As Boolean, Optional ByVal varOther As Variant) As String
    On Error GoTo 0
    If Not IsMissing(varOther) Then
        PassFail = "NotUsed"
        Exit Function
    End If
    If blnPassFail Then
        PassFail = "Pass"
    Else
        PassFail = "Fail"
    End If
End Function

Private Sub aegitClassLoggingTestA(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    Dim oDbObjects As aegit_expClass
    Set oDbObjects = New aegit_expClass

    Dim blnTestA As Boolean

TestA:
    '=============
    ' TEST A
    '=============
    If IsMissing(varDebug) Then
        aeBeginLogging "DocumentTheDatabase"
        blnTestA = oDbObjects.DocumentTheDatabase()
    Else
        aeBeginLogging "DocumentTheDatabase", "WithDebugging"
        blnTestA = oDbObjects.DocumentTheDatabase("WithDebugging")
    End If

RESULTS:
    Debug.Print "Test A: DocumentTheDatabase"
    Debug.Print PassFail(blnTestA)
    aeEndLogging "DocumentTheDatabase"

PROC_EXIT:
    Set oDbObjects = Nothing
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aegitClassLoggingTestA"
            Stop
    End Select

End Sub

Private Sub aegitClassLoggingTestB(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    Dim oDbObjects As aegit_expClass
    Set oDbObjects = New aegit_expClass

    Dim blnTestB As Boolean

TestB:
    '=============
    ' TEST B
    '=============
    If IsMissing(varDebug) Then
        aeBeginLogging "Exists"
        blnTestB = oDbObjects.Exists("Modules", "aegit_expClass")
    Else
        aeBeginLogging "Exists", "WithDebugging"
        blnTestB = oDbObjects.Exists("Modules", "aegit_expClass", "WithDebugging")
    End If

RESULTS:
    Debug.Print "Test B: Exists"
    Debug.Print PassFail(blnTestB)
    aeEndLogging "Exists"

PROC_EXIT:
    Set oDbObjects = Nothing
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aegitClassLoggingTestB"
            Stop
    End Select

End Sub

Private Sub aegitClassLoggingTestC(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    Dim oDbObjects As aegit_expClass
    Set oDbObjects = New aegit_expClass

    Dim blnTestC As Boolean

TestC:
    '=============
    ' TEST C
    '=============
    If IsMissing(varDebug) Then
        aeBeginLogging "GetReferences"
        blnTestC = oDbObjects.GetReferences()
    Else
        aeBeginLogging "GetReferences", "WithDebugging"
        blnTestC = oDbObjects.GetReferences("WithDebugging")
    End If

RESULTS:
    Debug.Print "Test C: GetReferences"
    Debug.Print PassFail(blnTestC)
    aeEndLogging "GetReferences"

PROC_EXIT:
    Set oDbObjects = Nothing
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aegitClassLoggingTestC"
            Stop
    End Select

End Sub

Private Sub TestLogging()
    On Error GoTo 0
    aeBeginLogging "MyProc", "Some Parameter varOne", "varTwo", "varThree"
    ' more code here
    aePrintLog "Some Status"
    ' more code here
    aeEndLogging "MyProc"
End Sub

Private Sub aeBeginLogging(ByVal strProcName As String, Optional ByVal varOne As Variant = vbNullString, _
        Optional ByVal varTwo As Variant = vbNullString, Optional ByVal varThree As Variant = vbNullString)
    On Error GoTo 0
    mlngStartTime = timeGetTime()
    'Debug.Print ">aeBeginLogging"; Space$(1); "mlngStartTime=" & mlngStartTime
    If aeLog.blnNoTrace Then
        'Debug.Print "B1: aeBeginLogging", "blnNoTrace=" & aeLog.blnNoTrace
        Exit Sub
    End If
    If Not aeLog.blnNoTimer Then
        'Debug.Print "B2: aeBeginLogging", "blnNoTimer=" & aeLog.blnNoTimer
        Debug.Print Format$(mlngStartTime, "0.00"); Space$(2);
    End If
    Debug.Print Space$(lngIndent * 4); strProcName; Space$(1); "'" & varOne & "'"; Space$(1); "'" & varTwo & "'"; Space$(1); "'" & varThree & "'"
    lngIndent = lngIndent + 1
End Sub

Private Sub aeEndLogging(ByVal strProcName As String, Optional ByVal varOne As Variant = vbNullString, _
        Optional ByVal varTwo As Variant = vbNullString, Optional ByVal varThree As Variant = vbNullString)
    On Error GoTo 0
    If aeLog.blnNoTrace Then
        'Debug.Print "E1: aeEndLogging", "blnNoTrace=" & aeLog.blnNoTrace
        Exit Sub
    End If
    lngIndent = lngIndent - 1
    mlngEndTime = timeGetTime()
    If Not aeLog.blnNoEnd Then
        If Not aeLog.blnNoTimer Then
            'Debug.Print ">aeEndLogging"; Space$(1); "mlngEndTime=" & mlngEndTime
            mlngEndTime = timeGetTime()
            'Debug.Print "E2: aeEndLogging", "blnNoTimer=" & aeLog.blnNoTimer
            Debug.Print Format$(mlngEndTime, "0.00"); Space$(2);
        End If
        Debug.Print Space$(lngIndent * 4); "End " & lngIndent; Space$(1); varOne; Space$(1); varTwo; Space$(1); varThree
        Debug.Print "It took " & (mlngEndTime - mlngStartTime) / 1000 & " seconds to process " & "'" & strProcName & "' procedure"
    End If
End Sub

Public Sub aePrintLog(Optional ByVal varOne As Variant = vbNullString, _
        Optional ByVal varTwo As Variant = vbNullString, Optional ByVal varThree As Variant = vbNullString)
    On Error GoTo 0
    If aeLog.blnNoTrace Or aeLog.blnNoPrint Then
        Exit Sub
    End If
    If Not aeLog.blnNoTimer Then
        Debug.Print Format$(timeGetTime(), "0.00"); Space$(2);
    End If
    Debug.Print Space$(lngIndent * 4); "'" & varOne & "'"; Space$(1); "'" & varTwo; "'"; Space$(1); "'" & varThree; "'"
End Sub