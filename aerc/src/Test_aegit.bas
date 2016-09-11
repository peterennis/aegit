Option Compare Database
Option Explicit
Option Private Module

Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long

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

Private Sub TestLoggingProcessExample()
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

Private Sub aePrintLog(Optional ByVal varOne As Variant = vbNullString, _
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

Public Sub Run_aeTestLogging()
'    Test_DocumentTheDatabase
'    Test_Exists
'    Test_GetReferences
'    Test_DocumentTables
'    Test_DocumentRelations
'    Test_DocumentTablesXML
    Test_SchemaFile "WithDebugging"

End Sub

Private Sub Test_DocumentTheDatabase(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    Dim oDbObjects As aegit_expClass
    Set oDbObjects = New aegit_expClass

    Dim blnTest As Boolean

    If IsMissing(varDebug) Then
        aeBeginLogging "DocumentTheDatabase"
        blnTest = oDbObjects.DocumentTheDatabase()
    Else
        aeBeginLogging "DocumentTheDatabase", "WithDebugging"
        blnTest = oDbObjects.DocumentTheDatabase("WithDebugging")
    End If

RESULTS:
    Debug.Print "Test: DocumentTheDatabase"
    Debug.Print PassFail(blnTest)
    aeEndLogging "DocumentTheDatabase"

PROC_EXIT:
    Set oDbObjects = Nothing
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_DocumentTheDatabase"
            Stop
    End Select

End Sub

Private Sub Test_Exists(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    Dim oDbObjects As aegit_expClass
    Set oDbObjects = New aegit_expClass

    Dim blnTest As Boolean

    If IsMissing(varDebug) Then
        aeBeginLogging "Exists"
        blnTest = oDbObjects.Exists("Modules", "aegit_expClass")
    Else
        aeBeginLogging "Exists", "WithDebugging"
        blnTest = oDbObjects.Exists("Modules", "aegit_expClass", "WithDebugging")
    End If

RESULTS:
    Debug.Print "Test: Exists"
    Debug.Print PassFail(blnTest)
    aeEndLogging "Exists"

PROC_EXIT:
    Set oDbObjects = Nothing
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_Exists"
            Stop
    End Select

End Sub

Private Sub Test_GetReferences(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    Dim oDbObjects As aegit_expClass
    Set oDbObjects = New aegit_expClass

    Dim blnTest As Boolean

    If IsMissing(varDebug) Then
        aeBeginLogging "GetReferences"
        blnTest = oDbObjects.GetReferences()
    Else
        aeBeginLogging "GetReferences", "WithDebugging"
        blnTest = oDbObjects.GetReferences("WithDebugging")
    End If

RESULTS:
    Debug.Print "Test: GetReferences"
    Debug.Print PassFail(blnTest)
    aeEndLogging "GetReferences"

PROC_EXIT:
    Set oDbObjects = Nothing
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_GetReferences"
            Stop
    End Select

End Sub

Private Sub Test_DocumentTables(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    Dim oDbObjects As aegit_expClass
    Set oDbObjects = New aegit_expClass

    Dim blnTest As Boolean

    If IsMissing(varDebug) Then
        aeBeginLogging "DocumentTables"
        blnTest = oDbObjects.DocumentTables()
    Else
        aeBeginLogging "DocumentTables", "WithDebugging"
        blnTest = oDbObjects.DocumentTables("WithDebugging")
    End If

RESULTS:
    Debug.Print "Test: DocumentTables"
    Debug.Print PassFail(blnTest)
    aeEndLogging "DocumentTables"

PROC_EXIT:
    Set oDbObjects = Nothing
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_DocumentTables"
            Stop
    End Select

End Sub

Private Sub Test_DocumentRelations(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    Dim oDbObjects As aegit_expClass
    Set oDbObjects = New aegit_expClass

    Dim blnTest As Boolean

    If IsMissing(varDebug) Then
        aeBeginLogging "DocumentRelations"
        blnTest = oDbObjects.DocumentRelations()
    Else
        aeBeginLogging "DocumentRelations", "WithDebugging"
        blnTest = oDbObjects.DocumentRelations("WithDebugging")
    End If

RESULTS:
    Debug.Print "Test: DocumentRelations"
    Debug.Print PassFail(blnTest)
    aeEndLogging "DocumentRelations"

PROC_EXIT:
    Set oDbObjects = Nothing
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_DocumentRelations"
            Stop
    End Select

End Sub

Private Sub Test_DocumentTablesXML(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    Dim oDbObjects As aegit_expClass
    Set oDbObjects = New aegit_expClass

    Dim blnTest As Boolean

    If Application.VBE.ActiveVBProject.Name = "aegit" Then
        Dim myArray() As Variant
        myArray = Array("aeItems", "aetlkpStates", "USysRibbons")
        oDbObjects.TablesExportToXML = myArray()
        oDbObjects.ExcludeFiles = False
        Debug.Print , "oDbObjects.ExcludeFiles = " & oDbObjects.ExcludeFiles
    End If

    If IsMissing(varDebug) Then
        aeBeginLogging "DocumentTablesXML"
        blnTest = oDbObjects.DocumentTablesXML()
    Else
        aeBeginLogging "DocumentTablesXML", "WithDebugging"
        blnTest = oDbObjects.DocumentTablesXML("WithDebugging")
    End If

RESULTS:
    Debug.Print "Test: DocumentTablesXML"
    Debug.Print PassFail(blnTest)
    aeEndLogging "DocumentTablesXML"

PROC_EXIT:
    Set oDbObjects = Nothing
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_DocumentTablesXML"
            Stop
    End Select

End Sub

Private Sub Test_SchemaFile(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    Dim oDbObjects As aegit_expClass
    Set oDbObjects = New aegit_expClass

    Dim blnTest As Boolean

    If IsMissing(varDebug) Then
        aeBeginLogging "SchemaFile"
        blnTest = oDbObjects.SchemaFile
    Else
        aeBeginLogging "SchemaFile", "WithDebugging"
        blnTest = oDbObjects.SchemaFile("WithDebugging")
    End If

RESULTS:
    Debug.Print "Test: SchemaFile"
    Debug.Print PassFail(blnTest)
    aeEndLogging "SchemaFile"

PROC_EXIT:
    Set oDbObjects = Nothing
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_SchemaFile"
            Stop
    End Select

End Sub