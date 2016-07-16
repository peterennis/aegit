Option Compare Database
Option Explicit

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

Public Sub TestLogging()
    On Error GoTo 0
    aeBeginLogging "MyProc", "Some Parameter varOne", "varTwo", "varThree"
    ' more code here
    aePrintLog "Some Status"
    ' more code here
    aeEndLogging
End Sub

Private Sub aeBeginLogging(ByVal strProcName As String, Optional ByVal varOne As Variant = vbNullString, _
        Optional ByVal varTwo As Variant = vbNullString, Optional ByVal varThree As Variant = vbNullString)
    On Error GoTo 0
    mlngStartTime = timeGetTime()
    Debug.Print ">aeBeginLogging"; Space$(1); "mlngStartTime=" & mlngStartTime
    If aeLog.blnNoTrace Then
        Debug.Print "B1: aeBeginLogging", "blnNoTrace=" & aeLog.blnNoTrace
        Exit Sub
    End If
    If Not aeLog.blnNoTimer Then
        Debug.Print "B2: aeBeginLogging", "blnNoTimer=" & aeLog.blnNoTimer
        Debug.Print Format$(mlngStartTime, "0.00"); Space$(2);
    End If
    Debug.Print Space$(lngIndent * 4); strProcName; Space$(1); "'" & varOne & "'"; Space$(1); "'" & varTwo & "'"; Space$(1); "'" & varThree & "'"
    lngIndent = lngIndent + 1
End Sub

Private Sub aeEndLogging(Optional ByVal varOne As Variant = vbNullString, _
        Optional ByVal varTwo As Variant = vbNullString, Optional ByVal varThree As Variant = vbNullString)
    On Error GoTo 0
    If aeLog.blnNoTrace Then
        Debug.Print "E1: aeEndLogging", "blnNoTrace=" & aeLog.blnNoTrace
        Exit Sub
    End If
    lngIndent = lngIndent - 1
    mlngEndTime = timeGetTime()
    If Not aeLog.blnNoEnd Then
        If Not aeLog.blnNoTimer Then
            Debug.Print ">aeEndLogging"; Space$(1); "mlngEndTime=" & mlngEndTime
            mlngEndTime = timeGetTime()
            Debug.Print "E2: aeEndLogging", "blnNoTimer=" & aeLog.blnNoTimer
            Debug.Print Format$(mlngEndTime, "0.00"); Space$(2);
        End If
        Debug.Print Space$(lngIndent * 4); "End " & lngIndent; Space$(1); varOne; Space$(1); varTwo; Space$(1); varThree
        Debug.Print "It took " & (mlngEndTime - mlngStartTime) / 1000 & " seconds to process this procedure"
    End If
End Sub

Public Sub aePrintLog(Optional ByVal varOne As Variant = vbNullString, _
        Optional ByVal varTwo As Variant = vbNullString, Optional ByVal varThree As Variant = vbNullString)
    On Error GoTo 0
    If aeLog.blnNoTrace Or aeLog.blnNoPrint Then
        Exit Sub
    End If
    If Not aeLog.blnNoTimer Then
        Debug.Print Format$(Timer, "0.00"); Space$(2);
    End If
    Debug.Print Space$(lngIndent * 4); "'" & varOne & "'"; Space$(1); "'" & varTwo; "'"; Space$(1); "'" & varThree; "'"
End Sub

Public Function IsMacHidden(ByVal strMacroName As String) As Boolean
    On Error GoTo 0
    If IsNull(strMacroName) Or strMacroName = vbNullString Then
        IsMacHidden = False
        'Debug.Print "IsMacHidden Null Test", strMacroName, IsMacHidden
    Else
        IsMacHidden = GetHiddenAttribute(acMacro, strMacroName)
        'Debug.Print "IsMacHidden Attribute Test", strMacroName, IsMacHidden
    End If
End Function

Public Sub NoBOM(ByVal strFileName As String)
' Ref: http://www.experts-exchange.com/Programming/Languages/Q_27478996.html
' Use the same file name for input and output
    On Error GoTo 0

    ' Define needed constants
'    Const ForReading As Integer = 1
    Const ForWriting As Integer = 2
'    Const TriStateUseDefault As Integer = -2
    Const adTypeText As Integer = 2
    Dim strContent As String

    ' Convert UTF-8 file to ANSI file
    Dim objStreamFile As Object
    Set objStreamFile = CreateObject("Adodb.Stream")
    With objStreamFile
        .Charset = "UTF-8"
        .Type = adTypeText
        .Open
        .LoadFromFile strFileName
        strContent = .ReadText
        .Close
    End With
    Set objStreamFile = Nothing
    Kill strFileName
    'Stop

    DoEvents

    ' Write out after "conversion"
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFile As Object
    Set objFile = objFSO.OpenTextFile(strFileName, ForWriting, True)
    objFile.Write Right$(strContent, Len(strContent) - 2)
    objFile.Close

    Set objFile = Nothing

End Sub

'?strTran("Dev" & vbCrLf & "ashish", vbCrLf, " ")
'Dev ashish
'
'Function strTran(ByVal sInString As String, _
                            sFindString As String, _
                            sReplaceString As String, _
                            Optional iCount As Variant) As String
'   Dim iSpot As Integer, iCtr As Integer
'   If IsMissing(iCount) Then iCount = 1
'   If iCount = 0 Then iCount = 1000
'   For iCtr = 1 To iCount
'     iSpot = InStr(1, sInString, sFindString)
'     If iSpot > 0 Then
'       sInString = Left$(sInString, iSpot - 1) & _
'                         sReplaceString & _
'                         Mid$(sInString, iSpot + Len(sFindString))
'     Else
'       Exit For
'     End If
'   Next
'   strTran = sInString
'
'End Function
'
'http://computer-programming-forum.com/1-vba/34d339bb6472eb9d.htm

Public Sub CatalogUserCreatedObjects()
' Ref: http://blogannath.blogspot.com/2010/03/microsoft-access-tips-tricks-working.html#ixzz3WCBJcxwc
' Ref: http://stackoverflow.com/questions/5286620/saving-a-query-via-access-vba-code
    On Error GoTo 0

    Dim strSQL As String
    Const MY_QUERY_NAME As String = "zzzqryCatalogUserCreatedObjects"

    strSQL = "SELECT IIf(type = 1,""Table"", IIf(type = 6, ""Linked Table"", "
    strSQL = strSQL & vbCrLf & "IIf(type = 5,""Query"", IIf(type = -32768,""Form"", "
    strSQL = strSQL & vbCrLf & "IIf(type = -32764,""Report"", IIf(type=-32766,""Module"", "
    strSQL = strSQL & vbCrLf & "IIf(type = -32761,""Module"", ""Unknown""))))))) as [Object Type], "
    strSQL = strSQL & vbCrLf & "MSysObjects.Name, MSysObjects.DateCreate, MSysObjects.DateUpdate "
    strSQL = strSQL & vbCrLf & "FROM MSysObjects "
    strSQL = strSQL & vbCrLf & "WHERE Type IN (1, 5, 6, -32768, -32764, -32766, -32761) "
    strSQL = strSQL & vbCrLf & "AND Left$(Name, 4) <> ""MSys"" AND Left$(Name, 1) <> ""~"" "
    strSQL = strSQL & vbCrLf & "ORDER BY IIf(type=1,""Table"",IIf(type=6,""Linked Table"",IIf(type=5,""Query"",IIf(type=-32768,""Form"",IIf(type=-32764,""Report"",IIf(type=-32766,""Module"",IIf(type=-32761,""Module"",""Unknown""))))))), MSysObjects.Name;"

    Debug.Print strSQL
    
    ' Using a query name and sql string, if the query does not exist, ...
    If IsNull(DLookup("Name", "MsysObjects", "Name='" & MY_QUERY_NAME & "'")) Then
        ' create it ...
        CurrentDb.CreateQueryDef MY_QUERY_NAME, strSQL
    Else
        ' other wise, update the sql
        CurrentDb.QueryDefs(MY_QUERY_NAME).SQL = strSQL
    End If

    DoCmd.OpenQuery MY_QUERY_NAME

End Sub