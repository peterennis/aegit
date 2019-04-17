Option Compare Database
Option Explicit

Public Function RemoveTableDuplicates(ByVal strTableName As String) As Boolean
    ' Author: James Kauffman
    ' Source: http://www.saplsmw.com
    ' Update: Peter F. Ennis
    ' Note a dependency on ADODB plug-in in earlier Access versions.

    On Error GoTo 0
    Dim rs As DAO.Recordset
    Dim nCurrent As Long
    Dim nFieldCount As Long
    Dim nRecordCount As Long
    Dim RetVal As Variant
    Dim nCurRec As Long
    Dim nCurSec As Long
    Dim strLastRecord As String
    Dim strThisRecord As String
    Dim strSQL As String
    Dim nTotalDeleted As Long

    Set rs = CurrentDb.OpenRecordset(strTableName)
    nFieldCount = rs.Fields.Count

    strSQL = "SELECT * FROM " & strTableName & " ORDER BY "

    For nCurrent = 0 To rs.Fields.Count - 1
        strSQL = strSQL & rs.Fields(nCurrent).Name
        If nCurrent < rs.Fields.Count - 1 Then
            strSQL = strSQL & ", "
        End If
    Next
    strSQL = strSQL & ";"
    rs.Close
    Set rs = CurrentDb.OpenRecordset(strSQL)
    
    nRecordCount = rs.RecordCount

    RetVal = SysCmd(acSysCmdInitMeter, "Removing duplicates from " & strTableName & ". . .", nRecordCount)
    Do While Not rs.EOF
        nCurRec = nCurRec + 1
        If Second(Now()) <> nCurSec And nCurRec <> rs.RecordCount Then
            nCurSec = Second(Now())
            RetVal = SysCmd(acSysCmdUpdateMeter, nCurRec)
            RetVal = DoEvents()
        End If
        
        strThisRecord = vbNullString
        For nCurrent = 0 To rs.Fields.Count - 1
            strThisRecord = strThisRecord & rs.Fields(nCurrent).Value
        Next
        If strThisRecord = strLastRecord Then
            rs.Delete
            nTotalDeleted = nTotalDeleted + 1
        End If
        strLastRecord = strThisRecord
        rs.MoveNext
    Loop
    rs.Close
    RemoveTableDuplicates = True
    RetVal = SysCmd(acSysCmdRemoveMeter)

End Function

Public Function ExportToText(ByVal strTableName As String, ByVal strFileName As String, Optional ByVal strDelim As String = vbTab) As Boolean
    ' This function ONLY exports to Tab-delimited text files with the headers and without text idenitifiers (No quotes!)
    
    On Error GoTo 0
    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim nCurrent As Long
    Dim nFieldCount As Long
    Dim nRecordCount As Long
    Dim RetVal As Variant
    Dim nCurRec As Long
    Dim nCurSec As Long
    Dim strTest As String
    
    strSQL = "SELECT * FROM " & strTableName & ";"
    
    ' Check to see if strTableName is actually a query.  If so, use its SQL query.
    nCurrent = 0
    Do While nCurrent < CurrentDb.QueryDefs.Count
        If UCase$(CurrentDb.QueryDefs(nCurrent).Name) = UCase$(strTableName) Then
            strSQL = CurrentDb.QueryDefs(nCurrent).sql
        End If
        nCurrent = nCurrent + 1
    Loop
    
    Set rst = CurrentDb.OpenRecordset(strSQL)
    nFieldCount = rst.Fields.Count
    
    If Not rst.EOF Then
        ' Now find the *actual* record count--returns a value of 1 record if we don't do these moves.
        rst.MoveLast
        rst.MoveFirst
    End If
    nRecordCount = rst.RecordCount

    RetVal = SysCmd(acSysCmdInitMeter, "Exporting " & strTableName & " to " & strFileName & ". . .", nRecordCount)

    Open strFileName For Output As #1
    For nCurrent = 0 To nFieldCount - 1
        If Right$(rst.Fields(nCurrent).Name, 1) = "_" Then
            Print #1, Left$(rst.Fields(nCurrent).Name, Len(rst.Fields(nCurrent).Name) - 1) & strDelim;
        Else
            Print #1, rst.Fields(nCurrent).Name & strDelim;
        End If
    Next
    Print #1, vbNullString        ' New line.
    nCurSec = Second(Now())
    Do While nCurSec = Second(Now())
    Loop
    nCurSec = Second(Now())
    Do While Not rst.EOF
        nCurRec = nCurRec + 1
        If Second(Now()) <> nCurSec And nCurRec <> rst.RecordCount Then
            nCurSec = Second(Now())
            RetVal = SysCmd(acSysCmdUpdateMeter, nCurRec)
            RetVal = DoEvents()
        End If
        strTest = vbNullString
        For nCurrent = 0 To nFieldCount - 1  'Check for blank lines--no need to export those!
            If IsNull(rst.Fields) Then
                strTest = strTest & vbNullString
            Else
                strTest = strTest & rst.Fields(nCurrent).Value
            End If
        Next
        If Len(Trim$(strTest)) > 0 Then  'Check for blank lines--no need to export those!
            For nCurrent = 0 To nFieldCount - 1
                If Not IsNull(rst.Fields(nCurrent).Value) Then
                    Print #1, Trim$(rst.Fields(nCurrent));
                End If
                If nCurrent < nFieldCount - 1 Then
                    Print #1, strDelim;
                Else       'new line.
                    Print #1, vbNullString
                End If
            Next
        End If
        rst.MoveNext
    Loop
    Close #1
    rst.Close
    Set rst = Nothing
    ExportToText = True
    RetVal = SysCmd(acSysCmdRemoveMeter)

End Function

Public Sub TestExportToTextUnicode()
    On Error GoTo 0
    Dim bln As Boolean
    bln = ExportToTextUnicode("Items", "C:\Temp\ExportedItemsUnicode.txt")
End Sub

Public Function ExportToTextUnicode(ByVal strTableName As String, ByVal strFileName As String, Optional ByVal strDelim As String = vbTab) As Boolean
    ' Written by Jimbo at SAPLSMW.com
    ' Special thanks: accessblog.net/2007/06/how-to-write-out-unicode-text-files-in.html

    On Error GoTo 0
    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim nCurrent As Long
    Dim nFieldCount As Long
    Dim nRecordCount As Long
    Dim RetVal As Variant
    Dim nCurRec As Long
    Dim nCurSec As Long
    Dim strTest As String

    strSQL = "SELECT * FROM " & strTableName & ";"

    ' Check to see if strTableName is actually a query.  If so, use its SQL query.
    nCurrent = 0
    Do While nCurrent < CurrentDb.QueryDefs.Count
        If UCase$(CurrentDb.QueryDefs(nCurrent).Name) = UCase$(strTableName) Then
            strSQL = CurrentDb.QueryDefs(nCurrent).sql
        End If
        nCurrent = nCurrent + 1
    Loop
    Set rst = CurrentDb.OpenRecordset(strSQL)
    nFieldCount = rst.Fields.Count

    If Not rst.EOF Then
        ' Now find the *actual* record count--returns a value of 1 record if we don't do these moves.
        rst.MoveLast
        rst.MoveFirst
    End If

    nRecordCount = rst.RecordCount
    RetVal = SysCmd(acSysCmdInitMeter, "Exporting " & strTableName & " to " & strFileName & ". . .", nRecordCount)
    'Create a binary stream
    Dim UnicodeStream As Object
    Set UnicodeStream = CreateObject("ADODB.Stream")
    UnicodeStream.Charset = "UTF-8"
    UnicodeStream.Open

    For nCurrent = 0 To nFieldCount - 1
        If Right$(rst.Fields(nCurrent).Name, 1) = "_" Then
            UnicodeStream.WriteText Left$(rst.Fields(nCurrent).Name, Len(rst.Fields(nCurrent).Name) - 1) & strDelim
        Else
            UnicodeStream.WriteText rst.Fields(nCurrent).Name & strDelim
        End If
    Next

    UnicodeStream.WriteText vbCrLf
    nCurSec = Second(Now())

    Do While Not rst.EOF
        nCurRec = nCurRec + 1
        If Second(Now()) <> nCurSec And nCurRec <> rst.RecordCount Then
            nCurSec = Second(Now())
            RetVal = SysCmd(acSysCmdUpdateMeter, nCurRec)
            RetVal = DoEvents()
        End If
        strTest = vbNullString
        For nCurrent = 0 To nFieldCount - 1  ' Check for blank lines--no need to export those!
            If IsNull(rst.Fields) Then
                strTest = strTest & vbNullString
            Else
                strTest = strTest & rst.Fields(nCurrent).Value
            End If
        Next
        If Len(Trim$(strTest)) > 0 Then  ' Check for blank lines--no need to export those!
            For nCurrent = 0 To nFieldCount - 1
                If Not IsNull(rst.Fields(nCurrent).Value) Then
                    UnicodeStream.WriteText Trim$(rst.Fields(nCurrent).Value)
                End If
                If nCurrent = (nFieldCount - 1) Then
                    UnicodeStream.WriteText vbCrLf 'new line.
                Else
                    UnicodeStream.WriteText strDelim
                End If
            Next
        End If
        rst.MoveNext
    Loop

    ' Check to ensure that the file doesn't already exist.
    If Len(Dir$(strFileName)) > 0 Then
        Kill strFileName  ' The file exists, so we must delete it before it be created again.
    End If
    UnicodeStream.SaveToFile strFileName
    UnicodeStream.Close
    rst.Close
    Set rst = Nothing
    ExportToTextUnicode = True
    RetVal = SysCmd(acSysCmdRemoveMeter)

End Function

Public Function ImportFromAccess(ByVal strSourceFile As String, ByVal strSourceTable As String, _
    ByVal strTargetTable As String) As Boolean

    On Error GoTo 0
    Dim nCurrent As Long
    Dim nRecordCount As Long
    Dim nFileLen As Long
    Dim nTotalBytes As Long
    Dim RetVal As Variant
    Dim nCurRec As Long
    Dim nCurSec As Long

    Dim dbs As DAO.Database
    Set dbs = OpenDatabase(strSourceFile)

    Dim rs1 As DAO.Recordset
    Set rs1 = dbs.OpenRecordset(strSourceTable)

    Dim rs As DAO.Recordset
    rs.OpenRecordset (strTargetTable)

    nRecordCount = rs1.RecordCount

    RetVal = SysCmd(acSysCmdInitMeter, "Importing " & strSourceTable & " from " & strSourceFile & "...", nFileLen)

    Do While Not rs1.EOF
        nCurRec = nCurRec + 1
        If Second(Now()) <> nCurSec Then ' And nCurRec <> rs.RecordCount Then
            nCurSec = Second(Now())
            RetVal = SysCmd(acSysCmdUpdateMeter, nTotalBytes)
            RetVal = DoEvents()
        End If
        rs.AddNew
        nCurrent = 0
        Do While nCurrent < rs1.Fields.Count
            rs.Fields(nCurrent).Value = rs1.Fields(nCurrent).Value
            nCurrent = nCurrent + 1
            rs.Update
        Loop
        rs1.MoveNext
    Loop
    rs.Close
    rs1.Close
    dbs.Close
    RetVal = SysCmd(acSysCmdRemoveMeter)
    ImportFromAccess = True

End Function

Public Function TableScrub(ByVal strTableName As String) As Long
    ' This function removes leading spaces and trailing spaces from every string field in a table.

    On Error GoTo 0
    Dim strTemp As String
    Dim A As Integer
    Dim nLength As Long
    Dim RetVal As Variant
    Dim nCurRec As Long
    Dim nCurSec As Integer
    Dim nTotalSeconds As Integer
    Dim nSecondsLeft As Integer
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset(strTableName)

    nCurSec = Second(Now())
    TableScrub = 0
    RetVal = SysCmd(acSysCmdInitMeter, "Killing excess spaces in " & strTableName & " . . . ", rs.RecordCount)

    rs.MoveFirst
    Do While Not rs.EOF
        nCurRec = nCurRec + 1
        If Second(Now()) <> nCurSec And nCurRec <> rs.RecordCount Then
            nTotalSeconds = nTotalSeconds + 1
            If nTotalSeconds > 3 Then
                nSecondsLeft = Int(((nTotalSeconds / nCurRec) * rs.RecordCount) * ((rs.RecordCount - nCurRec) / rs.RecordCount))
                RetVal = SysCmd(acSysCmdRemoveMeter)
                RetVal = SysCmd(acSysCmdInitMeter, "Killing excess spaces in " & strTableName & ", " & nSecondsLeft & " seconds remaining.", rs.RecordCount())
                RetVal = SysCmd(acSysCmdUpdateMeter, nCurRec)
                RetVal = DoEvents()
            End If
            nCurSec = Second(Now())
        End If
        rs.Edit
        For A = 0 To rs.Fields.Count - 1
            nLength = 0
            If rs.Fields(A).Type = dbText And Len(rs.Fields(A).Value) > 0 Then
                nLength = Len(rs.Fields(A).Value)
                strTemp = Trim$(rs.Fields(A).Value)
                If Len(strTemp) = 0 Then
                    rs.Fields(A).Value = Null
                Else
                    rs.Fields(A).Value = strTemp
                End If
                nLength = nLength - Len(strTemp)
            End If
            TableScrub = TableScrub + nLength

        Next
        rs.Update
        rs.MoveNext
    Loop
    RetVal = SysCmd(acSysCmdRemoveMeter)
    rs.Close
    Set rs = Nothing

End Function

Public Function FixCase(ByVal strText As String) As String
    ' Convert to sentence case: UPPER CASE COMPANY NAME-->Upper Case Company Name
    Dim strParse As String
    On Error GoTo 0
    strParse = Trim$(strText & vbNullString)
    Dim nCurrent As Long
    For nCurrent = 2 To Len(strParse)
        If Mid$(strParse, nCurrent - 1, 1) <> " " And Mid$(strParse, nCurrent - 1, 1) <> "." Then
            strParse = Left$(strParse, nCurrent - 1) & LCase$(Mid$(strParse, nCurrent, 1)) & Mid$(strParse, nCurrent + 1)
        End If
    Next
    FixCase = strParse
End Function

Public Function Deduplicate(ByVal strValue As String) As Boolean
    On Error GoTo 0
    Static sValue As String
    If strValue = sValue Then
        Deduplicate = True
    Else
        Deduplicate = False
        sValue = strValue
    End If
End Function

Public Function DeleteRecords(ByVal strTableName As String) As Boolean
    ' Delete all records from a table--easier than creating a delete query.
    On Error GoTo 0
    CurrentDb.Execute ("DELETE * FROM " & strTableName)
    DeleteRecords = True
End Function