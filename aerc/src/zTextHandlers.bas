Option Compare Database
Option Explicit

' Author: James Kauffman
' Source: http://www.saplsmw.com
' Update: Peter F. Ennis

'Note a dependency on ADODB plug-in in earlier Access versions.

Public Function RemoveTableDuplicates(ByVal strTableName As String) As Boolean

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

Public Function ImportFromText(ByVal strTableName As String, ByVal strFileName As String, Optional ByVal strDelim As String = vbTab) As Boolean
' This function should be used only for importing extraordinarily large text files.
' Files of normal length should be imported using the Access import utility.
  
    On Error GoTo 0
    Dim rs As DAO.Recordset
    Dim nCurrent As Long
    Dim RetVal As Variant
    Dim nCurRec As Long
    Dim nCurSec As Long
    Dim nTotalSeconds As Long
    Dim nSecondsLeft As Long
    Dim nTotalBytes As Long
    Dim nFileLen As Long
    Dim strTest As Variant
    Dim strHeadersIn() As String
    Dim strHeaders(999) As String
    Const nReadAhead As Long = 30000
    Dim nSizes(999) As Long
    Dim strRecords(nReadAhead) As String
    Dim nRecords As Long
    Dim nLoaded As Long
    Dim strFields() As String
    Dim nHeaders As Long
    Dim isSAP As Boolean

    nFileLen = FileLen(strFileName)
    RetVal = SysCmd(acSysCmdSetStatus, "Preparing to import " & strTableName & " from " & strFileName & "...")
    RetVal = DoEvents()

    Open strFileName For Input As #1
    Line Input #1, strTest
    If Left$(strTest, 6) = "Table:" Then ' This is an SAP extract!
        isSAP = True
        Line Input #1, strTest
        Line Input #1, strTest
        Line Input #1, strTest  ' Fourth line has the headers!
    Else
        isSAP = False
    End If

    If InStr(1, strTest, "|", vbTextCompare) Then
        strDelim = "|"
    End If

    nTotalBytes = nTotalBytes + Len(strTest) + 2 ' +2 for vbCrLf--This line prevents div by zero later...
    strTest = Trim$(strTest)
    If Right$(strTest, 1) = strDelim Then
        strTest = Left$(strTest, Len(strTest) - 1)
    End If
    strHeadersIn = Split(Trim$(strTest), strDelim)
    nHeaders = 0
    
    For Each strTest In strHeadersIn
        nHeaders = nHeaders + 1
        strTest = Replace(Replace(strTest, " ", vbNullString), ".", vbNullString)
        strTest = Replace(Replace(strTest, " ", vbNullString), ".", vbNullString)
        If Len(Trim$(strTest)) = 0 Then
            strHeaders(nHeaders) = "HEADER" & Right$("000" & nHeaders, 3)
        Else
            strHeaders(nHeaders) = Trim$(strTest)
        End If
        For nCurrent = 1 To nHeaders - 1
            If strHeaders(nHeaders) = strHeaders(nCurrent) Then
                strHeaders(nHeaders) = strHeaders(nHeaders) & nHeaders
            End If
        Next
    Next
    strHeaders(0) = nHeaders
    RetVal = SysCmd(acSysCmdClearStatus)
    RetVal = SysCmd(acSysCmdInitMeter, "Preparing to import " & strTableName & " from " & strFileName & "...", nReadAhead)
    RetVal = DoEvents()

    Do While Not EOF(1) And nRecords < nReadAhead ' Read through the file and get the maximum sizes for fields in advance.
        Line Input #1, strTest
        strTest = Trim$(strTest)
        If Right$(strTest, 1) = strDelim Then
            strTest = Left$(strTest, Len(strTest) - 1)
        End If
        If isSAP And Left$(strTest, 20) = "--------------------" Then
            strTest = vbNullString  ' Skip this line!
        End If
        If Len(strTest) > 0 Then
            nRecords = nRecords + 1
            strRecords(nRecords) = strTest
            strFields = Split(strTest, strDelim)
            nCurrent = 0
            For Each strTest In strFields
                nCurrent = nCurrent + 1
                If Len(strTest) > nSizes(nCurrent) Then
                    nSizes(nCurrent) = Len(strTest)
                End If
            Next
            If Second(Now) <> nCurSec Then
                nCurSec = Second(Now)
                RetVal = SysCmd(acSysCmdUpdateMeter, nRecords)
                RetVal = DoEvents()
            End If
        End If
    Loop
   
    If CreateTable(strTableName, strHeaders, nSizes) Then
        If isSAP Then
            For nCurrent = 1 To nHeaders
                If Left$(strHeaders(nCurrent), 8) = "HEADER00" Then
                    strHeaders(nCurrent) = vbNullString  ' Don't bother importing this field.
                End If
            Next
        End If
        Set rs = CurrentDb.OpenRecordset(strTableName)
        nLoaded = 0
        nTotalSeconds = 0
        Do While Not EOF(1) Or nLoaded < nRecords
            nCurRec = nCurRec + 1
            If Second(Now()) <> nCurSec Then
                nCurSec = Second(Now())
                nTotalSeconds = nTotalSeconds + 1
                'RetVal = DoEvents()
                If nTotalSeconds > 3 Then
                    'nSecondsLeft = Int(((nTotalSeconds / nCurRec) * rs.RecordCount) * ((rs.RecordCount - nCurRec) / rs.RecordCount))
                    nSecondsLeft = Int(((nTotalSeconds / nTotalBytes) * nFileLen) * ((nFileLen - nTotalBytes) / nFileLen))
                    RetVal = SysCmd(acSysCmdRemoveMeter)
                    RetVal = SysCmd(acSysCmdInitMeter, "Importing " & strTableName & " from " & strFileName & "... " & nSecondsLeft & " seconds remaining.", nFileLen)
                    RetVal = SysCmd(acSysCmdUpdateMeter, nTotalBytes)
                    RetVal = DoEvents()
                End If
            End If
            If nLoaded < nRecords Then
                nLoaded = nLoaded + 1
                strTest = strRecords(nLoaded)
            Else
                Line Input #1, strTest
            End If
            nTotalBytes = nTotalBytes + Len(strTest) + 2 'vbCrLf
            strTest = Trim$(strTest)
            If Right$(strTest, 1) = strDelim Then
                strTest = Left$(strTest, Len(strTest) - 1)
            End If
            If isSAP And Left$(strTest, 20) = "--------------------" Then
                strTest = vbNullString  ' Skip this line!
            End If
            If Len(strTest) > 0 Then
                strFields = Split(strTest, strDelim)
                nCurrent = 0
                rs.AddNew
                For Each strTest In strFields
                    nCurrent = nCurrent + 1
                    If Len(Trim$(strHeaders(nCurrent))) > 0 Then
                        rs.Fields(strHeaders(nCurrent)).Value = Trim$(strFields(nCurrent - 1))
                    End If
                Next
                rs.Update
            End If
        Loop
        rs.Close
    End If
    Close #1
    RetVal = SysCmd(acSysCmdRemoveMeter)

End Function

Public Function CreateTable(ByVal strTableName As String, ByRef strFields() As String, ByRef nSizes() As Long) As Boolean

    Dim nCounter As Long
    Dim dbs As DAO.Database
    ' Now create the database.  Rename the old database if necessary.
    Set dbs = CurrentDb
    Dim tdf As DAO.TableDef
    Dim fld1 As DAO.Field
    Dim fName As String
    Dim fType As Integer
    Dim fSize As Integer

    On Error GoTo ErrorHandler
    ' Check for existence of TargetTable
    nCounter = 0
    Do While nCounter < dbs.TableDefs.Count
        If dbs.TableDefs(nCounter).Name = strTableName Then
            ' Delete TargetTable--must start from scratch
            dbs.TableDefs.Delete (strTableName)
        End If
        nCounter = nCounter + 1
    Loop
    
    Set tdf = dbs.CreateTableDef(strTableName)
    For nCounter = 1 To Val(strFields(0))
        fName = strFields(nCounter)
        fType = dbText
        fSize = nSizes(nCounter) 'fSize = 255
        Set fld1 = tdf.CreateField(fName, fType, fSize)
        fld1.AllowZeroLength = True
        fld1.Required = False
        tdf.Fields.Append fld1
    Next
    ' Create the table in the database
    dbs.TableDefs.Append tdf
    dbs.TableDefs.Refresh
    CreateTable = True
    Exit Function

ErrorHandler:
    MsgBox "Error number " & Err.Number & ": " & Err.Description
    CreateTable = False
    Exit Function

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
    On Error GoTo 0
    strText = Trim$(strText & vbNullString)
    Dim nCurrent As Long
    For nCurrent = 2 To Len(strText)
        If Mid$(strText, nCurrent - 1, 1) <> " " And Mid$(strText, nCurrent - 1, 1) <> "." Then
            strText = Left$(strText, nCurrent - 1) & LCase$(Mid$(strText, nCurrent, 1)) & Mid$(strText, nCurrent + 1)
        End If
    Next
    FixCase = strText
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