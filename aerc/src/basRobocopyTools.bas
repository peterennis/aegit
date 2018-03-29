Option Compare Database
Option Explicit

Public Sub TestImportLogSummaryFromText()

    Dim bln As Boolean

    bln = ImportLogSummaryFromText("tblLogSummary", "doc\log_summary.csv", ",")

End Sub

Private Function ImportLogSummaryFromText(ByVal strTableName As String, _
    ByVal strFileName As String, Optional ByVal strDelim As String = vbTab) As Boolean
    ' Ref:  James Kauffman, http://www.saplsmw.com

    On Error GoTo 0

    Dim rst As DAO.Recordset
    Dim nCurrent As Long
    Dim lngCurRec As Long
    Dim lngCurSec As Long
    Dim lngTotalSeconds As Long
    Dim lngSecondsLeft As Long
    Dim lngTotalBytes As Long
    Dim lngFileLen As Long
    Dim strTest As Variant
    Dim strHeadersIn() As String
    Dim strHeaders(999) As String
    Const nReadAhead As Long = 30000
    Dim nSizes(999) As Long
    Dim strRecords(nReadAhead) As String
    Dim lngRecords As Long
    Dim lngLoaded As Long
    Dim strFields() As String
    Dim lngHeaders As Long
    Dim strFileNamePath As String

    Debug.Print "strFileName = " & strFileName
    strFileNamePath = CurrentProject.Path & "\" & strFileName
    Debug.Print "strFileNamePath = " & strFileNamePath
    lngFileLen = FileLen(strFileNamePath)

    Open strFileNamePath For Input As #1
    Line Input #1, strTest

    lngTotalBytes = lngTotalBytes + Len(strTest) + 2    ' +2 for vbCrLf--This line prevents div by zero later...
    strTest = Trim$(strTest)
    If Right$(strTest, 1) = strDelim Then
        strTest = Left$(strTest, Len(strTest) - 1)
    End If
    strHeadersIn = Split(Trim$(strTest), strDelim)
    lngHeaders = 0

    Debug.Print "A"
    For Each strTest In strHeadersIn
        lngHeaders = lngHeaders + 1
        strTest = Replace(Replace(strTest, " ", vbNullString), ".", vbNullString)
        strTest = Replace(Replace(strTest, " ", vbNullString), ".", vbNullString)
        If Len(Trim$(strTest)) = 0 Then
            strHeaders(lngHeaders) = "HEADER" & Right$("000" & lngHeaders, 3)
        Else
            strHeaders(lngHeaders) = Trim$(strTest)
        End If
        For nCurrent = 1 To lngHeaders - 1
            If strHeaders(lngHeaders) = strHeaders(nCurrent) Then
                strHeaders(lngHeaders) = strHeaders(lngHeaders) & lngHeaders
            End If
        Next
    Next
    strHeaders(0) = lngHeaders

    Debug.Print "B"
    Do While Not EOF(1) And lngRecords < nReadAhead     ' Read through the file and get the maximum sizes for fields in advance.
        Line Input #1, strTest
        strTest = Trim$(strTest)
        If Right$(strTest, 1) = strDelim Then
            strTest = Left$(strTest, Len(strTest) - 1)
        End If
        If Len(strTest) > 0 Then
            lngRecords = lngRecords + 1
            strRecords(lngRecords) = strTest
            strFields = Split(strTest, strDelim)
            nCurrent = 0
            For Each strTest In strFields
                nCurrent = nCurrent + 1
                If Len(strTest) > nSizes(nCurrent) Then
                    nSizes(nCurrent) = Len(strTest)
                End If
            Next
            If Second(Now) <> lngCurSec Then
                lngCurSec = Second(Now)
            End If
        End If
    Loop

    Debug.Print "C"
    If CreateTable(strTableName, strHeaders, nSizes) Then
        Set rst = CurrentDb.OpenRecordset(strTableName)
        lngLoaded = 0
        lngTotalSeconds = 0
        Do While Not EOF(1) Or lngLoaded < lngRecords
            lngCurRec = lngCurRec + 1
            If Second(Now()) <> lngCurSec Then
                lngCurSec = Second(Now())
                lngTotalSeconds = lngTotalSeconds + 1
                If lngTotalSeconds > 3 Then
                    'lngSecondsLeft = Int(((lngTotalSeconds / lngCurRec) * rs.RecordCount) * ((rs.RecordCount - lngCurRec) / rs.RecordCount))
                    lngSecondsLeft = Int(((lngTotalSeconds / lngTotalBytes) * lngFileLen) * ((lngFileLen - lngTotalBytes) / lngFileLen))
                End If
            End If
            If lngLoaded < lngRecords Then
                lngLoaded = lngLoaded + 1
                strTest = strRecords(lngLoaded)
            Else
                Line Input #1, strTest
            End If
            lngTotalBytes = lngTotalBytes + Len(strTest) + 2 ' vbCrLf
            strTest = Trim$(strTest)
            If Right$(strTest, 1) = strDelim Then
                strTest = Left$(strTest, Len(strTest) - 1)
            End If
            If Len(strTest) > 0 Then
                strFields = Split(strTest, strDelim)
                nCurrent = 0
                rst.AddNew
                For Each strTest In strFields
                    nCurrent = nCurrent + 1
                    If Len(Trim$(strHeaders(nCurrent))) > 0 Then
                        rst.Fields(strHeaders(nCurrent)).Value = Trim$(strFields(nCurrent - 1))
                    End If
                Next
                rst.Update
            End If
        Loop
        rst.Close
    End If
    Close #1
    Debug.Print "DONE!!!"

End Function

Private Function CreateTable(ByVal strTableName As String, ByRef strFields() As String, ByRef nSizes() As Long) As Boolean

    Dim nCounter As Long
    Dim dbs As DAO.Database
    ' Now create the database.  Rename the old database if necessary.
    Set dbs = CurrentDb
    Dim tdf As DAO.TableDef
    Dim fld1 As DAO.Field
    Dim fName As String
    Dim fType As Integer
    Dim fSize As Integer

    On Error GoTo PROC_ERR
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

PROC_ERR:
    MsgBox "Error number " & Err.Number & ": " & Err.Description, "Function CreateTable"
    CreateTable = False
    Exit Function

End Function