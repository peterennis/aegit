Option Compare Database
Option Explicit

Public Sub TestImportLogSummaryFromText()

    Dim bln As Boolean

    bln = ImportLogSummaryFromText("tblLogSummary", "log_summary.csv", ",")

End Sub

Private Function ImportLogSummaryFromText(ByVal strTableName As String, _
    ByVal strFileName As String, Optional ByVal strDelim As String = vbTab) As Boolean
    'Ref:  James Kauffman, http://www.saplsmw.com

    On Error GoTo 0

    Dim rs As DAO.Recordset
    Dim nCurrent As Long
''    Dim RetVal As Variant
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

    lngTotalBytes = lngTotalBytes + Len(strTest) + 2 ' +2 for vbCrLf--This line prevents div by zero later...
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
    Do While Not EOF(1) And lngRecords < nReadAhead ' Read through the file and get the maximum sizes for fields in advance.
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
        Set rs = CurrentDb.OpenRecordset(strTableName)
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
            lngTotalBytes = lngTotalBytes + Len(strTest) + 2 'vbCrLf
            strTest = Trim$(strTest)
            If Right$(strTest, 1) = strDelim Then
                strTest = Left$(strTest, Len(strTest) - 1)
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
    Debug.Print "DONE!!!"

End Function