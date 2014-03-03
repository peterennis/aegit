'* * * * * * * * * * * * * * * * * * * *
'*                                     *  +--------------------------+
'*      Written by James Kauffman      *  |                          |
'*                                     *  |  http://www.saplsmw.com  |
'*     Ver 1.20 Updated 17Jun2010      *  |                          |
'*                                     *  +--------------------------+
'* * * * * * * * * * * * * * * * * * * *

Option Compare Database

Function ExportToExcel(strTableName, strFileName, Optional strTabName As String = "Sheet1") As Boolean

    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    '*                                                         *
    '*   Requires reference to Microsoft Excel Object Library  *
    '*                                                         *
    '*   Tools->References...->Microsoft Excel Ojbect Library  *
    '*                                                         *
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    Dim wb As Object
    Set wb = objExcel.Workbook
    Dim ws As Object
    Set ws = objExcel.Worksheet
    Dim ws2 As Object
    Set ws2 = objExcel.Worksheet
    Dim nCurrent As Long, isFound As Boolean

    On Error Resume Next
    isFound = False
    If Len(Dir(strFileName, vbNormal)) > 0 Then 'File exists!
        Set wb = objExcel.Workbooks.Open(strFileName, False, False)
        nCurrent = wb.Worksheets.Count
        Do While nCurrent > 0
            If wb.Worksheets(nCurrent).Name = strTabName Then
                objExcel.DisplayAlerts = False
                Set ws2 = wb.Worksheets(nCurrent)
                Set ws = wb.Worksheets.Add
                ws2.Delete
                ws.Name = strTabName
                isFound = True
            End If
            nCurrent = nCurrent - 1
        Loop
        If Not isFound Then
            Set ws = wb.Worksheets.Add
            ws.Name = strTabName
        End If
    Else
        Set wb = objExcel.Workbooks.Add
        wb.Worksheets(3).Delete
        wb.Worksheets(2).Delete
        wb.Worksheets(1).Name = strTabName
        Set ws = wb.Worksheets(1)
    End If
        
    Dim rs As DAO.Recordset
    Dim nFieldCount As Long
    Dim nRecordCount As Long
    Dim RetVal As Variant
    Dim nCurRec As Long
    Dim nCurSec As Long
    Dim nTotalSeconds As Long
    Dim nSecondsLeft As Long
    Dim strTest As String
    
    Set rs = CurrentDb.OpenRecordset("select * from " & strTableName)
    nFieldCount = rs.Fields.Count
    
    If Not rs.EOF Then
        rs.MoveLast
        nRecordCount = rs.RecordCount
        rs.MoveFirst
    End If
    
    RetVal = SysCmd(acSysCmdInitMeter, "Exporting " & strTableName & " to " & strFileName & ". . .", nRecordCount)

    
    For nCurrent = 0 To nFieldCount - 1
        strTest = rs.Fields(nCurrent).Name
        Do While InStr(strTest, "/") > 0
            strTest = Replace(strTest, "/", "")
        Loop
        ws.Range(FindExcelCell(nCurrent + 1, 1)) = strTest
        ws.Range(FindExcelCell(nCurrent + 1, 1)).Font.Bold = True
        ws.Range(FindExcelCell(nCurrent + 1, 1)).Interior.Color = RGB(222, 222, 222)
    Next
    
    nCurSec = Second(Now())
    Do While nCurSec = Second(Now())
    Loop
    nCurSec = Second(Now())
    Do While Not rs.EOF
        nCurRec = nCurRec + 1
        If Second(Now()) <> nCurSec And nCurRec < nRecordCount Then
            nCurSec = Second(Now())
            nTotalSeconds = nTotalSeconds + 1
            If nTotalSeconds > 3 Then
                nSecondsLeft = Int(((nTotalSeconds / nCurRec) * nRecordCount) * ((nRecordCount - nCurRec) / nRecordCount))
                RetVal = SysCmd(acSysCmdRemoveMeter)
                RetVal = SysCmd(acSysCmdInitMeter, "Exporting " & strTableName & " to tab " & strTabName & " in " & strFileName & ". . .  " & nSecondsLeft & " seconds remaining.", rs.RecordCount())
                RetVal = SysCmd(acSysCmdUpdateMeter, nCurRec)
                RetVal = DoEvents()
            End If
        End If
        strTest = ""
        'Check for blank lines--no need to export those!
        For nCurrent = 0 To nFieldCount - 1
            strTest = strTest & IIf(IsNull(rs.Fields), "", rs.Fields(nCurrent).Value)
        Next
        If Len(Trim(strTest)) > 0 Then
            For nCurrent = 0 To nFieldCount - 1
                If Not IsNull(rs.Fields(nCurrent).Value) Then
                    If rs.Fields(nCurrent).Value <> "" Then
                        If IsNumeric(rs.Fields(nCurrent).Value & "") Then
                            ws.Range(FindExcelCell(nCurrent + 1, nCurRec + 1)) = "'" & Trim(rs.Fields(nCurrent).Value)
                        Else
                            ws.Range(FindExcelCell(nCurrent + 1, nCurRec + 1)) = Trim(rs.Fields(nCurrent).Value)
                        End If
                    End If
                End If
            Next
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    ws.Range("A1").Select  'Move the cursor to the very first field
    If Len(Dir(strFileName, vbNormal)) > 0 Then
        'File already exists, just close and save.
        wb.Close (True)
    Else
        'File must be created.  Save and then close without saving.
        objExcel.Workbooks(1).SaveAs (strFileName)
        objExcel.Workbooks(1).Close (False)
    End If
    objExcel.DisplayAlerts = True
    objExcel.Quit
    Set objExcel = Nothing
    
    ExportToExcel = True
    RetVal = SysCmd(acSysCmdRemoveMeter)
End Function

Function IsNumeric(strCheck As String) As Boolean
    Dim nCurrent As Long
    IsNumeric = True
    nCurrent = 0
    strCheck = Trim(strCheck)
    Do While nCurrent < Len(strCheck) And IsNumeric = True
        nCurrent = nCurrent + 1
        If InStr("01234567890", Mid(strCheck, nCurrent, 1)) < 1 Then
            IsNumeric = False 'Part of the string is not a digit!
        End If
    Loop
End Function

Function FindExcelCell(nX As Long, nY As Long) As String
    Dim nPower1 As Long, nPower2 As Long
    nPower2 = 0
    If nX > 26 Then
        nPower2 = Int((nX - 1) / 26)
    End If
    nPower1 = nX - (26 * nPower2)
    If nPower2 > 0 Then
        FindExcelCell = Chr(64 + nPower2) & Chr(64 + nPower1) & nY
    Else
        FindExcelCell = Chr(64 + nPower1) & nY
    End If
End Function