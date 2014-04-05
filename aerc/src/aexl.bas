Option Compare Database
Option Explicit

Public Sub TestExportToExcel()
    On Error GoTo 0
    Dim bln As Boolean
    bln = ExportToExcel("aeItems", "C:\TEMP\Exported_aeItems.xls")
End Sub

Public Function ExportToExcel(ByVal strTableName, ByVal strFileName, Optional ByVal strTabName As String = "Sheet1") As Boolean
' Original example Ref: http://www.saplsmw.com, James Kauffman, Ver 1.20 Updated 17 Jun 2010
' Ref: http://www.granite.ab.ca/access/latebinding.htm

    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    Dim wkb As Object
    Dim wks As Object
    Dim wks2 As Object
    Dim nCurrent As Long
    Dim isFound As Boolean

    On Error Resume Next
    isFound = False
    If Len(Dir$(strFileName, vbNormal)) > 0 Then 'File exists!
        Set wkb = objExcel.Workbooks.Open(strFileName, False, False)
        nCurrent = wkb.Worksheets.Count
        Do While nCurrent > 0
            If wkb.Worksheets(nCurrent).Name = strTabName Then
                objExcel.DisplayAlerts = False
                Set wks2 = wkb.Worksheets(nCurrent)
                Set wks = wkb.Worksheets.Add
                wks2.Delete
                wks.Name = strTabName
                isFound = True
            End If
            nCurrent = nCurrent - 1
        Loop
        If Not isFound Then
            Set wks = wkb.Worksheets.Add
            wks.Name = strTabName
        End If
    Else
        Set wkb = objExcel.Workbooks.Add
        wkb.Worksheets(3).Delete
        wkb.Worksheets(2).Delete
        wkb.Worksheets(1).Name = strTabName
        Set wks = wkb.Worksheets(1)
    End If
        
    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset("select * from " & strTableName)

    Dim nFieldCount As Long
    Dim nRecordCount As Long
    Dim RetVal As Variant
    Dim nCurRec As Long
    Dim nCurSec As Long
    Dim nTotalSeconds As Long
    Dim nSecondsLeft As Long
    Dim strTest As String
    
    nFieldCount = rst.Fields.Count
    
    If Not rst.EOF Then
        rst.MoveLast
        nRecordCount = rst.RecordCount
        rst.MoveFirst
    End If
    
    RetVal = SysCmd(acSysCmdInitMeter, "Exporting " & strTableName & " to " & strFileName & ". . .", nRecordCount)

    
    For nCurrent = 0 To nFieldCount - 1
        strTest = rst.Fields(nCurrent).Name
        Do While InStr(strTest, "/") > 0
            strTest = Replace(strTest, "/", vbNullString)
        Loop
        wks.Range(FindExcelCell(nCurrent + 1, 1)) = strTest
        wks.Range(FindExcelCell(nCurrent + 1, 1)).Font.Bold = True
        wks.Range(FindExcelCell(nCurrent + 1, 1)).Interior.Color = RGB(222, 222, 222)
    Next
    
    nCurSec = Second(Now())
    Do While nCurSec = Second(Now())
    Loop
    nCurSec = Second(Now())
    Do While Not rst.EOF
        nCurRec = nCurRec + 1
        If Second(Now()) <> nCurSec And nCurRec < nRecordCount Then
            nCurSec = Second(Now())
            nTotalSeconds = nTotalSeconds + 1
            If nTotalSeconds > 3 Then
                nSecondsLeft = Int(((nTotalSeconds / nCurRec) * nRecordCount) * ((nRecordCount - nCurRec) / nRecordCount))
                RetVal = SysCmd(acSysCmdRemoveMeter)
                RetVal = SysCmd(acSysCmdInitMeter, "Exporting " & strTableName & " to tab " & strTabName & " in " & strFileName & ". . .  " & nSecondsLeft & " seconds remaining.", rst.RecordCount())
                RetVal = SysCmd(acSysCmdUpdateMeter, nCurRec)
                RetVal = DoEvents()
            End If
        End If
        strTest = vbNullString
        'Check for blank lines--no need to export those!
        For nCurrent = 0 To nFieldCount - 1
            strTest = strTest & IIf(IsNull(rst.Fields), vbNullString, rst.Fields(nCurrent).Value)
        Next
        If Len(Trim$(strTest)) > 0 Then
            For nCurrent = 0 To nFieldCount - 1
                If Not IsNull(rst.Fields(nCurrent).Value) Then
                    If rst.Fields(nCurrent).Value <> vbNullString Then
                        If IsNumeric(rst.Fields(nCurrent).Value & vbNullString) Then
                            wks.Range(FindExcelCell(nCurrent + 1, nCurRec + 1)) = "'" & Trim$(rst.Fields(nCurrent).Value)
                        Else
                            wks.Range(FindExcelCell(nCurrent + 1, nCurRec + 1)) = Trim$(rst.Fields(nCurrent).Value)
                        End If
                    End If
                End If
            Next
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    wks.Range("A1").Select  'Move the cursor to the very first field
    If Len(Dir$(strFileName, vbNormal)) > 0 Then
        'File already exists, just close and save.
        wkb.Close (True)
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

Private Function IsNumeric(ByVal strCheck As String) As Boolean

    On Error GoTo 0
    Dim nCurrent As Long
    IsNumeric = True
    nCurrent = 0
    strCheck = Trim$(strCheck)
    Do While nCurrent < Len(strCheck) And IsNumeric = True
        nCurrent = nCurrent + 1
        If InStr("01234567890", Mid$(strCheck, nCurrent, 1)) < 1 Then
            IsNumeric = False 'Part of the string is not a digit!
        End If
    Loop

End Function

Private Function FindExcelCell(ByVal nX As Long, ByVal nY As Long) As String

    On Error GoTo 0
    Dim nPower1 As Long
    Dim nPower2 As Long
    nPower2 = 0
    If nX > 26 Then
        nPower2 = Int((nX - 1) / 26)
    End If
    nPower1 = nX - (26 * nPower2)
    If nPower2 > 0 Then
        FindExcelCell = Chr$(64 + nPower2) & Chr$(64 + nPower1) & nY
    Else
        FindExcelCell = Chr$(64 + nPower1) & nY
    End If

End Function