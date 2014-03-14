Option Compare Database
Option Explicit

Public Function FoundkeyWordInLine(strLine As String) As Boolean

    FoundkeyWordInLine = False
    If InStr(1, strLine, "NameMap = Begin", vbTextCompare) > 0 Then
        FoundkeyWordInLine = True
    End If
    If InStr(1, strLine, "PrtMip = Begin", vbTextCompare) > 0 Then
        FoundkeyWordInLine = True
    End If
    If InStr(1, strLine, "PrtDevMode = Begin", vbTextCompare) > 0 Then
        FoundkeyWordInLine = True
    End If
    If InStr(1, strLine, "PrtDevNames = Begin", vbTextCompare) > 0 Then
        FoundkeyWordInLine = True
    End If
    If InStr(1, strLine, "PrtDevModeW = Begin", vbTextCompare) > 0 Then
        FoundkeyWordInLine = True
    End If
    If InStr(1, strLine, "PrtDevNamesW = Begin", vbTextCompare) > 0 Then
        FoundkeyWordInLine = True
    End If
    If InStr(1, strLine, "OleData = Begin", vbTextCompare) > 0 Then
        FoundkeyWordInLine = True
    End If

End Function

Public Sub CreateFormReportTextFile()
' Ref: http://social.msdn.microsoft.com/Forums/office/en-US/714d453c-d97a-4567-bd5f-64651e29c93a/how-to-read-text-a-file-into-a-string-1line-at-a-time-search-it-for-keyword-data?forum=accessdev
' Ref: http://bytes.com/topic/access/insights/953655-vba-standard-text-file-i-o-statements
' Ref: http://www.java2s.com/Code/VBA-Excel-Access-Word/File-Path/ExamplesoftheVBAOpenStatement.htm
' Ref: http://www.techonthenet.com/excel/formulas/instr.php
' Ref: http://stackoverflow.com/questions/8680640/vba-how-to-conditionally-skip-a-for-loop-iteration
'
' "Checksum =" , "NameMap = Begin",  "PrtMap = Begin",  "PrtDevMode = Begin"
' "PrtDevNames = Begin", "PrtDevModeW = Begin", "PrtDevNamesW = Begin"
' "OleData = Begin"

    Dim fleIn As Integer
    Dim fleOut As Integer
    Dim strFileIn As String
    Dim strFileOut As String
    Dim strIn As String
    Dim strOut As String
    Dim i As Integer

    fleIn = FreeFile()
    strFileIn = "C:\TEMP\_chtQAQC.frm"
    Open strFileIn For Input As #fleIn

    fleOut = FreeFile()
    strFileOut = "C:\TEMP\_chtQAQC_frm.txt"
    Open strFileOut For Output As #fleOut

    Debug.Print "fleIn=" & fleIn, "fleOut=" & fleOut

    i = 0
    Do While Not EOF(fleIn)
        i = i + 1
        Line Input #fleIn, strIn
        If Left(strIn, Len("Checksum =")) = "Checksum =" Then
            Exit Do
        Else
            Print #fleOut, strIn
            Debug.Print i, strIn
        End If
    Loop
    Do While Not EOF(fleIn)
        i = i + 1
        Line Input #fleIn, strIn
        If FoundkeyWordInLine(strIn) Then
            Debug.Print strIn
            Stop
        Else
            Print #fleOut, strIn
            Debug.Print i, strIn
        End If
    Loop

    Close fleIn
    Close fleOut

End Sub

Public Sub SaveTableMacros()

    ' Export Table Data to XML
    ' Ref: http://technet.microsoft.com/en-us/library/ee692914.aspx
'    Application.ExportXML acExportTable, "aeItems", "C:\Temp\aeItemsData.xml"

    ' Save table macros as XML
    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=99179
    Application.SaveAsText acTableDataMacro, "aeItems", "C:\Temp\aeItems.xml"
    Debug.Print , "Items table macros saved to C:\Temp\aeItems.xml"

    PrettyXML "C:\Temp\aeItems.xml"

End Sub

Public Sub ApplicationInformation()
' Ref: http://msdn.microsoft.com/en-us/library/office/aa223101(v=office.11).aspx
' Ref: http://msdn.microsoft.com/en-us/library/office/aa173218(v=office.11).aspx
' Ref: http://msdn.microsoft.com/en-us/library/office/ff845735(v=office.15).aspx

    Dim intProjType As Integer
    Dim strProjType As String
    Dim lng As Long

    intProjType = Application.CurrentProject.ProjectType

    Select Case intProjType
        Case 0 ' acNull
            strProjType = "acNull"
        Case 1 ' acADP
            strProjType = "acADP"
        Case 2 ' acMDB
            strProjType = "acMDB"
        Case Else
            MsgBox "Can't determine ProjectType"
    End Select

    Debug.Print Application.CurrentProject.FullName
    Debug.Print "Project Type", intProjType, strProjType
    lng = CodeLinesInProjectCount

End Sub

Public Sub TestRegKey()

    Dim strKey As String

    ' Office 2010
    'strKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Common\General\ReportAddinCustomUIErrors"
    ' Office 2013
    strKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Common\General\ReportAddinCustomUIErrors"
    If RegKeyExists(strKey) Then
        Debug.Print strKey & " Exists!"
    Else
        Debug.Print strKey & " Does NOT Exist!"
    End If

End Sub

Public Function RegKeyExists(strRegKey As String) As Boolean
' Return True if the registry key i_RegKey was found and False if not
' Ref: http://vba-corner.livejournal.com/3054.html

    On Error GoTo ErrorHandler

    ' Use Windows scripting and try to read the registry key
    Dim myWS As Object
    Set myWS = CreateObject("WScript.Shell")

    myWS.RegRead strRegKey
    ' Key was found
    RegKeyExists = True
    Exit Function
  
ErrorHandler:
    ' Key was not found
    RegKeyExists = False

End Function