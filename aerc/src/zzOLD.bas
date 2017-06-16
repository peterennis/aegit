Option Compare Database
Option Explicit

Public Sub ListIndexes()

    On Error GoTo 0

    Const adSchemaIndexes As Long = 12
    Dim cnn As Object ' ADODB.Connection
    Dim rst As Object ' ADODB.Recordset
    'Dim i As Long

    Set cnn = CurrentProject.Connection
    Set rst = cnn.OpenSchema(adSchemaIndexes)
    With rst
        'For i = 0 To (.Fields.Count - 1)
        '    Debug.Print .Fields(i).Name
        'Next i
        Do While Not .EOF
            If Left$(!TABLE_NAME, 4) <> "MSys" Then
                Debug.Print !TABLE_NAME, !INDEX_NAME, !PRIMARY_KEY
                .MoveNext
            Else
                .MoveNext
            End If
        Loop
        .Close
    End With
    Set rst = Nothing
    Set cnn = Nothing
End Sub

Public Function IsPK(ByVal tdf As DAO.TableDef, ByVal strField As String) As Boolean
    'Debug.Print "isPK"
    On Error GoTo 0

    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tdf.Indexes
        If idx.Primary Then
            For Each fld In idx.Fields
                If strField = fld.Name Then
                    IsPK = True
                    Exit Function
                End If
            Next fld
        End If
    Next idx
End Function

Public Function IsIndex(ByVal tdf As DAO.TableDef, ByVal strField As String) As Boolean
    'Debug.Print "isIndex"
    On Error GoTo 0

    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tdf.Indexes
        For Each fld In idx.Fields
            If strField = fld.Name Then
                IsIndex = True
                Exit Function
            End If
        Next fld
    Next idx
End Function

Public Function IsFK(ByVal tdf As DAO.TableDef, ByVal strField As String) As Boolean
    'Debug.Print "isFK"
    On Error GoTo 0
    
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tdf.Indexes
        If idx.Foreign Then
            For Each fld In idx.Fields
                If strField = fld.Name Then
                    IsFK = True
                    Exit Function
                End If
            Next fld
        End If
    Next idx
End Function

'VBA-Inspector:Ignore

'    Dim objChart As Object
'    Set objChart = Me.chtChart.Object
'    Dim objAxis As Object
'    Set objAxis = objChart.Axes(1)
'    objChart.PlotOnX = 0
'    ' Ref: http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.axis.scaletype.aspx
'    objAxis.ScaleType = xlScaleLinear
'    'objAxis.ScaleType = xlScaleLogarithmic

'

Public Function SpFolder(ByVal SpName As String) As String

    On Error GoTo 0

    Dim objShell As Object
    Set objShell = CreateObject("Shell.Application")
    Dim objFolder As Object
    Set objFolder = objShell.Namespace(SpName)
    Dim objFolderItem As Object
    Set objFolderItem = objFolder.Self

    SpFolder = objFolderItem.path

End Function
   
Public Sub ExportAllModulesToFile()
    ' Ref: http://wiki.lessthandot.com/index.php/Code_and_Code_Windows
    ' Ref: http://stackoverflow.com/questions/2794480/exporting-code-from-microsoft-access
    ' The reference for the FileSystemObject Object is Windows Script Host Object Model
    ' but it not necessary to add the reference for this procedure.
    On Error GoTo 0

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fil As Object
    Dim strMod As String
    Dim mdl As Object
    Dim i As Integer
    Dim strTxtFile As String
    Const Desktop As Long = &H10&

    ' Set up the file
    Debug.Print "CurrentProject.Name = " & CurrentProject.Name
    strTxtFile = SpFolder(Desktop) & "\" & Replace(CurrentProject.Name, ".", "_") & ".txt"
    Debug.Print "strTxtFile = " & strTxtFile
    Set fil = fso.CreateTextFile(SpFolder(Desktop) & "\" _
        & Replace(CurrentProject.Name, ".", " ") & ".txt")

    ' For each component in the project ...
    For Each mdl In VBE.ActiveVBProject.VBComponents
        ' using the count of lines ...
        If Left$(mdl.Name, 3) <> "zzz" Then
            Debug.Print mdl.Name
            i = VBE.ActiveVBProject.VBComponents(mdl.Name).CodeModule.CountOfLines
            ' put the code in a string ...
            If i > 0 Then
                strMod = VBE.ActiveVBProject.VBComponents(mdl.Name).CodeModule.Lines(1, i)
            End If
            ' and then write it to a file, first marking the start with
            ' some equal signs and the component name.
            fil.WriteLine String$(15, "=") & vbCrLf & mdl.Name _
                & vbCrLf & String$(15, "=") & vbCrLf & strMod
        End If
    Next
       
    ' Close eveything
    fil.Close
    Set fso = Nothing

End Sub

Public Sub SetRefToLibrary()
    ' http://www.exceltoolset.com/setting-a-reference-to-the-vba-extensibility-library-by-code/
    ' Adjusted for Microsoft Access
    ' Create a reference to the VBA Extensibility library
    On Error Resume Next        ' in case the reference already exits
    Access.Application.VBE.ActiveVBProject.References _
        .AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 0
End Sub

Public Function CodeLinesInProjectCount() As Long
    ' Ref: http://www.cpearson.com/excel/vbe.aspx
    ' Adjusted for Microsoft Access and Late Binding. No reference needed.
    ' Access.Application is used. Returns -1 if the VBProject is locked.
    On Error GoTo 0

    Dim VBP As Object               'VBIDE.VBProject
    Dim VBComp As Object            'VBIDE.VBComponent
    Dim LineCount As Long

    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=245480
    Const vbext_pp_locked As Integer = 1

    Set VBP = Access.Application.VBE.ActiveVBProject

    If VBP.Protection = vbext_pp_locked Then
        CodeLinesInProjectCount = -1
        Exit Function
    End If

    For Each VBComp In VBP.VBComponents
        If Left$(VBComp.Name, 3) <> "zzz" Then
            Debug.Print VBComp.Name, VBComp.CodeModule.CountOfLines
        End If
        LineCount = LineCount + VBComp.CodeModule.CountOfLines
    Next VBComp

    CodeLinesInProjectCount = LineCount

    Set VBP = Nothing

End Function

Public Sub GetAK()
    ' Ref: http://compgroups.net/comp.databases.ms-access/can-t-export-a-pass-through-query/357262

    On Error Resume Next
    CurrentDb.Execute "drop table t1"
    On Error GoTo 0
    CurrentDb.Execute "select *.* into t1 from pq"
    DoCmd.TransferText acExportDelim, , "t1", "c:\test.txt", True

End Sub

Public Sub IsAppOpen(ByVal strAppName As String)
    ' Ref: http://www.ehow.com/how_12111794_determine-excel-already-running-vba.html
    ' Ref: http://msdn.microsoft.com/en-us/library/office/aa164798(v=office.10).aspx

    Const ERR_APP_NOTRUNNING As Long = 429

    On Error GoTo Err_IsAppOpen:

    Dim objApp As Object

    Select Case strAppName
        Case "Access"
            Set objApp = GetObject(, "Access.Application")
            'Debug.Print objApp.Name
            If (objApp.Name = "Microsoft Access") Then
                Debug.Print "Access is running!"
            End If
        Case "Excel"
            'Debug.Print objApp.Name
            Set objApp = GetObject(, "Excel.Application")
            If (objApp = "Microsoft Excel") Then
                Debug.Print "Excel is running!"
            End If
        Case Else
            Debug.Print "Invalid App Name"
    End Select

    Set objApp = Nothing

Exit_IsAppOpen:
    Exit Sub

Err_IsAppOpen:
    If Err.Number = ERR_APP_NOTRUNNING Then
        Debug.Print strAppName & " is not running!"
    End If
    Set objApp = Nothing

End Sub

Public Sub TestPropertiesOutput()
    ' Ref: http://www.everythingaccess.com/tutorials.asp?ID=Accessing-detailed-file-information-provided-by-the-Operating-System
    ' Ref: http://www.techrepublic.com/article/a-simple-solution-for-tracking-changes-to-access-data/
    ' Ref: http://social.msdn.microsoft.com/Forums/office/en-US/480c17b3-e3d1-4f98-b1d6-fa16b23c6a0d/please-help-to-edit-the-table-query-form-and-modules-modified-date
    '
    ' Ref: http://perfectparadigm.com/tip001.html
    'SELECT MSysObjects.DateCreate, MSysObjects.DateUpdate,
    'MSysObjects.Name , MSysObjects.Type
    'FROM MSysObjects;
    On Error GoTo 0

    Debug.Print ">>>frm_Dummy"
    Debug.Print "DateCreated", DBEngine(0)(0).Containers("Forms")("frm_Dummy").Properties("DateCreated").Value
    Debug.Print "LastUpdated", DBEngine(0)(0).Containers("Forms")("frm_Dummy").Properties("LastUpdated").Value

    ' *** Ref: http://support.microsoft.com/default.aspx?scid=kb%3Ben-us%3B299554 ***
    'When the user initially creates a new Microsoft Access specific-object, such as a form), the database engine still
    'enters the current date and time into the DateCreate and DateUpdate columns in the MSysObjects table. However, when
    'the user modifies and saves the object, Microsoft Access does not notify the database engine; therefore, the
    'DateUpdate column always stays the same.
    
    ' Ref: http://questiontrack.com/how-can-i-display-a-last-modified-time-on-ms-access-form-995507.html

    Dim obj As AccessObject
    Dim dbs As Object

    Set dbs = Application.CurrentData
    Set obj = dbs.AllTables("tblThisTableHasSomeReallyLongNameButItCouldBeMuchLonger")
    Debug.Print ">>>" & obj.Name
    Debug.Print "DateCreated: " & obj.DateCreated
    Debug.Print "DateModified: " & obj.DateModified

End Sub

Public Sub ObjectCounts()

    On Error GoTo 0
 
    Dim qry As DAO.QueryDef
    Dim cnt As DAO.Container
 
    ' Delete all TEMP queries ...
    For Each qry In CurrentDb.QueryDefs
        If Left$(qry.Name, 1) = "~" Then
            CurrentDb.QueryDefs.Delete qry.Name
            CurrentDb.QueryDefs.Refresh
        End If
    Next qry
 
    ' Print the values to the immediate window
    With CurrentDb
 
        Debug.Print "--- From the DAO.Database ---"
        Debug.Print "-----------------------------"
        Debug.Print "Tables (Inc. System tbls): " & .TableDefs.Count
        Debug.Print "Querys: " & .QueryDefs.Count & vbCrLf
 
        For Each cnt In .Containers
            Debug.Print cnt.Name & ":" & cnt.Documents.Count
        Next cnt
 
    End With
 
    ' Use the "Project" collections to get the counts of objects
    With CurrentProject
        Debug.Print vbCrLf & "--- From the Access 'Project' ---"
        Debug.Print "---------------------------------"
        Debug.Print "Forms: " & .AllForms.Count
        Debug.Print "Reports: " & .AllReports.Count
        Debug.Print "DataAccessPages: " & .AllDataAccessPages.Count
        Debug.Print "Modules: " & .AllModules.Count
        Debug.Print "Macros (aka Scripts): " & .AllMacros.Count
    End With
 
End Sub

Public Sub PrettyXML(ByVal strPathFileName As String, Optional ByVal varDebug As Variant)

    On Error GoTo 0

    ' Beautify XML in VBA with MSXML6 only
    ' Ref: http://social.msdn.microsoft.com/Forums/en-US/409601d4-ca95-448a-aafc-aa0ee1ad67cd/beautify-xml-in-vba-with-msxml6-only?forum=xmlandnetfx
    Dim objXMLStyleSheet As Object
    Dim strXMLStyleSheet As String
    Dim objXMLDOMDoc As Object

    strXMLStyleSheet = "<xsl:stylesheet" & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "  xmlns:xsl=""http://www.w3.org/1999/XSL/Transform""" & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "  version=""1.0"">" & vbCrLf & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "<xsl:output method=""xml"" indent=""yes""/>" & vbCrLf & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "<xsl:template match=""@* | node()"">" & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "  <xsl:copy>" & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "    <xsl:apply-templates select=""@* | node()""/>" & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "  </xsl:copy>" & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "</xsl:template>" & vbCrLf & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "</xsl:stylesheet>"

    Set objXMLStyleSheet = CreateObject("Msxml2.DOMDocument.6.0")

    With objXMLStyleSheet
        ' Turn off Async I/O
        .async = False
        .validateOnParse = False
        .resolveExternals = False
    End With

    objXMLStyleSheet.LoadXML (strXMLStyleSheet)
    If objXMLStyleSheet.parseError.errorCode <> 0 Then
        Debug.Print "Some Error..."
        Exit Sub
    End If

    Set objXMLDOMDoc = CreateObject("Msxml2.DOMDocument.6.0")
    With objXMLDOMDoc
        ' Turn off Async I/O
        .async = False
        .validateOnParse = False
        .resolveExternals = False
    End With

    ' Ref: http://msdn.microsoft.com/en-us/library/ms762722(v=vs.85).aspx
    ' Ref: http://msdn.microsoft.com/en-us/library/ms754585(v=vs.85).aspx
    ' Ref: http://msdn.microsoft.com/en-us/library/aa468547.aspx
    objXMLDOMDoc.Load (strPathFileName)

    Dim strXMLResDoc As Object
    Set strXMLResDoc = CreateObject("Msxml2.DOMDocument.6.0")

    objXMLDOMDoc.transformNodeToObject objXMLStyleSheet, strXMLResDoc
    strXMLResDoc = strXMLResDoc.XML
    strXMLResDoc = Replace(strXMLResDoc, vbTab, Chr$(32) & Chr$(32), , , vbBinaryCompare)
    If Not IsMissing(varDebug) Then Debug.Print "Pretty XML Sample Output"
    Debug.Print strXMLResDoc

    Set objXMLDOMDoc = Nothing
    Set objXMLStyleSheet = Nothing

End Sub

Public Sub FormUseDefaultPrinter()
    ' Ref: http://msdn.microsoft.com/en-us/library/office/ff845464(v=office.15).aspx
    On Error GoTo 0

    Dim obj As Object
    For Each obj In CurrentProject.AllForms
        DoCmd.OpenForm FormName:=obj.Name, View:=acViewDesign
        If Not Forms(obj.Name).UseDefaultPrinter Then
            Forms(obj.Name).UseDefaultPrinter = True
            DoCmd.Save ObjectType:=acForm, ObjectName:=obj.Name
        End If
        DoCmd.Close
    Next obj

End Sub

Public Sub ReportUseDefaultPrinter()
    ' Ref: http://msdn.microsoft.com/en-us/library/office/ff845464(v=office.15).aspx
    On Error GoTo 0

    Dim obj As Object
    For Each obj In CurrentProject.AllReports
        DoCmd.OpenReport ReportName:=obj.Name, View:=acViewDesign
        If Not Reports(obj.Name).UseDefaultPrinter Then
            Reports(obj.Name).UseDefaultPrinter = True
            DoCmd.Save ObjectType:=acReport, ObjectName:=obj.Name
        End If
        DoCmd.Close
    Next obj

End Sub

Public Sub TestForCreateFormReportTextFile()
    ' Ref: http://social.msdn.microsoft.com/Forums/office/en-US/714d453c-d97a-4567-bd5f-64651e29c93a/how-to-read-text-a-file-into-a-string-1line-at-a-time-search-it-for-keyword-data?forum=accessdev
    ' Ref: http://bytes.com/topic/access/insights/953655-vba-standard-text-file-i-o-statements
    ' Ref: http://www.java2s.com/Code/VBA-Excel-Access-Word/File-Path/ExamplesoftheVBAOpenStatement.htm
    ' Ref: http://www.techonthenet.com/excel/formulas/instr.php
    '
    ' "Checksum =" , "NameMap = Begin",  "PrtMap = Begin",  "PrtDevMode = Begin"
    ' "PrtDevNames = Begin", "PrtDevModeW = Begin", "PrtDevNamesW = Begin"
    ' "OleData = Begin"
    On Error GoTo 0

    Dim fleIn As Integer
    Dim fleOut As Integer
    Dim strFileIn As String
    Dim strFileOut As String
    Dim strIn As String
    Dim i As Integer

    i = 0
    fleIn = FreeFile()
    strFileIn = "C:\TEMP\_chtQAQC.frm"
    Open strFileIn For Input As #fleIn

    fleOut = FreeFile()
    strFileOut = "C:\TEMP\_chtQAQC_frm.txt"
    Open strFileOut For Output As #fleOut

    Debug.Print "fleIn=" & fleIn, "fleOut=" & fleOut

    Do While Not EOF(fleIn)
        i = i + 1
        Line Input #fleIn, strIn
        If Left$(strIn, Len("Checksum =")) = "Checksum =" Then
            Print #fleOut, strIn
            Debug.Print i, strIn
        ElseIf InStr(1, strIn, "NameMap = Begin", vbTextCompare) > 0 Then
            Print #fleOut, strIn
            Debug.Print i, strIn
        ElseIf InStr(1, strIn, "PrtMip = Begin", vbTextCompare) > 0 Then
            Print #fleOut, strIn
            Debug.Print i, strIn
        ElseIf InStr(1, strIn, "PrtDevMode = Begin", vbTextCompare) > 0 Then
            Print #fleOut, strIn
            Debug.Print i, strIn
        ElseIf InStr(1, strIn, "PrtDevNames = Begin", vbTextCompare) > 0 Then
            Print #fleOut, strIn
            Debug.Print i, strIn
        ElseIf InStr(1, strIn, "PrtDevModeW = Begin", vbTextCompare) > 0 Then
            Print #fleOut, strIn
            Debug.Print i, strIn
        ElseIf InStr(1, strIn, "PrtDevNamesW = Begin", vbTextCompare) > 0 Then
            Print #fleOut, strIn
            Debug.Print i, strIn
        ElseIf InStr(1, strIn, "OleData = Begin", vbTextCompare) > 0 Then
            Print #fleOut, strIn
            Debug.Print i, strIn
        End If
    Loop

    Close fleIn
    Close fleOut

End Sub

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
    On Error GoTo 0

    Dim fleIn As Integer
    Dim fleOut As Integer
    Dim strFileIn As String
    Dim strFileOut As String
    Dim strIn As String
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
        If Left$(strIn, Len("Checksum =")) = "Checksum =" Then
            Exit Do
        Else
            Debug.Print i, strIn
            Print #fleOut, strIn
        End If
    Loop
    Do While Not EOF(fleIn)
        i = i + 1
        Line Input #fleIn, strIn
NextIteration:
        If FoundKeywordInLine(strIn) Then
            Debug.Print i & ">", strIn
            Print #fleOut, strIn
            Do While Not EOF(fleIn)
                i = i + 1
                Line Input #fleIn, strIn
                If Not FoundKeywordInLine(strIn, "End") Then
                    'Debug.Print "Not Found!!!", i
                    'GoTo SearchForEnd
                Else
                    Debug.Print i & ">", "Found End!!!"
                    Print #fleOut, strIn
                    i = i + 1
                    Line Input #fleIn, strIn
                    Debug.Print i & ":", strIn
                    'Stop
                    GoTo NextIteration
                End If
                'Stop
SearchForEnd:
            Loop
        Else
            'Stop
            Print #fleOut, strIn
            Debug.Print i, strIn
        End If
    Loop

    Close fleIn
    Close fleOut

End Sub

Public Function FoundKeywordInLine(ByVal strLine As String, Optional ByVal varEnd As Variant) As Boolean

    On Error GoTo 0

    FoundKeywordInLine = False
    If Not IsMissing(varEnd) Then
        If InStr(1, strLine, "End", vbTextCompare) > 0 Then
            FoundKeywordInLine = True
            Exit Function
        End If
    End If
    If InStr(1, strLine, "NameMap = Begin", vbTextCompare) > 0 Then
        FoundKeywordInLine = True
        Exit Function
    End If
    If InStr(1, strLine, "PrtMip = Begin", vbTextCompare) > 0 Then
        FoundKeywordInLine = True
        Exit Function
    End If
    If InStr(1, strLine, "PrtDevMode = Begin", vbTextCompare) > 0 Then
        FoundKeywordInLine = True
        Exit Function
    End If
    If InStr(1, strLine, "PrtDevNames = Begin", vbTextCompare) > 0 Then
        FoundKeywordInLine = True
        Exit Function
    End If
    If InStr(1, strLine, "PrtDevModeW = Begin", vbTextCompare) > 0 Then
        FoundKeywordInLine = True
        Exit Function
    End If
    If InStr(1, strLine, "PrtDevNamesW = Begin", vbTextCompare) > 0 Then
        FoundKeywordInLine = True
        Exit Function
    End If
    If InStr(1, strLine, "OleData = Begin", vbTextCompare) > 0 Then
        FoundKeywordInLine = True
        Exit Function
    End If

End Function

Public Sub SaveTableMacros()

    On Error GoTo 0
    ' Export Table Data to XML
    ' Ref: http://technet.microsoft.com/en-us/library/ee692914.aspx
    'Application.ExportXML acExportTable, "aeItems", "C:\Temp\aeItemsData.xml"

    ' Save table macros as XML
    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=99179
    Application.SaveAsText acTableDataMacro, "aeItems", "C:\Temp\aeItems.xml"
    Debug.Print , "Items table macros saved to C:\Temp\aeItems.xml"

    PrettyXML "C:\Temp\aeItems.xml"

End Sub

Public Function IncrementReset() As Long
    ' This function returns an incremented number each time it's called.  Resets after 2 seconds.
    On Error GoTo 0
    Static nIncrement As Long
    'Now we put in a reset based on time!
    Static nLastSecond As Long
    Dim nNowSecond As Long
    nNowSecond = 3600 * Hour(Now) + 60 * Minute(Now) + Second(Now)
    If Math.Abs(nNowSecond - nLastSecond) < 2 Then
        nIncrement = nIncrement + 1
    Else
        nIncrement = 1
    End If
    nLastSecond = nNowSecond
    IncrementReset = nIncrement
End Function