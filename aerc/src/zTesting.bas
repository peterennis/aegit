Option Compare Database
Option Explicit

' Remove this after integration with aegitClass
Public Const THE_SOURCE_FOLDER = "C:\ae\aegit\aerc\src\"

Public Sub TestCreateDbScript()
    'CreateDbScript "C:\Temp\Schema.txt"
    Debug.Print "THE_SOURCE_FOLDER=" & THE_SOURCE_FOLDER
    CreateDbScript THE_SOURCE_FOLDER & "Schema.txt"
End Sub

Public Sub CreateDbScript(strScriptFile As String)
' Remou - Ref: http://stackoverflow.com/questions/698839/how-to-extract-the-schema-of-an-access-mdb-database/9910716#9910716

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim ndx As DAO.Index
    Dim strSQL As String
    Dim strFlds As String
    Dim strCn As String
    Dim strLinkedTablePath As String
    Dim fs As Object
    Dim f As Object

    Set db = CurrentDb
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.CreateTextFile(strScriptFile)

    strSQL = "Public Sub CreateTheDb()" & vbCrLf
    f.WriteLine strSQL
    strSQL = "Dim strSQL As String"
    f.WriteLine strSQL
    strSQL = "On Error GoTo ErrorTrap"
    f.WriteLine strSQL

    For Each tdf In db.TableDefs
        If Not (Left(tdf.Name, 4) = "MSys" _
                Or Left(tdf.Name, 4) = "~TMP" _
                Or Left(tdf.Name, 3) = "zzz") Then

            strLinkedTablePath = GetLinkedTableCurrentPath(tdf.Name)
            If strLinkedTablePath <> "" Then
                f.WriteLine vbCrLf & "'OriginalLink=>" & strLinkedTablePath
            Else
                f.WriteLine vbCrLf & "'Local Table"
            End If

            strSQL = "strSQL=""CREATE TABLE [" & tdf.Name & "] ("
            strFlds = ""

            For Each fld In tdf.Fields

                strFlds = strFlds & ",[" & fld.Name & "] "

                Select Case fld.Type
                    Case dbText
                        'No look-up fields
                        strFlds = strFlds & "Text (" & fld.Size & ")"
                    Case dbLong
                        If (fld.Attributes And dbAutoIncrField) = 0& Then
                            strFlds = strFlds & "Long"
                        Else
                            strFlds = strFlds & "Counter"
                        End If
                    Case dbBoolean
                        strFlds = strFlds & "YesNo"
                    Case dbByte
                        strFlds = strFlds & "Byte"
                    Case dbInteger
                        strFlds = strFlds & "Integer"
                    Case dbCurrency
                        strFlds = strFlds & "Currency"
                    Case dbSingle
                        strFlds = strFlds & "Single"
                    Case dbDouble
                        strFlds = strFlds & "Double"
                    Case dbDate
                        strFlds = strFlds & "DateTime"
                    Case dbBinary
                        strFlds = strFlds & "Binary"
                    Case dbLongBinary
                        strFlds = strFlds & "OLE Object"
                    Case dbMemo
                        If (fld.Attributes And dbHyperlinkField) = 0& Then
                            strFlds = strFlds & "Memo"
                        Else
                            strFlds = strFlds & "Hyperlink"
                        End If
                    Case dbGUID
                        strFlds = strFlds & "GUID"
                End Select

            Next

            strSQL = strSQL & Mid(strFlds, 2) & " )""" & vbCrLf & "Currentdb.Execute strSQL"
            f.WriteLine vbCrLf & strSQL

            'Indexes
            For Each ndx In tdf.Indexes

                If ndx.Unique Then
                    strSQL = "strSQL=""CREATE UNIQUE INDEX "
                Else
                    strSQL = "strSQL=""CREATE INDEX "
                End If

                strSQL = strSQL & "[" & ndx.Name & "] ON [" & tdf.Name & "] ("
                strFlds = ""

                For Each fld In tdf.Fields
                    strFlds = ",[" & fld.Name & "]"
                Next

                strSQL = strSQL & Mid(strFlds, 2) & ") "
                strCn = ""

                If ndx.Primary Then
                    strCn = " PRIMARY"
                End If

                If ndx.Required Then
                    strCn = strCn & " DISALLOW NULL"
                End If

                If ndx.IgnoreNulls Then
                    strCn = strCn & " IGNORE NULL"
                End If

                If Trim(strCn) <> vbNullString Then
                    strSQL = strSQL & " WITH" & strCn & " "
                End If

                f.WriteLine vbCrLf & strSQL & """" & vbCrLf & "Currentdb.Execute strSQL"
            Next
        End If
    Next

    'strSQL = vbCrLf & "Debug.Print " & """" & "Done" & """"
    'f.WriteLine strSQL
    f.WriteLine
    f.WriteLine "'Access 2010 - Compact And Repair"
    strSQL = "SendKeys " & """" & "%F{END}{ENTER}%F{TAB}{TAB}{ENTER}" & """" & ", False"
    f.WriteLine strSQL
    strSQL = "Exit Sub"
    f.WriteLine strSQL
    strSQL = "ErrorTrap:"
    f.WriteLine strSQL
    'MsgBox "Erl=" & Erl & vbCrLf & "Err.Number=" & Err.Number & vbCrLf & "Err.Description=" & Err.Description
    strSQL = "MsgBox " & """" & "Erl=" & """" & " & vbCrLf & " & _
                """" & "Err.Number=" & """" & " & Err.Number & vbCrLf & " & _
                """" & "Err.Description=" & """" & " & Err.Description"
    f.WriteLine strSQL & vbCrLf
    strSQL = "End Sub"
    f.WriteLine strSQL

    f.Close
    Debug.Print "Done"

End Sub

Public Function GetLinkedTableCurrentPath(MyLinkedTable As String) As String
' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=198057
' To test in the Immediate window:       ? getcurrentpath("Const")
'====================================================================
' Procedure : GetLinkedTableCurrentPath
' DateTime  : 08/23/2010
' Author    : Rx
' Purpose   : Returns Current Path of a Linked Table in Access
' Updates   : Peter F. Ennis
' Updated   : All notes moved to change log
' History   : See comment details, basChangeLog, commit messages on github
'====================================================================
    On Error GoTo PROC_ERR
    GetLinkedTableCurrentPath = Mid(CurrentDb.TableDefs(MyLinkedTable).Connect, InStr(1, CurrentDb.TableDefs(MyLinkedTable).Connect, "=") + 1)
        ' non-linked table returns blank - the Instr removes the "Database="

PROC_EXIT:
    On Error Resume Next
    Exit Function

PROC_ERR:
    Select Case Err.Number
        'Case ###         ' Add your own error management or log error to logging table
        Case Else
            'a custom log usage function commented out
            'function LogUsage(ByVal strFormName As String, strCallingProc As String, Optional ControlName) As Boolean
            'call LogUsage Err.Number, "basRelinkTables", "GetCurrentPath" ()
    End Select
    Resume PROC_EXIT
End Function

' Ref: http://www.utteraccess.com/forum/lofiversion/index.php/t1995627.html
'-------------------------------------------------------------------------------------------------
' Procedure : ExecSQL
' DateTime  : 30/03/2009 10:19
' Author    : Dial222
' Purpose   : Execute SQL Select statements in the Immediate window
' Context   : Module basSQL2IMM
' Notes     : No error trapping whatsover - this is a 1.0 technology!
'             Max out at 194 data rows since immediate only displays 100!
'
' Usage     : in the immediate pane: ?execsql("select * from zstblprofile","|")
'
' Revision History
' Version   Date        Who             What
' 01        30/03/2009  Dial222         Function 'ExecSQL' Created
' 02        30/03/2009  Dial222         Added code for left/right align of text/numeric data
'                                       Added MaxRowLen and vbCrLF parsing functionality
'                                       Uprated cMaxRows to 194
'-------------------------------------------------------------------------------------------------
'

Public Function ExecSQL(strSQL As String, Optional strColumDelim As String = "|") As Boolean

    Dim rs              As DAO.Recordset
    Dim aintLen()       As Integer
    Dim i               As Integer
    Dim str             As String
    Dim lngRowCOunt     As Long

    Const cMaxRows      As Integer = 194
    Const cMaxRowLen    As Integer = 1023  ' Max width of immediate pane in characters, truncate after this.

    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)

    With rs
        .MoveLast
        .MoveFirst

        lngRowCOunt = .RecordCount
        If lngRowCOunt > 0 Then
            If lngRowCOunt > cMaxRows Then
                Debug.Print "Too many rows to return, will only print first " & cMaxRows & " rows."
            End If

            ReDim Preserve aintLen(.Fields.Count)

            For i = 0 To .Fields.Count - 1
                ' Initialise field len to field name len
                aintLen(i) = Len(.Fields(i).Name) + 3
            Next i

            ' On this pass just get length of field data for formatting
            Do Until .EOF
                If .AbsolutePosition = cMaxRows Then
                    ' Stop at the magic number
                    Exit Do
                Else
                    For i = 0 To rs.Fields.Count - 1
                        ' Test and update field len
                        If Len(CStr(Nz(.Fields(i).Value, ""))) > aintLen(i) Then
                            aintLen(i) = Len(CStr(.Fields(i).Value)) + 3
                        End If
                    Next i
                End If
                .MoveNext
            Loop

            ' Print Column Headers
            str = "Row " & strColumDelim & " "
            For i = 0 To rs.Fields.Count - 1
                ' Initialise field len to field name len
                str = str & Left(.Fields(i).Name & Space(aintLen(i)), aintLen(i)) & " " & strColumDelim & " "
            Next i

            ' Print the header row
            Debug.Print Left(str, cMaxRowLen)
            str = Space(Len(str))
            str = Replace(str, " ", "-")

            ' print underscores
            Debug.Print Left(str, cMaxRowLen)
            str = ""

            ' Start over for the data
            .MoveFirst

            Do Until .EOF
                If .AbsolutePosition = cMaxRows Then
                    Exit Do
                Else
                    str = Left(.AbsolutePosition + 1 & Space(3), 3) & " " & strColumDelim & " "
                    For i = 0 To .Fields.Count - 1
                        Select Case .Fields(i).Type
                            Case Is = 3, 4, 5, 6, 7, 8, 16, 19, 20, 21, 22, 23 ' The numeric DataTypeEnums
                                str = str & Right(Space(aintLen(i)) & .Fields(i).Value, aintLen(i)) & " " & strColumDelim & " "
                            Case Else
                                ' Is it number stored as text
                                If IsNumeric(.Fields(i).Value) Then
                                    ' Right align
                                    str = str & Right(Space(aintLen(i)) & .Fields(i).Value, aintLen(i)) & " " & strColumDelim & " "
                                Else
                                    ' Left align
                                    str = str & Left(.Fields(i).Value & Space(aintLen(i)), aintLen(i)) & " " & strColumDelim & " "
                                End If
                        End Select
                    Next i
                End If

                ' Parse out vbCrLf and dump data row to immediate
                Debug.Print Left(Replace(Replace(str, Chr(13), " "), Chr(10), " "), cMaxRowLen)
                .MoveNext
                str = ""
            Loop

            ExecSQL = True
        Else
            Debug.Print "No rows returned"
        End If
    End With

    Set rs = Nothing

End Function

Public Function SpFolder(SpName)

    Dim objShell As Object
    Dim objFolder As Object
    Dim objFolderItem As Object

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(SpName)

    Set objFolderItem = objFolder.Self

    SpFolder = objFolderItem.Path

End Function
   
Public Sub AllCodeToDesktop()
' Ref: http://wiki.lessthandot.com/index.php/Code_and_Code_Windows
' Ref: http://stackoverflow.com/questions/2794480/exporting-code-from-microsoft-access
' The reference for the FileSystemObject Object is Windows Script Host Object Model
' but it not necessary to add the reference for this procedure.

    Const Desktop = &H10&
    Const MyDocuments = &H5&

    Dim fs As Object
    Dim f As Object
    Dim strMod As String
    Dim mdl As Object
    Dim i As Integer
    Dim strTxtFile As String

    Set fs = CreateObject("Scripting.FileSystemObject")

    'Set up the file
    Debug.Print "CurrentProject.Name = " & CurrentProject.Name
    strTxtFile = SpFolder(Desktop) & "\" & Replace(CurrentProject.Name, ".", " ") & ".txt"
    Debug.Print "strTxtFile = " & strTxtFile
    Set f = fs.CreateTextFile(SpFolder(Desktop) & "\" _
        & Replace(CurrentProject.Name, ".", " ") & ".txt")

    'For each component in the project ...
    For Each mdl In VBE.ActiveVBProject.VBComponents
        'using the count of lines ...
        i = VBE.ActiveVBProject.VBComponents(mdl.Name).codemodule.CountOfLines
        'put the code in a string ...
        If i > 0 Then
            strMod = VBE.ActiveVBProject.VBComponents(mdl.Name).codemodule.Lines(1, i)
        End If
        'and then write it to a file, first marking the start with
        'some equal signs and the component name.
        f.WriteLine String(15, "=") & vbCrLf & mdl.Name _
            & vbCrLf & String(15, "=") & vbCrLf & strMod
    Next
       
    'Close eveything
    f.Close
    Set fs = Nothing

End Sub

Public Function PropertyExists(obj As Object, strPropertyName As String) As Boolean
' Ref: http://www.utteraccess.com/forum/Description-property-Mic-t552348.html
' e.g. ? PropertyExists(CurrentDB. ("The Name Of Your Table"), "Description")
    Dim var As Variant

    On Error Resume Next
    Set var = obj.Properties(strPropertyName)
    If Err.Number > 0 Then
        PropertyExists = False
    Else
        PropertyExists = True
    End If

End Function

Public Sub GetPropertyDescription()
' Ref: http://www.dbforums.com/microsoft-access/1620765-read-ms-access-table-properties-using-vba.html

    Dim dbs As DAO.Database
    Dim obj As Object
    Dim prp As Property

    Set dbs = Application.CurrentDb
    Set obj = dbs.Containers("modules").Documents("aegitClass")

    On Error Resume Next
    For Each prp In obj.Properties
        Debug.Print prp.Name, prp.Value
    Next prp

    Set obj = Nothing
    Set dbs = Nothing

End Sub

Public Sub TestListAllProperties()
    'ListAllProperties ("modules")
    ListAllProperties ("tables")
End Sub

Public Sub ListGUID()
' Ref: http://stackoverflow.com/questions/8237914/how-to-get-the-guid-of-a-table-in-microsoft-access

    Dim i As Integer
    Dim arrGUID8() As Byte
    Dim strGUID As String

    arrGUID8 = CurrentDb.TableDefs("tblThisTableHasSomeReallyLongNameButItCouldBeMuchLonger").Properties("GUID").Value
    For i = 1 To 8
        strGUID = strGUID & Hex(arrGUID8(i)) & "-"
    Next
    Debug.Print Left(strGUID, 23)

End Sub

Public Function fListGUID(strTableName As String) As String
' Ref: http://stackoverflow.com/questions/8237914/how-to-get-the-guid-of-a-table-in-microsoft-access
' e.g. ?fListGUID("tblThisTableHasSomeReallyLongNameButItCouldBeMuchLonger")

    Dim i As Integer
    Dim arrGUID8() As Byte
    Dim strGUID As String

    arrGUID8 = CurrentDb.TableDefs(strTableName).Properties("GUID").Value
    For i = 1 To 8
        strGUID = strGUID & Hex(arrGUID8(i)) & "-"
    Next
    'Debug.Print Left(strGUID, 23)
    fListGUID = Left(strGUID, 23)

End Function

Public Sub ListAllProperties(strContainer As String)
' Ref: http://www.dbforums.com/microsoft-access/1620765-read-ms-access-table-properties-using-vba.html
' Ref: http://ms-access.veryhelper.com/q_ms-access-database_153855.html
' Ref: http://msdn.microsoft.com/en-us/library/office/aa139941(v=office.10).aspx
    
    Dim dbs As DAO.Database
    Dim obj As Object
    Dim prp As Property
    Dim doc As Document

    Set dbs = Application.CurrentDb
    Set obj = dbs.Containers(strContainer)

    'Debug.Print "Modules", obj.Documents.Count
    'Debug.Print "Modules", obj.Documents(1).Name
    'Debug.Print "Modules", obj.Documents(2).Name

    ' Ref: http://stackoverflow.com/questions/16642362/how-to-get-the-following-code-to-continue-on-error
    For Each doc In obj.Documents
        If Left(doc.Name, 4) <> "MSys" And Left(doc.Name, 3) <> "zzz" Then
            Debug.Print ">>>" & doc.Name
            For Each prp In doc.Properties
                On Error Resume Next
                    If prp.Name = "GUID" And strContainer = "tables" Then
                        Debug.Print prp.Name, fListGUID(doc.Name)
                    ElseIf prp.Name = "DOL" Then
                        Debug.Print prp.Name, "Track name AutoCorrect info is ON!"
                    ElseIf prp.Name = "NameMap" Then
                        Debug.Print prp.Name, "Track name AutoCorrect info is ON!"
                    Else
                        Debug.Print prp.Name, prp.Value
                    End If
                On Error GoTo 0
            Next
        End If
    Next

    Set obj = Nothing
    Set dbs = Nothing

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

Public Sub ListAccessApplicationOptions()
' Ref: http://msdn.microsoft.com/en-us/library/office/aa140020(v=office.10).aspx (2000)
' Ref: http://msdn.microsoft.com/en-us/library/office/aa189769(v=office.10).aspx (XP)
'   IME is Microsoft Global Input Method Editors (IMEs)
'   Ref: http://www.dbforums.com/microsoft-access/993286-what-ime.html
' Ref: http://msdn.microsoft.com/en-us/library/office/aa172326(v=office.11).aspx (2003)
' Ref: http://msdn.microsoft.com/en-us/library/office/bb256546(v=office.12).aspx (2007)
' Ref: http://msdn.microsoft.com/en-us/library/office/ff823177(v=office.14).aspx (2010)
' *** Ref: http://msdn.microsoft.com/en-us/library/office/ff823177.aspx (2013)
' Ref: http://office.microsoft.com/en-us/access-help/HV080750165.aspx (2013?)
' Set Options from Visual Basic

    Dim dbs As Database
    Set dbs = CurrentDb

    On Error Resume Next
    Debug.Print ">>>Standard Options"
    '2000 The following options are equivalent to the standard startup options found in the Startup Options dialog box.
    Debug.Print , "2000", "AppTitle", dbs.Properties!AppTitle                               'String  The title of an application, as displayed in the title bar.
    Debug.Print , "2000", "AppIcon", dbs.Properties!AppIcon                                 'String  The file name and path of an application's icon.
    Debug.Print , "2000", "StartupMenuBar", dbs.Properties!StartUpMenuBar                   'String  Sets the default menu bar for the application.
    Debug.Print , "2000", "AllowFullMenus", dbs.Properties!AllowFullMenus                   'True/False  Determines if the built-in Access menu bars are displayed.
    Debug.Print , "2000", "AllowShortcutMenus", dbs.Properties!AllowShortcutMenus           'True/False  Determines if the built-in Access shortcut menus are displayed.
    Debug.Print , "2000", "StartupForm", dbs.Properties!StartUpForm                         'String  Sets the form or data page to show when the application is first opened.
    Debug.Print , "2000", "StartupShowDBWindow", dbs.Properties!StartUpShowDBWindow         'True/False  Determines if the database window is displayed when the application is first opened.
    Debug.Print , "2000", "StartupShowStatusBar", dbs.Properties!StartUpShowStatusBar       'True/False  Determines if the status bar is displayed.
    Debug.Print , "2000", "StartupShortcutMenuBar", dbs.Properties!StartUpShortcutMenuBar   'String  Sets the shortcut menu bar to be used in all forms and reports.
    Debug.Print , "2000", "AllowBuiltInToolbars", dbs.Properties!AllowBuiltInToolbars       'True/False  Determines if the built-in Access toolbars are displayed.
    Debug.Print , "2000", "AllowToolbarChanges", dbs.Properties!AllowToolbarChanges         'True/False  Determined if toolbar changes can be made.
    Debug.Print ">>>Advanced Option"
    Debug.Print , "2000", "AllowSpecialKeys", dbs.Properties!AllowSpecialKeys               'option (True/False value) determines if the use of special keys is permitted. It is equivalent to the advanced startup option found in the Startup Options dialog box.
    Debug.Print ">>>Extra Options"
    'The following options are not available from the Startup Options dialog box or any other Access user interface component, they are only available in programming code.
    Debug.Print , "2000", "AllowBypassKey", dbs.Properties!AllowBypassKey                   'True/False  Determines if the SHIFT key can be used to bypass the application load process.
    Debug.Print , "2000", "AllowBreakIntoCode", dbs.Properties!AllowBreakIntoCode           'True/False  Determines if the CTRL+BREAK key combination can be used to stop code from running.
    Debug.Print , "2000", "HijriCalendar", dbs.Properties!HijriCalendar                     'True/False  Applies only to Arabic countries; determines if the application uses Hijri or Gregorian dates.
    Debug.Print ">>>View Tab"
    Debug.Print , "XP", "Show Status Bar", Application.GetOption("Show Status Bar")                             'Show, Status bar
    Debug.Print , "XP", "Show Startup Dialog Box", Application.GetOption("Show Startup Dialog Box")             'Show, Startup Task Pane
    Debug.Print , "XP", "Show New Object Shortcuts", Application.GetOption("Show New Object Shortcuts")         'Show, New object shortcuts
    Debug.Print , "XP", "Show Hidden Objects", Application.GetOption("Show Hidden Objects")                     'Show, Hidden objects
    Debug.Print , "XP", "Show System Objects", Application.GetOption("Show System Objects")                     'Show, System objects
    Debug.Print , "XP", "ShowWindowsInTaskbar", Application.GetOption("ShowWindowsInTaskbar")                   'Show, Windows in Taskbar
    Debug.Print , "XP", "Show Macro Names Column", Application.GetOption("Show Macro Names Column")             'Show in Macro Design, Names column
    Debug.Print , "XP", "Show Conditions Column", Application.GetOption("Show Conditions Column")               'Show in Macro Design, Conditions column
    Debug.Print , "XP", "Database Explorer Click Behavior", Application.GetOption("Database Explorer Click Behavior")      'Click options in database window
    Debug.Print ">>>General Tab"
    Debug.Print , "XP", "Left Margin", Application.GetOption("Left Margin")                                     'Print margins, Left margin
    Debug.Print , "XP", "Right Margin", Application.GetOption("Right Margin")                                   'Print margins, Right margin
    Debug.Print , "XP", "Top Margin", Application.GetOption("Top Margin")                                       'Print margins, Top margin
    Debug.Print , "XP", "Bottom Margin", Application.GetOption("Bottom Margin")                                 'Print margins, Bottom margin
    Debug.Print , "XP", "Four-Digit Year Formatting", Application.GetOption("Four-Digit Year Formatting")       'Use four-year digit year formatting, This database
    Debug.Print , "XP", "Four-Digit Year Formatting All Databases", Application.GetOption("Four-Digit Year Formatting All Databases")    'Use four-year digit year formatting, All databases  Four-Digit Year Formatting All Databases
    Debug.Print , "XP", "Track Name AutoCorrect Info", Application.GetOption("Track Name AutoCorrect Info")     'Name AutoCorrect, Track name AutoCorrect info
    Debug.Print , "XP", "Perform Name AutoCorrect", Application.GetOption("Perform Name AutoCorrect")           'Name AutoCorrect, Perform name AutoCorrect
    Debug.Print , "XP", "Log Name AutoCorrect Changes", Application.GetOption("Log Name AutoCorrect Changes")   'Name AutoCorrect, Log name AutoCorrect changes
    Debug.Print , "XP", "Enable MRU File List", Application.GetOption("Enable MRU File List")                   'Recently used file list
    Debug.Print , "XP", "Size of MRU File List", Application.GetOption("Size of MRU File List")                 'Recently used file list, (number of files)
    Debug.Print , "XP", "Provide Feedback with Sound", Application.GetOption("Provide Feedback with Sound")     'Provide feedback with sound
    Debug.Print , "XP", "Auto Compact", Application.GetOption("Auto Compact")                                   'Compact on Close
    Debug.Print , "XP", "New Database Sort Order", Application.GetOption("New Database Sort Order")             'New database sort order
    Debug.Print , "XP", "Remove Personal Information", Application.GetOption("Remove Personal Information")     'Remove personal information from this file
    Debug.Print , "XP", "Default Database Directory", Application.GetOption("Default Database Directory")       'Default database folder
    Debug.Print ">>>Edit/Find Tab"
    Debug.Print , "XP", "Default Find/Replace Behavior", Application.GetOption("Default Find/Replace Behavior") 'Default find/replace behavior
    Debug.Print , "XP", "Confirm Record Changes", Application.GetOption("Confirm Record Changes")               'Confirm, Record changes
    Debug.Print , "XP", "Confirm Document Deletions", Application.GetOption("Confirm Document Deletions")       'Confirm, Document deletions
    Debug.Print , "XP", "Confirm Action Queries", Application.GetOption("Confirm Action Queries")               'Confirm, Action queries
    Debug.Print , "XP", "Show Values in Indexed", Application.GetOption("Show Values in Indexed")               'Show list of values in, Local indexed fields
    Debug.Print , "XP", "Show Values in Non-Indexed", Application.GetOption("Show Values in Non-Indexed")       'Show list of values in, Local nonindexed fields
    Debug.Print , "XP", "Show Values in Remote", Application.GetOption("Show Values in Remote")                 'Show list of values in, ODBC fields
    Debug.Print , "XP", "Show Values in Snapshot", Application.GetOption("Show Values in Snapshot")             'Show list of values in, Records in local snapshot
    Debug.Print , "XP", "Show Values in Server", Application.GetOption("Show Values in Server")                 'Show list of values in, Records at server
    Debug.Print , "XP", "Show Values Limit", Application.GetOption("Show Values Limit")                         'Don't display lists where more than this number of records read
    Debug.Print ">>>Datasheet Tab"
    Debug.Print , "XP", "Default Font Color", Application.GetOption("Default Font Color")                       'Default colors, Font
    Debug.Print , "XP", "Default Background Color", Application.GetOption("Default Background Color")           'Default colors, Background
    Debug.Print , "XP", "Default Gridlines Color", Application.GetOption("Default Gridlines Color")             'Default colors, Gridlines
    Debug.Print , "XP", "Default Gridlines Horizontal", Application.GetOption("Default Gridlines Horizontal")   'Default gridlines showing, Horizontal
    Debug.Print , "XP", "Default Gridlines Vertical", Application.GetOption("Default Gridlines Vertical")       'Default gridlines showing, Vertical
    Debug.Print , "XP", "Default Column Width", Application.GetOption("Default Column Width")                   'Default column width
    Debug.Print , "XP", "Default Font Name", Application.GetOption("Default Font Name")                         'Default font, Font
    Debug.Print , "XP", "Default Font Weight", Application.GetOption("Default Font Weight")                     'Default font, Weight
    Debug.Print , "XP", "Default Font Size", Application.GetOption("Default Font Size")                         'Default font, Size
    Debug.Print , "XP", "Default Font Underline", Application.GetOption("Default Font Underline")               'Default font, Underline
    Debug.Print , "XP", "Default Font Italic", Application.GetOption("Default Font Italic")                     'Default font, Italic
    Debug.Print , "XP", "Default Cell Effect", Application.GetOption("Default Cell Effect")                     'Default cell effect
    Debug.Print , "XP", "Show Animations", Application.GetOption("Show Animations")                             'Show animations
    Debug.Print ">>>Keyboard Tab"
    Debug.Print , "XP", "Move After Enter", Application.GetOption("Move After Enter")                                   'Move after enter
    Debug.Print , "XP", "Behavior Entering Field", Application.GetOption("Behavior Entering Field")                     'Behavior entering field
    Debug.Print , "XP", "Arrow Key Behavior", Application.GetOption("Arrow Key Behavior")                               'Arrow key behavior
    Debug.Print , "XP", "Cursor Stops at First/Last Field", Application.GetOption("Cursor Stops at First/Last Field")   'Cursor stops at first/last field
    Debug.Print , "XP", "Ime Autocommit", Application.GetOption("Ime Autocommit")                                       'Auto commit
    Debug.Print , "XP", "Datasheet Ime Control", Application.GetOption("Datasheet Ime Control")                         'Datasheet IME control
    Debug.Print ">>>Tables/Queries Tab"
    Debug.Print , "XP", "Default Text Field Size", Application.GetOption("Default Text Field Size")             'Table design, Default field sizes - Text
    Debug.Print , "XP", "Default Number Field Size", Application.GetOption("Default Number Field Size")         'Table design, Default field sizes - Number
    Debug.Print , "XP", "Default Field Type", Application.GetOption("Default Field Type")                       'Table design, Default field type
    Debug.Print , "XP", "AutoIndex on Import/Create", Application.GetOption("AutoIndex on Import/Create")       'Table design, AutoIndex on Import/Create
    Debug.Print , "XP", "Show Table Names", Application.GetOption("Show Table Names")                           'Query design, Show table names
    Debug.Print , "XP", "Output All Fields", Application.GetOption("Output All Fields")                         'Query design, Output all fields
    Debug.Print , "XP", "Enable AutoJoin", Application.GetOption("Enable AutoJoin")                             'Query design, Enable AutoJoin
    Debug.Print , "XP", "Run Permissions", Application.GetOption("Run Permissions")                             'Query design, Run permissions
    Debug.Print , "XP", "ANSI Query Mode", Application.GetOption("ANSI Query Mode")                             'Query design, SQL Server Compatible Syntax (ANSI 92) - This database
    Debug.Print , "XP", "ANSI Query Mode Default", Application.GetOption("ANSI Query Mode Default")             'Query design, SQL Server Compatible Syntax (ANSI 92) - Default for new databases
    Debug.Print ">>>Forms/Reports Tab"
    Debug.Print , "XP", "Selection Behavior", Application.GetOption("Selection Behavior")                       'Selection behavior
    Debug.Print , "XP", "Form Template", Application.GetOption("Form Template")                                 'Form template
    Debug.Print , "XP", "Report Template", Application.GetOption("Report Template")                             'Report template
    Debug.Print , "XP", "Always Use Event Procedures", Application.GetOption("Always Use Event Procedures")     'Always use event procedures
    Debug.Print ">>>Advanced Tab"
    Debug.Print , "XP", "Ignore DDE Requests", Application.GetOption("Ignore DDE Requests")                         'DDE operations, Ignore DDE requests
    Debug.Print , "XP", "Enable DDE Refresh", Application.GetOption("Enable DDE Refresh")                           'DDE operations, Enable DDE refresh
    Debug.Print , "XP", "Default File Format", Application.GetOption("Default File Format")                         'Default File Format
    Debug.Print , "XP", "Row Limit", Application.GetOption("Row Limit")                                             'Client-server settings, Default max records
    Debug.Print , "XP", "Default Open Mode for Databases", Application.GetOption("Default Open Mode for Databases") 'Default open mode
    Debug.Print , "XP", "Command-Line Arguments", Application.GetOption("Command-Line Arguments")                   'Command-line arguments
    Debug.Print , "XP", "OLE/DDE Timeout (sec)", Application.GetOption("OLE/DDE Timeout (sec)")                     'OLE/DDE timeout
    Debug.Print , "XP", "Default Record Locking", Application.GetOption("Default Record Locking")                   'Default record locking
    Debug.Print , "XP", "Refresh Interval (sec)", Application.GetOption("Refresh Interval (sec)")                   'Refresh interval
    Debug.Print , "XP", "Number of Update Retries", Application.GetOption("Number of Update Retries")               'Number of update retries
    Debug.Print , "XP", "ODBC Refresh Interval (sec)", Application.GetOption("ODBC Refresh Interval (sec)")         'ODBC fresh interval
    Debug.Print , "XP", "Update Retry Interval (msec)", Application.GetOption("Update Retry Interval (msec)")       'Update retry interval
    Debug.Print , "XP", "Use Row Level Locking", Application.GetOption("Use Row Level Locking")                     'Open databases using record-level locking
    Debug.Print , "XP", "Save Login and Password", Application.GetOption("Save Login and Password")                 'Save login and password
    Debug.Print ">>>Pages Tab"
    Debug.Print , "XP", "Section Indent", Application.GetOption("Section Indent")                                   'Default Designer Properties, Section Indent
    Debug.Print , "XP", "Alternate Row Color", Application.GetOption("Alternate Row Color")                         'Default Designer Properties, Alternative Row Color
    Debug.Print , "XP", "Caption Section Style", Application.GetOption("Caption Section Style")                     'Default Designer Properties, Caption Section Style
    Debug.Print , "XP", "Footer Section Style", Application.GetOption("Footer Section Style")                       'Default Designer Properties, Footer Section Style
    Debug.Print , "XP", "Use Default Page Folder", Application.GetOption("Use Default Page Folder")                 'Default Database/Project Properties, Use Default Page Folder
    Debug.Print , "XP", "Default Page Folder", Application.GetOption("Default Page Folder")                         'Default Database/Project Properties, Default Page Folder
    Debug.Print , "XP", "Use Default Connection File", Application.GetOption("Use Default Connection File")         'Default Database/Project Properties, Use Default Connection File
    Debug.Print , "XP", "Default Connection File", Application.GetOption("Default Connection File")                 'Default Database/Project Properties, Default Connection File
    Debug.Print ">>>Spelling Tab"
'Dictionary Language Spelling dictionary language
'Add words to    Spelling add words to
'Suggest from main dictionary only   Spelling suggest from main dictionary only
'Ignore words in UPPERCASE   Spelling ignore words in UPPERCASE
'Ignore words with numbers   Spelling ignore words with number
'Ignore Internet and file addresses  Spelling ignore Internet and file addresses
'Language-specific, German: Use post-reform rules    Spelling use German post-reform rules
'Language-specific, Korean: Combine aux verb/adj.    Spelling combine aux verb/adj
'Language-specific, Korean: Use auto-change list Spelling use auto-change list
'Language-specific, Korean: Process compound nouns   Spelling process compound nouns
'Language-specific, Hebrew modes Spelling Hebrew modes
'Language-specific, Arabic modes Spelling Arabic modes
    Debug.Print ">>>Creating databases section"
    Debug.Print , "Default File Format", Application.GetOption("Default File Format")
    Debug.Print , "Default Database Directory", Application.GetOption("Default Database Directory")
    Debug.Print , "New Database Sort Order", Application.GetOption("New Database Sort Order")
    Debug.Print ">>>Application Options section"
    Debug.Print , "Auto Compact", Application.GetOption("Auto Compact")
    Debug.Print , "Remove Personal Information", Application.GetOption("Remove Personal Information")
    Debug.Print , "Themed Form Controls", Application.GetOption("Themed Form Controls")
    Debug.Print , "DesignWithData", Application.GetOption("DesignWithData")
    Debug.Print , "CheckTruncatedNumFields", Application.GetOption("CheckTruncatedNumFields")
    Debug.Print , "Picture Property Storage Format", Application.GetOption("Picture Property Storage Format")
    Debug.Print ">>>Name AutoCorrect Options section"
    Debug.Print , "Track Name AutoCorrect Info", Application.GetOption("Track Name AutoCorrect Info")
    Debug.Print , "Perform Name AutoCorrect", Application.GetOption("Perform Name AutoCorrect")
    Debug.Print , "Log Name AutoCorrect Changes", Application.GetOption("Log Name AutoCorrect Changes")
    Debug.Print ">>>Filter Lookup options for <Database Name> Database section"
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")
    Debug.Print "", Application.GetOption("")

    Set dbs = Nothing

End Sub