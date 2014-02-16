Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Copyright (c) 2011 Peter F. Ennis
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation;
'version 3.0.
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, visit
'http://www.gnu.org/licenses/lgpl-3.0.txt
'
' Ref: http://www.di-mgt.com.au/cl_Simple.html
'=======================================================================
' Author:   Peter F. Ennis
' Date:     February 24, 2011
' Comment:  Create class for revision control
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'=======================================================================

Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Private Const aegitVERSION As String = "0.7.7"
Private Const aegitVERSION_DATE As String = "February 15, 2014"
Private Const THE_DRIVE As String = "C"

Private Const gcfHandleErrors As Boolean = True

' Ref: http://www.cpearson.com/excel/sizestring.htm
' This enum is used by SizeString to indicate whether the supplied text
' appears on the left or right side of result string.
Private Enum SizeStringSide
    TextLeft = 1
    TextRight = 2
End Enum

Private Type mySetupType
    SourceFolder As String
    ImportFolder As String
    UseImportFolder As Boolean
    XMLFolder As String
End Type

' Current pointer to the array element of the call stack
Private mintStackPointer As Integer
' Array of procedure names in the call stack
Private mastrCallStack() As String
' The number of elements to increase the array
Private Const mcintIncrementStackSize As Integer = 10
Private mfInErrorHandler As Boolean

Private aegitSetup As Boolean
Private aegitType As mySetupType
Private aegitSourceFolder As String
Private aegitImportFolder As String
Private aegitXMLFolder As String
Private aegitDataXML() As Variant
Private aegitExportDataToXML As Boolean
Private aegitUseImportFolder As Boolean
Private aegitblnCustomSourceFolder As Boolean
Private aestrSourceLocation As String
Private aestrImportLocation As String
Private aestrXMLLocation As String
Private aeintLTN As Long                        ' Longest Table Name
Private aestrLFN As String                      ' Longest Field Name
Private aestrLFNTN As String
Private aeintFNLen As Long
Private aestrLFT As String                      ' Longest Field Type
Private aeintFTLen As Long                      ' Field Type Length
Private Const aeintFSize As Long = 4
Private aeintFDLen As Long
Private aestrLFD As String
Private Const aestr4 As String = "    "
Private Const aeSqlTxtFile = "OutputSqlCodeForQueries.txt"
Private Const aeTblTxtFile = "OutputTblSetupForTables.txt"
Private Const aeTblXMLFile = "OutputTblXMLSetupForTables.txt"
Private Const aeRefTxtFile = "OutputReferencesSetup.txt"
Private Const aeRelTxtFile = "OutputRelationsSetup.txt"
Private Const aePrpTxtFile = "OutputPropertiesBuiltIn.txt"
Private Const aeFLkCtrFile = "OutputFieldLookupControlTypeList.txt"
Private Const aeSchemaFile = "OutputSchemaFile.txt"
'

Private Sub Class_Initialize()
' Ref: http://www.cadalyst.com/cad/autocad/programming-with-class-part-1-5050
' Ref: http://www.bigresource.com/Tracker/Track-vb-cyJ1aJEyKj/
' Ref: http://stackoverflow.com/questions/1731052/is-there-a-way-to-overload-the-constructor-initialize-procedure-for-a-class-in

    ' provide a default value for the SourceFolder, ImportFolder and other properties
    aegitSourceFolder = "default"
    aegitImportFolder = "default"
    aegitXMLFolder = "default"
    ReDim Preserve aegitDataXML(1 To 1)
    aegitDataXML(1) = "tlkpStates"
    aegitUseImportFolder = False
    aegitExportDataToXML = False
    aegitType.SourceFolder = "C:\ae\aegit\aerc\src\"
    aegitType.ImportFolder = "C:\ae\aegit\aerc\imp\"
    aegitType.UseImportFolder = False
    aegitType.XMLFolder = "C:\ae\aegit\aerc\xml\"
    aeintLTN = LongestTableName
    LongestFieldPropsName

    Debug.Print "Class_Initialize"
    Debug.Print , "Default for aegitSourceFolder = " & aegitSourceFolder
    Debug.Print , "Default for aegitImportFolder = " & aegitImportFolder
    Debug.Print , "Default for aegitType.SourceFolder = " & aegitType.SourceFolder
    Debug.Print , "Default for aegitType.ImportFolder = " & aegitType.ImportFolder
    Debug.Print , "Default for aegitType.UseImportFolder = " & aegitType.UseImportFolder
    Debug.Print , "Default for aegitType.XMLFolder = " & aegitType.XMLFolder
    Debug.Print , "aeintLTN = " & aeintLTN
    Debug.Print , "aeintFNLen = " & aeintFNLen
    Debug.Print , "aeintFTLen = " & aeintFTLen
    Debug.Print , "aeintFSize = " & aeintFSize
    'Debug.Print , "aeintFDLen = " & aeintFDLen

End Sub

Private Sub Class_Terminate()
    Dim strFile As String
    strFile = aegitSourceFolder & "export.ini"
    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
    End If
    Debug.Print
    Debug.Print "Class_Terminate"
    Debug.Print , "aegit VERSION: " & aegitVERSION
    Debug.Print , "aegit VERSION_DATE: " & aegitVERSION_DATE
End Sub

Property Get SourceFolder() As String
    SourceFolder = aegitSourceFolder
End Property

Property Let SourceFolder(ByVal strSourceFolder As String)
    ' Ref: http://www.techrepublic.com/article/build-your-skills-using-class-modules-in-an-access-database-solution/5031814
    ' Ref: http://www.utteraccess.com/wiki/index.php/Classes
    aegitSourceFolder = strSourceFolder
End Property

Property Get ImportFolder() As String
    ImportFolder = aegitImportFolder
End Property

Property Let ImportFolder(ByVal strImportFolder As String)
    aegitImportFolder = strImportFolder
End Property

Property Let UseImportFolder(ByVal blnUseImportFolder As Boolean)
    aegitUseImportFolder = blnUseImportFolder
End Property

Property Get XMLFolder() As String
    XMLFolder = aegitXMLFolder
End Property

Property Let XMLFolder(ByVal strXMLFolder As String)
    aegitXMLFolder = strXMLFolder
End Property

Property Let TablesExportToXML(ByRef avarTables() As Variant)
    MsgBox "Let TablesExportToXML: LBound(aegitDataXML())=" & LBound(aegitDataXML()) & _
        vbCrLf & "UBound(aegitDataXML())=" & UBound(aegitDataXML()), vbInformation, "CHECK"
    'aegitDataXML = avarTables
End Property

Property Get DocumentTheDatabase(Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get DocumentTheDatabase"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentTheDatabase = aeDocumentTheDatabase
    Else
        Debug.Print "Get DocumentTheDatabase"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTheDatabase = aeDocumentTheDatabase(DebugTheCode)
    End If
End Property

Property Get Exists(strAccObjType As String, _
                        strAccObjName As String, _
                        Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get Exists"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeExists"
        Debug.Print , "DEBUGGING IS OFF"
        Exists = aeExists(strAccObjType, strAccObjName)
    Else
        Debug.Print "Get Exists"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeExists"
        Debug.Print , "DEBUGGING TURNED ON"
        Exists = aeExists(strAccObjType, strAccObjName, DebugTheCode)
    End If
End Property

Property Get ReadDocDatabase(blnImport As Boolean, Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get ReadDocDatabase"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeReadDocDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        ReadDocDatabase = aeReadDocDatabase(blnImport)
    Else
        Debug.Print "Get ReadDocDatabase"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeReadDocDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
        ReadDocDatabase = aeReadDocDatabase(blnImport, DebugTheCode)
    End If
End Property

Property Get GetReferences(Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get GetReferences"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeGetReferences"
        Debug.Print , "DEBUGGING IS OFF"
        GetReferences = aeGetReferences
    Else
        Debug.Print "Get GetReferences"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeGetReferences"
        Debug.Print , "DEBUGGING TURNED ON"
        GetReferences = aeGetReferences(DebugTheCode)
    End If
End Property

Property Get DocumentRelations(Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get DocumentRelations"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeDocumentRelations"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentRelations = aeDocumentRelations
    Else
        Debug.Print "Get DocumentRelations"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeDocumentRelations"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentRelations = aeDocumentRelations(DebugTheCode)
    End If
End Property

Property Get DocumentTables(Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get DocumentTables"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeDocumentTables"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentTables = aeDocumentTables
    Else
        Debug.Print "Get DocumentTables"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeDocumentTables"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTables = aeDocumentTables(DebugTheCode)
    End If
End Property

Property Get DocumentTablesXML(Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get DocumentTablesXML"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeDocumentTablesXML"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentTablesXML = aeDocumentTablesXML
    Else
        Debug.Print "Get DocumentTablesXML"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeDocumentTablesXML"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTablesXML = aeDocumentTablesXML(DebugTheCode)
    End If
End Property

Property Get CompactAndRepair(Optional varTrueFalse As Variant) As Boolean
' Automation for Compact and Repair

    Dim blnRun As Boolean

    Debug.Print "CompactAndRepair"
    If Not IsMissing(varTrueFalse) Then
        blnRun = False
        Debug.Print , "varTrueFalse IS NOT MISSING so blnRun of CompactAndRepair is set to False"
        Debug.Print , "RUN CompactAndRepair IS OFF"
    Else
        blnRun = True
        Debug.Print , "varTrueFalse IS MISSING so blnRun of CompactAndRepair is set to True"
        Debug.Print , "RUN CompactAndRepair IS ON..."
    End If

' TableDefs not refreshed after create
' Ref: http://support.microsoft.com/kb/104339
' So force a compact and repair
' Ref: http://msdn.microsoft.com/en-us/library/office/aa202943(v=office.10).aspx
' Not a "good practice" but for this use it is simple and works
' From the Access window
' Access 2003: SendKeys "%(TDC)", False
' Access 2007: SendKeys "%(FMC)", False
' Access 2010: SendKeys "%(YC)", False
' From the Immediate window
    
    If blnRun Then
        ' Close VBA
        SendKeys "%F{END}{ENTER}", False
        ' Run Compact and Repair
        SendKeys "%F{TAB}{TAB}{ENTER}", False
        CompactAndRepair = True
    Else
        CompactAndRepair = False
    End If
    
End Property

Private Function aeReadWriteStream(strPathFileName As String) As Boolean

    'Debug.Print "aeReadWriteStream Entry strPathFileName=" & strPathFileName

    ' Use a call stack and global error handler
    'If gcfHandleErrors Then On Error GoTo PROC_ERR
    'PushCallStack "aeReadWriteStream"

    On Error GoTo PROC_ERR

    Dim fName As String
    Dim fname2 As String
    Dim fnr As Integer
    Dim fnr2 As Integer
    Dim tstring As String * 1
    Dim i As Integer

    aeReadWriteStream = False

    ' If the file has no Byte Order Mark (BOM)
    ' Ref: http://msdn.microsoft.com/en-us/library/windows/desktop/dd374101%28v=vs.85%29.aspx
    ' then do nothing
    fName = strPathFileName
    fname2 = strPathFileName & ".clean.txt"

    fnr = FreeFile()
    Open fName For Binary Access Read As #fnr
    Get #fnr, , tstring
    ' #FFFE, #FFFF, #0000
    ' If no BOM then it is a txt file and header stripping is not needed
    If Asc(tstring) <> 254 And Asc(tstring) <> 255 And _
                Asc(tstring) <> 0 Then
        Close #fnr
        aeReadWriteStream = False
        Exit Function
    End If

    fnr2 = FreeFile()
    Open fname2 For Binary Lock Read Write As #fnr2

    While Not EOF(fnr)
        Get #fnr, , tstring
        If Asc(tstring) = 254 Or Asc(tstring) = 255 Or _
                Asc(tstring) = 0 Then
        Else
            Put #fnr2, , tstring
        End If
    Wend

PROC_EXIT:
    Close #fnr
    Close #fnr2
    aeReadWriteStream = True
    'PopCallStack
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 9
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadWriteStream of Class aegitClass" & _
                    vbCrLf & "aeReadWriteStream Entry strPathFileName=" & strPathFileName, vbCritical, "aeReadWriteStream ERROR=9"
            'If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadWriteStream of Class aegitClass"
            'GlobalErrHandler
            Resume Next
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadWriteStream of Class aegitClass"
            'GlobalErrHandler
            Resume Next
    End Select

End Function

Private Sub ListOfAllHiddenQueries(Optional varDebug As Variant)
' Ref: http://www.pcreview.co.uk/forums/runtime-error-7874-a-t2922352.html

    Const strTempTable As String = "zzzTmpTblQueries"
    ' NOTE: Use zzz* for the table name so that it will be ignored by aegit code export if it exists
    Const strSQL As String = "SELECT m.Name INTO " & strTempTable & " " & vbCrLf & _
                                "FROM MSysObjects AS m " & vbCrLf & _
                                "WHERE (((m.Name) Not ALike ""~%"") AND ((IIf(IsQryHidden([Name]),1,0))=1) AND ((m.Type)=5)) " & vbCrLf & _
                                "ORDER BY m.Name;"
    
'    "SELECT m.Name, IIf(IsQryHidden([Name]),1,0) AS Hidden INTO " & strTempTable & " " & vbCrLf & _
'                                "FROM MSysObjects AS m " & vbCrLf & _
'                                "WHERE (((m.Name) Not ALike ""~%"") AND ((IIf(IsQryHidden([Name]),1,0))=1) AND ((m.Type)=5)) " & vbCrLf & _
'                                "ORDER BY m.Name;"

    ' RunSQL works for Action queries
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    If Not IsMissing(varDebug) Then Debug.Print "The number of hidden queries in the database is: " & DCount("Name", strTempTable)
    DoCmd.TransferText acExportDelim, "", strTempTable, aestrSourceLocation & "ListOfAllHiddenQueries.txt", False
    CurrentDb.Execute "DROP TABLE " & strTempTable
    DoCmd.SetWarnings True

End Sub

Private Sub ListOfAccessApplicationOptions(Optional varDebug As Variant)

' Note: If you are developing a database application, add-in, library database, or referenced database, make sure that the
' Error Trapping option is set to 2 (Break On Unhandled Errors) when you have finished debugging your code.
'
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
'
' Ref: http://www.fmsinc.com/tpapers/vbacode/debug.asp
' Break on Unhandled Errors: works in most cases but is problematic while debugging class modules.
' During development, if Error Trapping is set to 'Break on Unhandled Errors' and an error occurs in a class module,
' the debugger stops on the line calling the class rather than the offending line in the class.
' This makes finding and fixing the problem a real pain.

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "ListOfAccessApplicationOptions"

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim str As String
    Dim fle As Integer
    Dim Debugit As Boolean

    If IsMissing(varDebug) Then
        Debug.Print "ListOfAccessApplicationOptions"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to ListOfAccessApplicationOptions"
        Debug.Print , "DEBUGGING IS OFF"
        Debugit = False
    Else
        Debug.Print "ListOfAccessApplicationOptions"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to ListOfAccessApplicationOptions"
        Debug.Print , "DEBUGGING TURNED ON"
        Debugit = True
    End If

    fle = FreeFile()
    If Not IsMissing(varDebug) Then Debug.Print "aegitSourceFolder=" & aegitSourceFolder
    'Stop

    If aegitSourceFolder = "default" Then
        aegitSourceFolder = aegitType.SourceFolder
    End If

    Open aegitSourceFolder & "\ListOfAccessApplicationOptions.txt" For Output As #fle

    'On Error Resume Next
    Print #fle, ">>>Standard Options"
    '2000 The following options are equivalent to the standard startup options found in the Startup Options dialog box.
    Print #fle, , "2000", "AppTitle              ", dbs.Properties!AppTitle                     'String  The title of an application, as displayed in the title bar.
    Print #fle, , "2000", "AppIcon               ", dbs.Properties!AppIcon                      'String  The file name and path of an application's icon.
    Print #fle, , "2000", "StartupMenuBar        ", dbs.Properties!StartUpMenuBar               'String  Sets the default menu bar for the application.
    Print #fle, , "2000", "AllowFullMenus        ", dbs.Properties!AllowFullMenus               'True/False  Determines if the built-in Access menu bars are displayed.
    Print #fle, , "2000", "AllowShortcutMenus    ", dbs.Properties!AllowShortcutMenus           'True/False  Determines if the built-in Access shortcut menus are displayed.
    Print #fle, , "2000", "StartupForm           ", dbs.Properties!StartUpForm                  'String  Sets the form or data page to show when the application is first opened.
    Print #fle, , "2000", "StartupShowDBWindow   ", dbs.Properties!StartUpShowDBWindow          'True/False  Determines if the database window is displayed when the application is first opened.
    Print #fle, , "2000", "StartupShowStatusBar  ", dbs.Properties!StartUpShowStatusBar         'True/False  Determines if the status bar is displayed.
    Print #fle, , "2000", "StartupShortcutMenuBar", dbs.Properties!StartUpShortcutMenuBar       'String  Sets the shortcut menu bar to be used in all forms and reports.
    Print #fle, , "2000", "AllowBuiltInToolbars  ", dbs.Properties!AllowBuiltInToolbars         'True/False  Determines if the built-in Access toolbars are displayed.
    Print #fle, , "2000", "AllowToolbarChanges   ", dbs.Properties!AllowToolbarChanges          'True/False  Determined if toolbar changes can be made.
    Print #fle, ">>>Advanced Option"
    Print #fle, , "2000", "AllowSpecialKeys      ", dbs.Properties!AllowSpecialKeys             'option (True/False value) determines if the use of special keys is permitted. It is equivalent to the advanced startup option found in the Startup Options dialog box.
    Print #fle, ">>>Extra Options"
    'The following options are not available from the Startup Options dialog box or any other Access user interface component, they are only available in programming code.
    Print #fle, , "2000", "AllowBypassKey        ", dbs.Properties!AllowBypassKey               'True/False  Determines if the SHIFT key can be used to bypass the application load process.
    Print #fle, , "2000", "AllowBreakIntoCode    ", dbs.Properties!AllowBreakIntoCode           'True/False  Determines if the CTRL+BREAK key combination can be used to stop code from running.
    Print #fle, , "2000", "HijriCalendar         ", dbs.Properties!HijriCalendar                'True/False  Applies only to Arabic countries; determines if the application uses Hijri or Gregorian dates.
    Print #fle, ">>>View Tab"
    Print #fle, , "XP, 2003", "Show Status Bar                 ", Application.GetOption("Show Status Bar")                    'Show, Status bar
    Print #fle, , "XP, 2003", "Show Startup Dialog Box         ", Application.GetOption("Show Startup Dialog Box")            'Show, Startup Task Pane
    Print #fle, , "XP, 2003", "Show New Object Shortcuts       ", Application.GetOption("Show New Object Shortcuts")          'Show, New object shortcuts
    Print #fle, , "XP, 2003", "Show Hidden Objects             ", Application.GetOption("Show Hidden Objects")                'Show, Hidden objects
    Print #fle, , "XP, 2003", "Show System Objects             ", Application.GetOption("Show System Objects")                'Show, System objects
    Print #fle, , "XP, 2003", "ShowWindowsInTaskbar            ", Application.GetOption("ShowWindowsInTaskbar")               'Show, Windows in Taskbar
    Print #fle, , "XP, 2003", "Show Macro Names Column         ", Application.GetOption("Show Macro Names Column")            'Show in Macro Design, Names column
    Print #fle, , "XP, 2003", "Show Conditions Column          ", Application.GetOption("Show Conditions Column")             'Show in Macro Design, Conditions column
    Print #fle, , "XP, 2003", "Database Explorer Click Behavior", Application.GetOption("Database Explorer Click Behavior")   'Click options in database window
    Print #fle, ">>>General Tab"
    Print #fle, , "XP, 2003", "Left Margin                 ", Application.GetOption("Left Margin")                                            'Print margins, Left margin
    Print #fle, , "XP, 2003", "Right Margin                ", Application.GetOption("Right Margin")                                           'Print margins, Right margin
    Print #fle, , "XP, 2003", "Top Margin                  ", Application.GetOption("Top Margin")                                             'Print margins, Top margin
    Print #fle, , "XP, 2003", "Bottom Margin               ", Application.GetOption("Bottom Margin")                                          'Print margins, Bottom margin
    Print #fle, , "XP, 2003", "Four-Digit Year Formatting  ", Application.GetOption("Four-Digit Year Formatting")                             'Use four-year digit year formatting, This database
    Print #fle, , "XP, 2003", "Four-Digit Year Formatting All Databases", Application.GetOption("Four-Digit Year Formatting All Databases")   'Use four-year digit year formatting, All databases  Four-Digit Year Formatting All Databases
    Print #fle, , "XP, 2003", "Track Name AutoCorrect Info ", Application.GetOption("Track Name AutoCorrect Info")                            'Name AutoCorrect, Track name AutoCorrect info
    Print #fle, , "XP, 2003", "Perform Name AutoCorrect    ", Application.GetOption("Perform Name AutoCorrect")                               'Name AutoCorrect, Perform name AutoCorrect
    Print #fle, , "XP, 2003", "Log Name AutoCorrect Changes", Application.GetOption("Log Name AutoCorrect Changes")                           'Name AutoCorrect, Log name AutoCorrect changes
    Print #fle, , "XP, 2003", "Enable MRU File List        ", Application.GetOption("Enable MRU File List")                                   'Recently used file list
    Print #fle, , "XP, 2003", "Size of MRU File List       ", Application.GetOption("Size of MRU File List")                                  'Recently used file list, (number of files)
    Print #fle, , "XP, 2003", "Provide Feedback with Sound ", Application.GetOption("Provide Feedback with Sound")                            'Provide feedback with sound
    Print #fle, , "XP, 2003", "Auto Compact                ", Application.GetOption("Auto Compact")                                           'Compact on Close
    Print #fle, , "XP, 2003", "New Database Sort Order     ", Application.GetOption("New Database Sort Order")                                'New database sort order
    Print #fle, , "XP, 2003", "Remove Personal Information ", Application.GetOption("Remove Personal Information")                            'Remove personal information from this file
    Print #fle, , "XP, 2003", "Default Database Directory  ", Application.GetOption("Default Database Directory")                             'Default database folder
    Print #fle, ">>>Edit/Find Tab"
    Print #fle, , "XP, 2003", "Default Find/Replace Behavior", Application.GetOption("Default Find/Replace Behavior")       'Default find/replace behavior
    Print #fle, , "XP, 2003", "Confirm Record Changes       ", Application.GetOption("Confirm Record Changes")              'Confirm, Record changes
    Print #fle, , "XP, 2003", "Confirm Document Deletions   ", Application.GetOption("Confirm Document Deletions")          'Confirm, Document deletions
    Print #fle, , "XP, 2003", "Confirm Action Queries       ", Application.GetOption("Confirm Action Queries")              'Confirm, Action queries
    Print #fle, , "XP, 2003", "Show Values in Indexed       ", Application.GetOption("Show Values in Indexed")              'Show list of values in, Local indexed fields
    Print #fle, , "XP, 2003", "Show Values in Non-Indexed   ", Application.GetOption("Show Values in Non-Indexed")          'Show list of values in, Local nonindexed fields
    Print #fle, , "XP, 2003", "Show Values in Remote        ", Application.GetOption("Show Values in Remote")               'Show list of values in, ODBC fields
    Print #fle, , "XP, 2003", "Show Values in Snapshot      ", Application.GetOption("Show Values in Snapshot")             'Show list of values in, Records in local snapshot
    Print #fle, , "XP, 2003", "Show Values in Server        ", Application.GetOption("Show Values in Server")               'Show list of values in, Records at server
    Print #fle, , "XP, 2003", "Show Values Limit            ", Application.GetOption("Show Values Limit")                   'Don't display lists where more than this number of records read
    Print #fle, ">>>Datasheet Tab"
    Print #fle, , "XP, 2003", "Default Font Color           ", Application.GetOption("Default Font Color")                  'Default colors, Font
    Print #fle, , "XP, 2003", "Default Background Color     ", Application.GetOption("Default Background Color")            'Default colors, Background
    Print #fle, , "XP, 2003", "Default Gridlines Color      ", Application.GetOption("Default Gridlines Color")             'Default colors, Gridlines
    Print #fle, , "XP, 2003", "Default Gridlines Horizontal ", Application.GetOption("Default Gridlines Horizontal")        'Default gridlines showing, Horizontal
    Print #fle, , "XP, 2003", "Default Gridlines Vertical   ", Application.GetOption("Default Gridlines Vertical")          'Default gridlines showing, Vertical
    Print #fle, , "XP, 2003", "Default Column Width         ", Application.GetOption("Default Column Width")                'Default column width
    Print #fle, , "XP, 2003", "Default Font Name            ", Application.GetOption("Default Font Name")                   'Default font, Font
    Print #fle, , "XP, 2003", "Default Font Weight          ", Application.GetOption("Default Font Weight")                 'Default font, Weight
    Print #fle, , "XP, 2003", "Default Font Size            ", Application.GetOption("Default Font Size")                   'Default font, Size
    Print #fle, , "XP, 2003", "Default Font Underline       ", Application.GetOption("Default Font Underline")              'Default font, Underline
    Print #fle, , "XP, 2003", "Default Font Italic          ", Application.GetOption("Default Font Italic")                 'Default font, Italic
    Print #fle, , "XP, 2003", "Default Cell Effect          ", Application.GetOption("Default Cell Effect")                 'Default cell effect
    Print #fle, , "XP, 2003", "Show Animations              ", Application.GetOption("Show Animations")                     'Show animations
    Print #fle, , "    2003", "Show Smart Tags on Datasheets", Application.GetOption("Show Smart Tags on Datasheets")       'Show Smart Tags on Datasheets
    Print #fle, ">>>Keyboard Tab"
    Print #fle, , "XP, 2003", "Move After Enter                ", Application.GetOption("Move After Enter")                   'Move after enter
    Print #fle, , "XP, 2003", "Behavior Entering Field         ", Application.GetOption("Behavior Entering Field")            'Behavior entering field
    Print #fle, , "XP, 2003", "Arrow Key Behavior              ", Application.GetOption("Arrow Key Behavior")                 'Arrow key behavior
    Print #fle, , "XP, 2003", "Cursor Stops at First/Last Field", Application.GetOption("Cursor Stops at First/Last Field")   'Cursor stops at first/last field
    Print #fle, , "XP, 2003", "Ime Autocommit                  ", Application.GetOption("Ime Autocommit")                     'Auto commit
    Print #fle, , "XP, 2003", "Datasheet Ime Control           ", Application.GetOption("Datasheet Ime Control")              'Datasheet IME control
    Print #fle, ">>>Tables/Queries Tab"
    Print #fle, , "XP, 2003", "Default Text Field Size             ", Application.GetOption("Default Text Field Size")              'Table design, Default field sizes - Text
    Print #fle, , "XP, 2003", "Default Number Field Size           ", Application.GetOption("Default Number Field Size")            'Table design, Default field sizes - Number
    Print #fle, , "XP, 2003", "Default Field Type                  ", Application.GetOption("Default Field Type")                   'Table design, Default field type
    Print #fle, , "XP, 2003", "AutoIndex on Import/Create          ", Application.GetOption("AutoIndex on Import/Create")           'Table design, AutoIndex on Import/Create
    Print #fle, , "XP, 2003", "Show Table Names                    ", Application.GetOption("Show Table Names")                     'Query design, Show table names
    Print #fle, , "XP, 2003", "Output All Fields                   ", Application.GetOption("Output All Fields")                    'Query design, Output all fields
    Print #fle, , "XP, 2003", "Enable AutoJoin                     ", Application.GetOption("Enable AutoJoin")                      'Query design, Enable AutoJoin
    Print #fle, , "XP, 2003", "Run Permissions                     ", Application.GetOption("Run Permissions")                      'Query design, Run permissions
    Print #fle, , "XP, 2003", "ANSI Query Mode                     ", Application.GetOption("ANSI Query Mode")                      'Query design, SQL Server Compatible Syntax (ANSI 92) - This database
    Print #fle, , "XP, 2003", "ANSI Query Mode Default             ", Application.GetOption("ANSI Query Mode Default")              'Query design, SQL Server Compatible Syntax (ANSI 92) - Default for new databases
    Print #fle, , "    2003", "Query Design Font Name              ", Application.GetOption("Query Design Font Name")               'Query design, Query design font, Font
    Print #fle, , "    2003", "Query Design Font Size              ", Application.GetOption("Query Design Font Size")               'Query design, Query design font, Size
    Print #fle, , "    2003", "Show Property Update Options buttons", Application.GetOption("Show Property Update Options buttons") 'Show Property Update Options buttons
    Print #fle, ">>>Forms/Reports Tab"
    Print #fle, , "XP, 2003", "Selection Behavior         ", Application.GetOption("Selection Behavior")              'Selection behavior
    Print #fle, , "XP, 2003", "Form Template              ", Application.GetOption("Form Template")                   'Form template
    Print #fle, , "XP, 2003", "Report Template            ", Application.GetOption("Report Template")                 'Report template
    Print #fle, , "XP, 2003", "Always Use Event Procedures", Application.GetOption("Always Use Event Procedures")     'Always use event procedures
    Print #fle, , "    2003", "Show Smart Tags on Forms   ", Application.GetOption("Show Smart Tags on Forms")        'Show Smart Tags on Forms
    Print #fle, , "    2003", "Themed Form Controls       ", Application.GetOption("Themed Form Controls")            'Show Windows Themed Controls on Forms
    Print #fle, ">>>Advanced Tab"
    Print #fle, , "XP, 2003", "Ignore DDE Requests            ", Application.GetOption("Ignore DDE Requests")             'DDE operations, Ignore DDE requests
    Print #fle, , "XP, 2003", "Enable DDE Refresh             ", Application.GetOption("Enable DDE Refresh")              'DDE operations, Enable DDE refresh
    Print #fle, , "XP, 2003", "Default File Format            ", Application.GetOption("Default File Format")             'Default File Format
    Print #fle, , "XP      ", "Row Limit                      ", Application.GetOption("Row Limit")                       'Client-server settings, Default max records
    Print #fle, , "XP, 2003", "Default Open Mode for Databases", Application.GetOption("Default Open Mode for Databases") 'Default open mode
    Print #fle, , "XP, 2003", "Command-Line Arguments         ", Application.GetOption("Command-Line Arguments")          'Command-line arguments
    Print #fle, , "XP, 2003", "OLE/DDE Timeout (sec)          ", Application.GetOption("OLE/DDE Timeout (sec)")           'OLE/DDE timeout
    Print #fle, , "XP, 2003", "Default Record Locking         ", Application.GetOption("Default Record Locking")          'Default record locking
    Print #fle, , "XP, 2003", "Refresh Interval (sec)         ", Application.GetOption("Refresh Interval (sec)")          'Refresh interval
    Print #fle, , "XP, 2003", "Number of Update Retries       ", Application.GetOption("Number of Update Retries")        'Number of update retries
    Print #fle, , "XP, 2003", "ODBC Refresh Interval (sec)    ", Application.GetOption("ODBC Refresh Interval (sec)")     'ODBC fresh interval
    Print #fle, , "XP, 2003", "Update Retry Interval (msec)   ", Application.GetOption("Update Retry Interval (msec)")    'Update retry interval
    Print #fle, , "XP, 2003", "Use Row Level Locking          ", Application.GetOption("Use Row Level Locking")           'Open databases using record-level locking
    Print #fle, , "XP      ", "Save Login and Password        ", Application.GetOption("Save Login and Password")         'Save login and password
    Print #fle, ">>>Pages Tab"
    Print #fle, , "XP, 2003", "Section Indent             ", Application.GetOption("Section Indent")                      'Default Designer Properties, Section Indent
    Print #fle, , "XP, 2003", "Alternate Row Color        ", Application.GetOption("Alternate Row Color")                 'Default Designer Properties, Alternative Row Color
    Print #fle, , "XP, 2003", "Caption Section Style      ", Application.GetOption("Caption Section Style")               'Default Designer Properties, Caption Section Style
    Print #fle, , "XP, 2003", "Footer Section Style       ", Application.GetOption("Footer Section Style")                'Default Designer Properties, Footer Section Style
    Print #fle, , "XP, 2003", "Use Default Page Folder    ", Application.GetOption("Use Default Page Folder")             'Default Database/Project Properties, Use Default Page Folder
    Print #fle, , "XP, 2003", "Default Page Folder        ", Application.GetOption("Default Page Folder")                 'Default Database/Project Properties, Default Page Folder
    Print #fle, , "XP, 2003", "Use Default Connection File", Application.GetOption("Use Default Connection File")         'Default Database/Project Properties, Use Default Connection File
    Print #fle, , "XP, 2003", "Default Connection File    ", Application.GetOption("Default Connection File")             'Default Database/Project Properties, Default Connection File
    Print #fle, ">>>Spelling Tab"
    Print #fle, , "XP, 2003", "Spelling dictionary language               ", Application.GetOption("Spelling dictionary language")                 'Dictionary Language
    Print #fle, , "XP, 2003", "Spelling add words to                      ", Application.GetOption("Spelling add words to")                        'Add words to
    Print #fle, , "XP, 2003", "Spelling suggest from main dictionary only ", Application.GetOption("Spelling suggest from main dictionary only")   'Suggest from main dictionary only
    Print #fle, , "XP, 2003", "Spelling ignore words in UPPERCASE         ", Application.GetOption("Spelling ignore words in UPPERCASE")           'Ignore words in UPPERCASE
    Print #fle, , "XP, 2003", "Spelling ignore words with number          ", Application.GetOption("Spelling ignore words with number")            'Ignore words with numbers
    Print #fle, , "XP, 2003", "Spelling ignore Internet and file addresses", Application.GetOption("Spelling ignore Internet and file addresses")  'Ignore Internet and file addresses
    Print #fle, , "XP, 2003", "Spelling use German post-reform rules      ", Application.GetOption("Spelling use German post-reform rules")        'Language-specific, German: Use post-reform rules
    Print #fle, , "XP, 2003", "Spelling combine aux verb/adj              ", Application.GetOption("Spelling combine aux verb/adj")                'Language-specific, Korean: Combine aux verb/adj.
    Print #fle, , "XP, 2003", "Spelling use auto-change list              ", Application.GetOption("Spelling use auto-change list")                'Language-specific, Korean: Use auto-change list
    Print #fle, , "XP, 2003", "Spelling process compound nouns            ", Application.GetOption("Spelling process compound nouns")              'Language-specific, Korean: Process compound nouns
    Print #fle, , "XP, 2003", "Spelling Hebrew modes                      ", Application.GetOption("Spelling Hebrew modes")                        'Language-specific, Hebrew modes
    Print #fle, , "XP, 2003", "Spelling Arabic modes                      ", Application.GetOption("Spelling Arabic modes")                        'Language-specific, Arabic modes
    Print #fle, ">>>International Tab"
    Print #fle, , "    2003", "Default direction ", Application.GetOption("Default direction")       'Right-to-Left, Default direction
    Print #fle, , "    2003", "General alignment ", Application.GetOption("General alignment")       'Right-to-Left, General alignment
    Print #fle, , "    2003", "Cursor movement   ", Application.GetOption("Cursor movement")         'Right-to-Left, Cursor movement
    Print #fle, , "    2003", "Use Hijri Calendar", Application.GetOption("Use Hijri Calendar")      'Use Hijri Calendar
    Print #fle, ">>>Error Checking Tab"
    Print #fle, , "    2003", "Enable Error Checking                        ", Application.GetOption("Enable Error Checking")                          'Settings, Enable error checking
    Print #fle, , "    2003", "Error Checking Indicator Color               ", Application.GetOption("Error Checking Indicator Color")                 'Settings, Error indicator color
    Print #fle, , "    2003", "Unassociated Label and Control Error Checking", Application.GetOption("Unassociated Label and Control Error Checking")  'Form/Report Design Rules, Unassociated label and control
    Print #fle, , "    2003", "Keyboard Shortcut Errors Error Checking      ", Application.GetOption("Keyboard Shortcut Errors Error Checking")        'Form/Report Design Rules, Keyboard shortcut errors
    Print #fle, , "    2003", "Invalid Control Properties Error Checking    ", Application.GetOption("Invalid Control Properties Error Checking")      'Form/Report Design Rules, Invalid control properties
    Print #fle, , "    2003", "Common Report Errors Error Checking          ", Application.GetOption("Common Report Errors Error Checking")            'Form/Report Design Rules, Common report errors
    Print #fle, ">>>Popular Tab"
    Print #fle, "   >>>Creating databases section"
    Print #fle, , "2007, 2010, 2013", "Default File Format       ", Application.GetOption("Default File Format")            'Default file format
    Print #fle, , "2007, 2010, 2013", "Default Database Directory", Application.GetOption("Default Database Directory")     'Default database folder
    Print #fle, , "2007, 2010, 2013", "New Database Sort Order   ", Application.GetOption("New Database Sort Order")        'New database sort order
    Print #fle, ">>>Current Database Tab"
    Print #fle, "   >>>Application Options section"
    Print #fle, , "2007, 2010, 2013", "Auto Compact                   ", Application.GetOption("Auto Compact")                      'Compact on Close
    Print #fle, , "2007, 2010, 2013", "Remove Personal Information    ", Application.GetOption("Remove Personal Information")       'Remove personal information from file properties on save
    Print #fle, , "2007, 2010, 2013", "Themed Form Controls           ", Application.GetOption("Themed Form Controls")              'Use Windows-themed Controls on Forms
    Print #fle, , "2007, 2010, 2013", "DesignWithData                 ", Application.GetOption("DesignWithData")                    'Enable Layout View for this database
    Print #fle, , "2007, 2010, 2013", "CheckTruncatedNumFields        ", Application.GetOption("CheckTruncatedNumFields")           'Check for truncated number fields
    Print #fle, , "2007, 2010, 2013", "Picture Property Storage Format", Application.GetOption("Picture Property Storage Format")   'Picture Property Storage Format
    Print #fle, "   >>>Name AutoCorrect Options section"
    Print #fle, , "2007, 2010, 2013", "Track Name AutoCorrect Info ", Application.GetOption("Track Name AutoCorrect Info")   'Track name AutoCorrect info
    Print #fle, , "2007, 2010, 2013", "Perform Name AutoCorrect    ", Application.GetOption("Perform Name AutoCorrect")      'Perform name AutoCorrect
    Print #fle, , "2007, 2010, 2013", "Log Name AutoCorrect Changes", Application.GetOption("Log Name AutoCorrect Changes")  'Log name AutoCorrect changes
    Print #fle, "   >>>Filter Lookup options for <Database Name> Database section"
    Print #fle, , "2007, 2010, 2013", "Show Values in Indexed    ", Application.GetOption("Show Values in Indexed")         'Show list of values in, Local indexed fields
    Print #fle, , "2007, 2010, 2013", "Show Values in Non-Indexed", Application.GetOption("Show Values in Non-Indexed")     'Show list of values in, Local nonindexed fields
    Print #fle, , "2007, 2010, 2013", "Show Values in Remote     ", Application.GetOption("Show Values in Remote")          'Show list of values in, ODBC fields
    Print #fle, , "2007, 2010, 2013", "Show Values in Snapshot   ", Application.GetOption("Show Values in Snapshot")        'Show list of values in, Records in local snapshot
    Print #fle, , "2007, 2010, 2013", "Show Values in Server     ", Application.GetOption("Show Values in Server")          'Show list of values in, Records at server
    Print #fle, , "2007, 2010, 2013", "Show Values Limit         ", Application.GetOption("Show Values Limit")              'Don't display lists where more than this number of records read
    Print #fle, ">>>Datasheet Tab"
    Print #fle, "   >>>Default colors section"
    Print #fle, , "2007, 2010, 2013", "Default Font Color      ", Application.GetOption("Default Font Color")               'Font color
    Print #fle, , "2007, 2010, 2013", "Default Background Color", Application.GetOption("Default Background Color")         'Background color
    Print #fle, , "2007, 2010, 2013", "_64                     ", Application.GetOption("_64")                              'Alternate background color
    Print #fle, , "2007, 2010, 2013", "Default Gridlines Color ", Application.GetOption("Default Gridlines Color")          'Gridlines color
    Print #fle, "   >>>Gridlines and cell effects section"
    Print #fle, , "2007, 2010, 2013", "Default Gridlines Horizontal", Application.GetOption("Default Gridlines Horizontal") 'Default gridlines showing, Horizontal
    Print #fle, , "2007, 2010, 2013", "Default Gridlines Vertical  ", Application.GetOption("Default Gridlines Vertical")   'Default gridlines showing, Vertical
    Print #fle, , "2007, 2010, 2013", "Default Cell Effect         ", Application.GetOption("Default Cell Effect")          'Default cell effect
    Print #fle, , "2007, 2010, 2013", "Default Column Width        ", Application.GetOption("Default Column Width")         'Default column width
    Print #fle, "   >>>Default font section"
    Print #fle, , "2007, 2010, 2013", "Default Font Name     ", Application.GetOption("Default Font Name")                  'Font
    Print #fle, , "2007, 2010, 2013", "Default Font Size     ", Application.GetOption("Default Font Size")                  'Size
    Print #fle, , "2007, 2010, 2013", "Default Font Weight   ", Application.GetOption("Default Font Weight")                'Weight
    Print #fle, , "2007, 2010, 2013", "Default Font Underline", Application.GetOption("Default Font Underline")             'Underline
    Print #fle, , "2007, 2010, 2013", "Default Font Italic   ", Application.GetOption("Default Font Italic")                'Italic
    Print #fle, ">>>Object Designers Tab"
    Print #fle, "   >>>Table design section"
    Print #fle, , "2007, 2010, 2013", "Default Text Field Size             ", Application.GetOption("Default Text Field Size")              'Default text field size
    Print #fle, , "2007, 2010, 2013", "Default Number Field Size           ", Application.GetOption("Default Number Field Size")            'Default number field size
    Print #fle, , "2007, 2010, 2013", "Default Field Type                  ", Application.GetOption("Default Field Type")                   'Default field type
    Print #fle, , "2007, 2010, 2013", "AutoIndex on Import/Create          ", Application.GetOption("AutoIndex on Import/Create")           'AutoIndex on Import/Create
    Print #fle, , "2007, 2010, 2013", "Show Property Update Options Buttons", Application.GetOption("Show Property Update Options Buttons") 'Show Property Update Option Buttons
    Print #fle, "   >>>Query design section"
    Print #fle, , "2007, 2010, 2013", "Show Table Names       ", Application.GetOption("Show Table Names")                  'Show table names
    Print #fle, , "2007, 2010, 2013", "Output All Fields      ", Application.GetOption("Output All Fields")                 'Output all fields
    Print #fle, , "2007, 2010, 2013", "Enable AutoJoin        ", Application.GetOption("Enable AutoJoin")                   'Enable AutoJoin
    Print #fle, , "2007, 2010, 2013", "ANSI Query Mode        ", Application.GetOption("ANSI Query Mode")                   'SQL Server Compatible Syntax (ANSI 92), This database
    Print #fle, , "2007, 2010, 2013", "ANSI Query Mode Default", Application.GetOption("ANSI Query Mode Default")           'SQL Server Compatible Syntax (ANSI 92), Default for new databases
    Print #fle, , "2007, 2010, 2013", "Query Design Font Name ", Application.GetOption("Query Design Font Name")            'Query design font, Font
    Print #fle, , "2007, 2010, 2013", "Query Design Font Size ", Application.GetOption("Query Design Font Size")            'Query design font, Size
    Print #fle, "   >>>Forms/Reports section"
    Print #fle, , "2007, 2010, 2013", "Selection Behavior         ", Application.GetOption("Selection Behavior")            'Selection behavior
    Print #fle, , "2007, 2010, 2013", "Form Template              ", Application.GetOption("Form Template")                 'Form template
    Print #fle, , "2007, 2010, 2013", "Report Template            ", Application.GetOption("Report Template")               'Report template
    Print #fle, , "2007, 2010, 2013", "Always Use Event Procedures", Application.GetOption("Always Use Event Procedures")   'Always use event procedures
    Print #fle, "   >>>Error checking section"
    Print #fle, , "2007, 2010, 2013", "Enable Error Checking                        ", Application.GetOption("Enable Error Checking")                           'Enable error checking
    Print #fle, , "2007, 2010, 2013", "Error Checking Indicator Color               ", Application.GetOption("Error Checking Indicator Color")                  'Error indicator color
    Print #fle, , "2007, 2010, 2013", "Unassociated Label and Control Error Checking", Application.GetOption("Unassociated Label and Control Error Checking")   'Check for unassociated label and control
    Print #fle, , "2007, 2010, 2013", "New Unassociated Labels Error Checking       ", Application.GetOption("New Unassociated Labels Error Checking")          'Check for new unassociated labels
    Print #fle, , "2007, 2010, 2013", "Keyboard Shortcut Errors Error Checking      ", Application.GetOption("Keyboard Shortcut Errors Error Checking")         'Check for keyboard shortcut errors
    Print #fle, , "2007, 2010, 2013", "Invalid Control Properties Error Checking    ", Application.GetOption("Invalid Control Properties Error Checking")       'Check for invalid control properties
    Print #fle, , "2007, 2010, 2013", "Common Report Errors Error Checking          ", Application.GetOption("Common Report Errors Error Checking")             'Check for common report errors
    Print #fle, ">>>Proofing Tab"
    Print #fle, "   >>>When correcting spelling in Microsoft Office programs section"
    Print #fle, , "2007, 2010, 2013", "Spelling ignore words in UPPERCASE         ", Application.GetOption("Spelling ignore words in UPPERCASE")            'Ignore words in UPPERCASE
    Print #fle, , "2007, 2010, 2013", "Spelling ignore words with number          ", Application.GetOption("Spelling ignore words with number")             'Ignore words that contain numbers
    Print #fle, , "2007, 2010, 2013", "Spelling ignore Internet and file addresses", Application.GetOption("Spelling ignore Internet and file addresses")   'Ignore Internet and file addresses
    Print #fle, , "2007, 2010, 2013", "Spelling suggest from main dictionary only ", Application.GetOption("Spelling suggest from main dictionary only")    'Suggest from main dictionary only
    Print #fle, , "2007, 2010, 2013", "Spelling dictionary language               ", Application.GetOption("Spelling dictionary language")                  'Dictionary Language
    Print #fle, ">>>Advanced Tab"
    Print #fle, "   >>>Editing section"
    Print #fle, , "2007, 2010, 2013", "Move After Enter                ", Application.GetOption("Move After Enter")                     'Move after enter
    Print #fle, , "2007, 2010, 2013", "Behavior Entering Field         ", Application.GetOption("Behavior Entering Field")              'Behavior entering field
    Print #fle, , "2007, 2010, 2013", "Arrow Key Behavior              ", Application.GetOption("Arrow Key Behavior")                   'Arrow key behavior
    Print #fle, , "2007, 2010, 2013", "Cursor Stops at First/Last Field", Application.GetOption("Cursor Stops at First/Last Field")     'Cursor stops at first/last field
    Print #fle, , "2007, 2010, 2013", "Default Find/Replace Behavior   ", Application.GetOption("Default Find/Replace Behavior")        'Default find/replace behavior
    Print #fle, , "2007, 2010, 2013", "Confirm Record Changes          ", Application.GetOption("Confirm Record Changes")               'Confirm, Record changes
    Print #fle, , "2007, 2010, 2013", "Confirm Document Deletions      ", Application.GetOption("Confirm Document Deletions")           'Confirm, Document deletions
    Print #fle, , "2007, 2010, 2013", "Confirm Action Queries          ", Application.GetOption("Confirm Action Queries")               'Confirm, Action queries
    Print #fle, , "2007, 2010, 2013", "Default Direction               ", Application.GetOption("Default Direction")                    'Default direction
    Print #fle, , "2007, 2010, 2013", "General Alignment               ", Application.GetOption("General Alignment")                    'General alignment
    Print #fle, , "2007, 2010, 2013", "Cursor Movement                 ", Application.GetOption("Cursor Movement")                      'Cursor movement
    Print #fle, , "2007, 2010, 2013", "Datasheet Ime Control           ", Application.GetOption("Datasheet Ime Control")                'Datasheet IME control
    Print #fle, , "2007, 2010, 2013", "Use Hijri Calendar              ", Application.GetOption("Use Hijri Calendar")                   'Use Hijri Calendar
    Print #fle, "   >>>Display section"
    Print #fle, , "2007, 2010, 2013", "Size of MRU File List               ", Application.GetOption("Size of MRU File List")                'Show this number of Recent Documents
    Print #fle, , "2007, 2010, 2013", "Show Status Bar                     ", Application.GetOption("Show Status Bar")                      'Status bar
    Print #fle, , "2007, 2010, 2013", "Show Animations                     ", Application.GetOption("Show Animations")                      'Show animations
    Print #fle, , "2007, 2010, 2013", "Show Smart Tags on Datasheets       ", Application.GetOption("Show Smart Tags on Datasheets")        'Show Smart Tags on Datasheets
    Print #fle, , "2007, 2010, 2013", "Show Smart Tags on Forms and Reports", Application.GetOption("Show Smart Tags on Forms and Reports") 'Show Smart Tags on Forms and Reports
    Print #fle, , "2007, 2010, 2013", "Show Macro Names Column             ", Application.GetOption("Show Macro Names Column")              'Show in Macro Design, Names column
    Print #fle, , "2007, 2010, 2013", "Show Conditions Column              ", Application.GetOption("Show Conditions Column")               'Show in Macro Design, Conditions column
    Print #fle, "   >>>Printing section"
    Print #fle, , "2007, 2010, 2013", "Left Margin  ", Application.GetOption("Left Margin")         'Left margin
    Print #fle, , "2007, 2010, 2013", "Right Margin ", Application.GetOption("Right Margin")        'Right margin
    Print #fle, , "2007, 2010, 2013", "Top Margin   ", Application.GetOption("Top Margin")          'Top margin
    Print #fle, , "2007, 2010, 2013", "Bottom Margin", Application.GetOption("Bottom Margin")       'Bottom margin
    Print #fle, "   >>>General section"
    Print #fle, , "2007, 2010, 2013", "Provide Feedback with Sound             ", Application.GetOption("Provide Feedback with Sound")                  'Provide feedback with sound
    Print #fle, , "2007, 2010, 2013", "Four-Digit Year Formatting              ", Application.GetOption("Four-Digit Year Formatting")                   'Use four-year digit year formatting, This database
    Print #fle, , "2007, 2010, 2013", "Four-Digit Year Formatting All Databases", Application.GetOption("Four-Digit Year Formatting All Databases")     'Use four-year digit year formatting, All databases
    Print #fle, "   >>>Advanced section"
    Print #fle, , "2007, 2010, 2013", "Open Last Used Database When Access Starts", Application.GetOption("Open Last Used Database When Access Starts")     'Open last used database when Access starts
    Print #fle, , "2007, 2010, 2013", "Default Open Mode for Databases           ", Application.GetOption("Default Open Mode for Databases")                'Default open mode
    Print #fle, , "2007, 2010, 2013", "Default Record Locking                    ", Application.GetOption("Default Record Locking")                         'Default record locking
    Print #fle, , "2007, 2010, 2013", "Use Row Level Locking                     ", Application.GetOption("Use Row Level Locking")                          'Open databases by using record-level locking
    Print #fle, , "2007, 2010, 2013", "OLE/DDE Timeout (sec)                     ", Application.GetOption("OLE/DDE Timeout (sec)")                          'OLE/DDE timeout (sec)
    Print #fle, , "2007, 2010, 2013", "Refresh Interval (sec)                    ", Application.GetOption("Refresh Interval (sec)")                         'Refresh interval (sec)
    Print #fle, , "2007, 2010, 2013", "Number of Update Retries                  ", Application.GetOption("Number of Update Retries")                       'Number of update retries
    Print #fle, , "2007, 2010, 2013", "ODBC Refresh Interval (sec)               ", Application.GetOption("ODBC Refresh Interval (sec)")                    'ODBC refresh interval (sec)
    Print #fle, , "2007, 2010, 2013", "Update Retry Interval (msec)              ", Application.GetOption("Update Retry Interval (msec)")                   'Update retry interval (msec)
    Print #fle, , "2007, 2010, 2013", "Ignore DDE Requests                       ", Application.GetOption("Ignore DDE Requests")                            'DDE operations, Ignore DDE requests
    Print #fle, , "2007, 2010, 2013", "Enable DDE Refresh                        ", Application.GetOption("Enable DDE Refresh")                             'DDE operations, Enable DDE refresh
    Print #fle, , "2007, 2010, 2013", "Command-Line Arguments                    ", Application.GetOption("Command-Line Arguments")                         'Command-line arguments

PROC_EXIT:
    Set dbs = Nothing
    Close fle
    PopCallStack
    Exit Sub

PROC_ERR:
    If Err = 2091 Then          ''...' is an invalid name.
        If Debugit Then Debug.Print "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListOfAccessApplicationOptions of Class aegitClass"
        Print #fle, "!" & Err.Description
        Err.Clear
    ElseIf Err = 3270 Then      'Property not found.
        Err.Clear
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListOfAccessApplicationOptions of Class aegitClass"
        GlobalErrHandler
    End If
    Resume Next

End Sub

Private Sub ListOfApplicationProperties()
' Ref: http://www.granite.ab.ca/access/settingstartupoptions.htm

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "ListOfApplicationProperties"

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim fle As Integer

    fle = FreeFile()
    Open aegitSourceFolder & "\ListOfApplicationProperties.txt" For Output As #fle

    Dim prp As DAO.Property
    Dim i As Integer
    Dim strPropName As String
    Dim varPropValue As Variant
    Dim varPropType As Variant
    Dim varPropInherited As Variant
    Dim intPropPropCount As Integer
    Dim strError As String

    With dbs
        For i = 0 To (.Properties.Count - 1)
            strPropName = .Properties(i).Name
            varPropValue = Null
            varPropValue = .Properties(i).Value
            varPropType = .Properties(i).Type
            varPropInherited = .Properties(i).Inherited
            Print #fle, strPropName & ": " & varPropValue & ", " & _
                varPropType & ", " & varPropInherited
        Next i
    End With

PROC_EXIT:
    Set dbs = Nothing
    Close fle
    PopCallStack
    Exit Sub

PROC_ERR:
    If Err = 3251 Then
        Debug.Print "Erl=" & Erl & " Error " & Err.Number & " strPropName=" & strPropName & " (" & Err.Description & ") in procedure ListOfApplicationProperties of Class aegitClass"
        Print #fle, "!" & Err.Description, strPropName
        Err.Clear
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListOfApplicationProperties of Class aegitClass"
        GlobalErrHandler
    End If
    Resume Next

End Sub

Private Function Pause(NumberOfSeconds As Variant)
' Ref: http://www.access-programmers.co.uk/forums/showthread.php?p=952355

    On Error GoTo PROC_ERR

    Dim PauseTime As Variant
    Dim Start As Variant

    PauseTime = NumberOfSeconds
    Start = Timer
    Do While Timer < Start + PauseTime
        Sleep 100
        DoEvents
    Loop

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Pause of Class aegitClass"
    Resume PROC_EXIT

End Function

Private Sub WaitSeconds(intSeconds As Integer)
' Comments: Waits for a specified number of seconds
' Params  : intSeconds      Number of seconds to wait
' Source  : Total Visual SourceBook
' Ref     : http://www.fmsinc.com/MicrosoftAccess/modules/examples/AvoidDoEvents.asp

    On Error GoTo PROC_ERR

    Dim datTime As Date

    datTime = DateAdd("s", intSeconds, Now)

    Do
        ' Yield to other programs (better than using DoEvents which eats up all the CPU cycles)
        Sleep 100
        DoEvents
    Loop Until Now >= datTime

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure WaitSeconds of Class aegitClass"
    Resume PROC_EXIT
End Sub

Private Function aeGetReferences(Optional varDebug As Variant) As Boolean
' Ref: http://vbadud.blogspot.com/2008/04/get-references-of-vba-project.html
' Ref: http://www.pcreview.co.uk/forums/type-property-reference-object-vbulletin-project-t3793816.html
' Ref: http://www.cpearson.com/excel/missingreferences.aspx
' Ref: http://allenbrowne.com/ser-38.html
' Ref: http://access.mvps.org/access/modules/mdl0022.htm (References Wizard)
' Ref: http://www.accessmvp.com/djsteele/AccessReferenceErrors.html
'====================================================================
' Author:   Peter F. Ennis
' Date:     November 28, 2012
' Comment:  Added and adapted from aeladdin (tm) code
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim i As Integer
    Dim RefName As String
    Dim RefDesc As String
    Dim blnRefBroken As Boolean
    Dim blnDebug As Boolean
    Dim strFile As String

    Dim vbaProj As Object
    Set vbaProj = Application.VBE.ActiveVBProject

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeGetReferences"

    Debug.Print "aeGetReferences"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeGetReferences is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeGetReferences is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    strFile = aestrSourceLocation & aeRefTxtFile
    
    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    If blnDebug Then
        Debug.Print ">==> aeGetReferences >==>"
        Debug.Print , "vbaProj.Name = " & vbaProj.Name
        Debug.Print , "vbaProj.Type = '" & vbaProj.Type & "'"
        ' Display the versions of Access, ADO and DAO
        Debug.Print , "Access version = " & Application.Version
        Debug.Print , "ADO (ActiveX Data Object) version = " & CurrentProject.Connection.Version
        Debug.Print , "DAO (DbEngine)  version = " & Application.DBEngine.Version
        Debug.Print , "DAO (CodeDb)    version = " & Application.CodeDb.Version
        Debug.Print , "DAO (CurrentDb) version = " & Application.CurrentDb.Version
        Debug.Print , "<@_@>"
        Debug.Print , "     " & "References:"
    End If

        Print #1, ">==> The Project References >==>"
        Print #1, , "vbaProj.Name = " & vbaProj.Name
        Print #1, , "vbaProj.Type = '" & vbaProj.Type & "'"
        ' Display the versions of Access, ADO and DAO
        Print #1, , "Access version = " & Application.Version
        Print #1, , "ADO (ActiveX Data Object) version = " & CurrentProject.Connection.Version
        Print #1, , "DAO (DbEngine)  version = " & Application.DBEngine.Version
        Print #1, , "DAO (CodeDb)    version = " & Application.CodeDb.Version
        Print #1, , "DAO (CurrentDb) version = " & Application.CurrentDb.Version
        Print #1, , "<@_@>"
        Print #1, , "     " & "References:"

    For i = 1 To vbaProj.References.Count

        blnRefBroken = False

        ' Get the Name of the Reference
        RefName = vbaProj.References(i).Name

        ' Get the Description of Reference
        RefDesc = vbaProj.References(i).Description

        If blnDebug Then Debug.Print , , vbaProj.References(i).Name, vbaProj.References(i).Description
        If blnDebug Then Debug.Print , , , vbaProj.References(i).FullPath
        If blnDebug Then Debug.Print , , , vbaProj.References(i).Guid

        Print #1, , , vbaProj.References(i).Name, vbaProj.References(i).Description
        Print #1, , , , vbaProj.References(i).FullPath
        Print #1, , , , vbaProj.References(i).Guid

        ' Returns a Boolean value indicating whether or not the Reference object points to a valid reference in the registry. Read-only.
        If Application.VBE.ActiveVBProject.References(i).IsBroken = True Then
              blnRefBroken = True
              If blnDebug Then Debug.Print , , vbaProj.References(i).Name, "blnRefBroken=" & blnRefBroken
              Print #1, , , vbaProj.References(i).Name, "blnRefBroken=" & blnRefBroken
        End If
    Next
    If blnDebug Then Debug.Print , "<*_*>"
    If blnDebug Then Debug.Print "<==<"

    Print #1, , "<*_*>"
    Print #1, "<==<"

    aeGetReferences = True

PROC_EXIT:
    Set vbaProj = Nothing
    Close 1
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeGetReferences of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeGetReferences of Class aegitClass"
    aeGetReferences = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function LongestTableName() As Integer
'====================================================================
' Author:   Peter F. Ennis
' Date:     November 30, 2012
' Comment:  Return the length of the longest table name
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim intTNLen As Integer

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "LongestTableName"

    intTNLen = 0
    Set dbs = CurrentDb()
    For Each tdf In CurrentDb.TableDefs
        If Not (Left(tdf.Name, 4) = "MSys" _
                Or Left(tdf.Name, 4) = "~TMP" _
                Or Left(tdf.Name, 3) = "zzz") Then
            If Len(tdf.Name) > intTNLen Then
                intTNLen = Len(tdf.Name)
            End If
        End If
    Next tdf

    LongestTableName = intTNLen

PROC_EXIT:
    Set tdf = Nothing
    Set dbs = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestTableName of Class aegitClass"
    LongestTableName = 0
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function LongestFieldPropsName() As Boolean
'====================================================================
' Author:   Peter F. Ennis
' Date:     December 5, 2012
' Comment:  Return length of field properties for text output alignment
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim dbs As DAO.Database
    Dim tblDef As DAO.TableDef
    Dim fld As DAO.Field

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "LongestFieldPropsName"

    On Error GoTo PROC_ERR

    aeintFNLen = 0
    aeintFTLen = 0
    aeintFDLen = 0

    Set dbs = CurrentDb()

    For Each tblDef In CurrentDb.TableDefs
        If Not (Left(tblDef.Name, 4) = "MSys" _
                Or Left(tblDef.Name, 4) = "~TMP" _
                Or Left(tblDef.Name, 3) = "zzz") Then
            For Each fld In tblDef.Fields
                If Len(fld.Name) > aeintFNLen Then
                    aestrLFNTN = tblDef.Name
                    aestrLFN = fld.Name
                    aeintFNLen = Len(fld.Name)
                End If
                If Len(FieldTypeName(fld)) > aeintFTLen Then
                    aestrLFT = FieldTypeName(fld)
                    aeintFTLen = Len(FieldTypeName(fld))
                End If
                If Len(GetDescrip(fld)) > aeintFDLen Then
                    aestrLFD = GetDescrip(fld)
                    aeintFDLen = Len(GetDescrip(fld))
                End If
            Next
        End If
    Next tblDef

    LongestFieldPropsName = True

PROC_EXIT:
    Set fld = Nothing
    Set tblDef = Nothing
    Set dbs = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestFieldPropsName of Class aegitClass"
    LongestFieldPropsName = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function SizeString(Text As String, Length As Long, _
    Optional ByVal TextSide As SizeStringSide = TextLeft, _
    Optional PadChar As String = " ") As String
' Ref: http://www.cpearson.com/excel/sizestring.htm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SizeString
' This procedure creates a string of a specified length. Text is the original string
' to include, and Length is the length of the result string. TextSide indicates whether
' Text should appear on the left (in which case the result is padded on the right with
' PadChar) or on the right (in which case the string is padded on the left). When padding on
' either the left or right, padding is done using the PadChar. character. If PadChar is omitted,
' a space is used. If PadChar is longer than one character, the left-most character of PadChar
' is used. If PadChar is an empty string, a space is used. If TextSide is neither
' TextLeft or TextRight, the procedure uses TextLeft.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim sPadChar As String

    If Len(Text) >= Length Then
        ' if the source string is longer than the specified length, return the
        ' Length left characters
        SizeString = Left(Text, Length)
        Exit Function
    End If

    If Len(PadChar) = 0 Then
        ' PadChar is an empty string. use a space.
        sPadChar = " "
    Else
        ' use only the first character of PadChar
        sPadChar = Left(PadChar, 1)
    End If

    If (TextSide <> TextLeft) And (TextSide <> TextRight) Then
        ' if TextSide was neither TextLeft nor TextRight, use TextLeft.
        TextSide = TextLeft
    End If

    If TextSide = TextLeft Then
        ' if the text goes on the left, fill out the right with spaces
        SizeString = Text & String(Length - Len(Text), sPadChar)
    Else
        ' otherwise fill on the left and put the Text on the right
        SizeString = String(Length - Len(Text), sPadChar) & Text
    End If

End Function

Private Function GetLinkedTableCurrentPath(MyLinkedTable As String) As String
' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=198057
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
        ' Non-linked table returns blank - Instr removes the "Database="

PROC_EXIT:
    On Error Resume Next
    Exit Function

PROC_ERR:
    Select Case Err.Number
        ' Case ###         ' Add your own error management or log error to logging table
        Case Else
            ' Add your own custom log usage function
    End Select
    Resume PROC_EXIT
End Function

Private Function FileLocked(strFileName As String) As Boolean
' Ref: http://support.microsoft.com/kb/209189
    On Error Resume Next
    ' If the file is already opened by another process,
    ' and the specified type of access is not allowed,
    ' the Open operation fails and an error occurs.
    Open strFileName For Binary Access Read Write Lock Read Write As #1
    Close 1
    ' If an error occurs, the document is currently open.
    If Err.Number <> 0 Then
        ' Display the error number and description.
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FileLocked of Class aegitClass"
        FileLocked = True
        Err.Clear
    End If
End Function

Private Function TableInfo(strTableName As String, Optional varDebug As Variant) As Boolean
' Ref: http://allenbrowne.com/func-06.html
'====================================================================
' Purpose:  Display the field names, types, sizes and descriptions for a table
' Argument: Name of a table in the current database
' Updates:  Peter F. Ennis
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim sLen As Long
    Dim strLinkedTablePath As String
    
    Dim blnDebug As Boolean

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "TableInfo"

    On Error GoTo PROC_ERR

    strLinkedTablePath = ""

    If IsMissing(varDebug) Then
        blnDebug = False
        'Debug.Print , "varDebug IS missing so blnDebug of TableInfo is set to False"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        'Debug.Print , "varDebug IS NOT missing so blnDebug of TableInfo is set to True"
        'Debug.Print , "NOW DEBUGGING..."
    End If

    Set dbs = CurrentDb()
    Set tdf = dbs.TableDefs(strTableName)
    sLen = Len("TABLE: ") + Len(strTableName)

    strLinkedTablePath = GetLinkedTableCurrentPath(strTableName)

    aeintFDLen = LongestTableDescription(tdf.Name)

    If aeintFDLen < Len("DESCRIPTION") Then aeintFDLen = Len("DESCRIPTION")

    If blnDebug Then
    'If blnDebug And aeintFDLen <> 11 Then
        Debug.Print SizeString("-", sLen, TextLeft, "-")
        Debug.Print SizeString("TABLE: " & strTableName, sLen, TextLeft, " ")
        Debug.Print SizeString("-", sLen, TextLeft, "-")
        If strLinkedTablePath <> "" Then
            Debug.Print strLinkedTablePath
        End If
        Debug.Print SizeString("FIELD NAME", aeintFNLen, TextLeft, " ") _
                        & aestr4 & SizeString("FIELD TYPE", aeintFTLen, TextLeft, " ") _
                        & aestr4 & SizeString("SIZE", aeintFSize, TextLeft, " ") _
                        & aestr4 & SizeString("DESCRIPTION", aeintFDLen, TextLeft, " ")
        Debug.Print SizeString("=", aeintFNLen, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFTLen, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFSize, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFDLen, TextLeft, "=")
    End If
  
    'Debug.Print ">>>", SizeString("-", sLen, TextLeft, "-")
    Print #1, SizeString("-", sLen, TextLeft, "-")
    Print #1, SizeString("TABLE: " & strTableName, sLen, TextLeft, " ")
    Print #1, SizeString("-", sLen, TextLeft, "-")
    If strLinkedTablePath <> "" Then
        Print #1, "Linked=>" & strLinkedTablePath
    End If
    Print #1, SizeString("FIELD NAME", aeintFNLen, TextLeft, " ") _
                        & aestr4 & SizeString("FIELD TYPE", aeintFTLen, TextLeft, " ") _
                        & aestr4 & SizeString("SIZE", aeintFSize, TextLeft, " ") _
                        & aestr4 & SizeString("DESCRIPTION", aeintFDLen, TextLeft, " ")
    Print #1, SizeString("=", aeintFNLen, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFTLen, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFSize, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFDLen, TextLeft, "=")
    strLinkedTablePath = ""

    For Each fld In tdf.Fields
        If blnDebug Then
        'If blnDebug And aeintFDLen <> 11 Then
            Debug.Print SizeString(fld.Name, aeintFNLen, TextLeft, " ") _
                & aestr4 & SizeString(FieldTypeName(fld), aeintFTLen, TextLeft, " ") _
                & aestr4 & SizeString(fld.size, aeintFSize, TextLeft, " ") _
                & aestr4 & SizeString(GetDescrip(fld), aeintFDLen, TextLeft, " ")
        End If
        Print #1, SizeString(fld.Name, aeintFNLen, TextLeft, " ") _
            & aestr4 & SizeString(FieldTypeName(fld), aeintFTLen, TextLeft, " ") _
            & aestr4 & SizeString(fld.size, aeintFSize, TextLeft, " ") _
            & aestr4 & SizeString(GetDescrip(fld), aeintFDLen, TextLeft, " ")
    Next
    If blnDebug Then Debug.Print
    'If blnDebug And aeintFDLen <> 11 Then Debug.Print
    Print #1, vbCrLf

    TableInfo = True

PROC_EXIT:
    Set fld = Nothing
    Set tdf = Nothing
    Set dbs = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TableInfo of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TableInfo of Class aegitClass"
    TableInfo = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function GetDescrip(obj As Object) As String
    On Error Resume Next
    GetDescrip = obj.Properties("Description")
End Function

Private Function LongestTableDescription(strTblName As String) As Integer
' ?LongestTableDescription("tblCaseManager")

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim strLFD As String

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "LongestTableDescription"

    Set dbs = CurrentDb()
    Set tdf = dbs.TableDefs(strTblName)

    For Each fld In tdf.Fields
        If Len(GetDescrip(fld)) > aeintFDLen Then
            strLFD = GetDescrip(fld)
            aeintFDLen = Len(GetDescrip(fld))
        End If
    Next

    LongestTableDescription = aeintFDLen

PROC_EXIT:
    Set fld = Nothing
    Set tdf = Nothing
    Set dbs = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestTableDescription of Class aegitClass"
    LongestTableDescription = -1
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function FieldTypeName(fld As DAO.Field) As String
' Ref: http://allenbrowne.com/func-06.html
' Purpose: Converts the numeric results of DAO Field.Type to text
    Dim strReturn As String    ' Name to return

    Select Case CLng(fld.Type) ' fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            '  1
        Case dbByte: strReturn = "Byte"                 '  2
        Case dbInteger: strReturn = "Integer"           '  3
        Case dbLong                                     '  4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         '  5
        Case dbSingle: strReturn = "Single"             '  6
        Case dbDouble: strReturn = "Double"             '  7
        Case dbDate: strReturn = "Date/Time"            '  8
        Case dbBinary: strReturn = "Binary"             '  9 (no interface)
        Case dbText                                     ' 10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        ' (no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     ' 11
        Case dbMemo                                     ' 12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 ' 15

        ' Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        ' 16
        Case dbVarBinary: strReturn = "VarBinary"       ' 17
        Case dbChar: strReturn = "Char"                 ' 18
        Case dbNumeric: strReturn = "Numeric"           ' 19
        Case dbDecimal: strReturn = "Decimal"           ' 20
        Case dbFloat: strReturn = "Float"               ' 21
        Case dbTime: strReturn = "Time"                 ' 22
        Case dbTimeStamp: strReturn = "Time Stamp"      ' 23

        ' Constants for complex types don't work
        ' prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"             ' dbAttachment
        Case 102&: strReturn = "Complex Byte"           ' dbComplexByte
        Case 103&: strReturn = "Complex Integer"        ' dbComplexInteger
        Case 104&: strReturn = "Complex Long"           ' dbComplexLong
        Case 105&: strReturn = "Complex Single"         ' dbComplexSingle
        Case 106&: strReturn = "Complex Double"         ' dbComplexDouble
        Case 107&: strReturn = "Complex GUID"           ' dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"        ' dbComplexDecimal
        Case 109&: strReturn = "Complex Text"           ' dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn

End Function

Private Function aeDocumentTables(Optional varDebug As Variant) As Boolean
' Ref: http://www.tek-tips.com/faqs.cfm?fid=6905
' Document the tables, fields, and relationships
' Tables, field type, primary keys, foreign keys, indexes
' Relationships in the database with table, foreign table, primary keys, foreign keys
' Ref: http://allenbrowne.com/func-06.html

    'Dim strDoc As String
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim blnDebug As Boolean
    Dim blnResult As Boolean
    Dim intFailCount As Integer
    Dim strFile As String

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeDocumentTables"

    On Error GoTo PROC_ERR

    intFailCount = 0
    
    LongestFieldPropsName
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Name=" & aestrLFN
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Name Length=" & aeintFNLen
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Name Table Name=" & aestrLFNTN
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Description=" & aestrLFD
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Description Length=" & aeintFDLen
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Type=" & aestrLFT
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Type Length=" & aeintFTLen

    ' Reset values
    aestrLFN = ""
    If aeintFNLen < 11 Then aeintFNLen = 11     ' Minimum required by design
    'aestrLFNTN = ""
    'aestrLFD = ""
    aeintFDLen = 0
    'aestrLFT = ""
    'aeintFTLen = 0

    Debug.Print "aeDocumentTables"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeDocumentTables is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeDocumentTables is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    strFile = aestrSourceLocation & aeTblTxtFile
    
    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    For Each tdf In CurrentDb.TableDefs
        If Not (Left(tdf.Name, 4) = "MSys" _
                Or Left(tdf.Name, 4) = "~TMP" _
                Or Left(tdf.Name, 3) = "zzz") Then
            If blnDebug Then
                blnResult = TableInfo(tdf.Name, "WithDebugging")
                If Not blnResult Then intFailCount = intFailCount + 1
                If blnDebug And aeintFDLen <> 11 Then Debug.Print "aeintFDLen=" & aeintFDLen
            Else
                blnResult = TableInfo(tdf.Name)
                If Not blnResult Then intFailCount = intFailCount + 1
            End If
            'Debug.Print
            aeintFDLen = 0
        End If
    Next tdf

    'If intFailCount > 0 Then
    '    aeDocumentTables = False
    'Else
    '    aeDocumentTables = True
    'End If
    If blnDebug Then
        Debug.Print "intFailCount = " & intFailCount
        'Debug.Print "aeDocumentTables = " & aeDocumentTables
    End If

    aeDocumentTables = True

PROC_EXIT:
    Set fld = Nothing
    Set tdf = Nothing
    Close 1
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTables of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTables of Class aegitClass"
    aeDocumentTables = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function aeDocumentTablesXML(Optional varDebug As Variant) As Boolean
' Ref: http://stackoverflow.com/questions/4867727/how-to-use-ms-access-saveastext-with-queries-specifically-stored-procedures

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeDocumentTablesXML"
    
    Dim dbs As DAO.Database
    Dim tbl As DAO.TableDef
    Dim strObjName As String

    Set dbs = CurrentDb

    Dim blnDebug As Boolean
    Dim intFailCount As Integer

    If aegitXMLFolder = "default" Then
        aestrXMLLocation = aegitType.XMLFolder
    Else
        aestrXMLLocation = aegitXMLFolder
    End If

    If Not FolderExists(aestrXMLLocation) Then
        MsgBox aestrXMLLocation & " does not exist!", vbCritical, "Error"
        Stop
    End If

    'MsgBox "aeDocumentTablesXML: LBound(aegitDataXML())=" & LBound(aegitDataXML()) & _
        vbCrLf & "UBound(aegitDataXML())=" & UBound(aegitDataXML()), vbInformation, "CHECK"
    If aegitExportDataToXML Then OutputTheTableDataAsXML aegitDataXML()

    intFailCount = 0
    Debug.Print "aeDocumentTablesXML"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeDocumentTablesXML is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeDocumentTablesXML is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    If blnDebug Then Debug.Print ">List of tables exported as XML to " & aestrXMLLocation
    For Each tbl In dbs.TableDefs
        If tbl.Attributes = 0 Then      ' Ignore System Tables
            strObjName = tbl.Name
            If blnDebug Then Debug.Print , "- " & strObjName & ".xsd"
            'Debug.Print "aestrXMLLocation=" & aestrXMLLocation
            'Debug.Print "the XML file=" & aestrXMLLocation & strObjName
            Application.ExportXML acExportTable, strObjName, , _
                        aestrXMLLocation & "tables_" & strObjName & ".xsd"
        End If
    Next

    If intFailCount > 0 Then
        aeDocumentTablesXML = False
    Else
        aeDocumentTablesXML = True
    End If

    If blnDebug Then
        Debug.Print "intFailCount = " & intFailCount
        Debug.Print "aeDocumentTablesXML = " & aeDocumentTablesXML
    End If

PROC_EXIT:
    Set tbl = Nothing
    Set dbs = Nothing
    Close 1
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTablesXML of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTablesXML of Class aegitClass"
    aeDocumentTablesXML = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Sub OutputTheSchemaFile()               ' CreateDbScript()
' Remou - Ref: http://stackoverflow.com/questions/698839/how-to-extract-the-schema-of-an-access-mdb-database/9910716#9910716

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim ndx As DAO.Index
    Dim strSQL As String
    Dim strFlds As String
    Dim strCn As String
    Dim strLinkedTablePath As String
    Dim f As Object

    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.CreateTextFile(aegitSourceFolder & aeSchemaFile)

    strSQL = "Public Sub CreateTheDb()" & vbCrLf
    f.WriteLine strSQL
    strSQL = "Dim strSQL As String"
    f.WriteLine strSQL
    strSQL = "On Error GoTo ErrorTrap"
    f.WriteLine strSQL

    For Each tdf In dbs.TableDefs
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
                        strFlds = strFlds & "Text (" & fld.size & ")"
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
    'Debug.Print "Done"

End Sub

Private Function isPK(tdf As DAO.TableDef, strField As String) As Boolean
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tdf.Indexes
        If idx.Primary Then
            For Each fld In idx.Fields
                If strField = fld.Name Then
                    isPK = True
                    Exit Function
                End If
            Next fld
        End If
    Next idx
End Function

Private Function isIndex(tdf As DAO.TableDef, strField As String) As Boolean
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tdf.Indexes
        For Each fld In idx.Fields
            If strField = fld.Name Then
                isIndex = True
                Exit Function
            End If
        Next fld
    Next idx
End Function

Private Function isFK(tdf As DAO.TableDef, strField As String) As Boolean
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tdf.Indexes
        If idx.Foreign Then
            For Each fld In idx.Fields
                If strField = fld.Name Then
                    isFK = True
                    Exit Function
                End If
            Next fld
        End If
    Next idx
End Function

Private Function aeDocumentRelations(Optional varDebug As Variant) As Boolean
' Ref: http://www.tek-tips.com/faqs.cfm?fid=6905
  
    Dim strDocument As String
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim prop As DAO.Property
    Dim strFile As String
    
    Dim blnDebug As Boolean

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeDocumentRelations"

    Debug.Print "aeDocumentRelations"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeDocumentRelations is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeDocumentRelations is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    strFile = aestrSourceLocation & aeRelTxtFile
    
    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    For Each rel In CurrentDb.Relations
        If Not (Left(rel.Name, 4) = "MSys" _
                        Or Left(rel.Name, 4) = "~TMP" _
                        Or Left(rel.Name, 3) = "zzz") Then
            strDocument = strDocument & vbCrLf & "Name: " & rel.Name & vbCrLf
            strDocument = strDocument & "  " & "Table: " & rel.Table & vbCrLf
            strDocument = strDocument & "  " & "Foreign Table: " & rel.ForeignTable & vbCrLf
            For Each fld In rel.Fields
                strDocument = strDocument & "  PK: " & fld.Name & "   FK:" & fld.ForeignName
                strDocument = strDocument & vbCrLf
            Next fld
        End If
    Next rel
    If blnDebug Then Debug.Print strDocument
    Print #1, strDocument
    
    aeDocumentRelations = True

PROC_EXIT:
    Set prop = Nothing
    Set idx = Nothing
    Set fld = Nothing
    Set rel = Nothing
    Close 1
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentRelations of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentRelations of Class aegitClass"
    aeDocumentRelations = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function OutputQueriesSqlText() As Boolean
' Ref: http://www.pcreview.co.uk/forums/export-sql-saved-query-into-text-file-t2775525.html
'====================================================================
' Author:   Peter F. Ennis
' Date:     December 3, 2012
' Comment:  Output the sql code of all queries to a text file
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim dbs As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strFile As String

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "OutputQueriesSqlText"

    strFile = aestrSourceLocation & aeSqlTxtFile
    
    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    Set dbs = CurrentDb
    For Each qdf In dbs.QueryDefs
        If Not (Left(qdf.Name, 4) = "MSys" Or Left(qdf.Name, 4) = "~sq_" _
                        Or Left(qdf.Name, 4) = "~TMP" _
                        Or Left(qdf.Name, 3) = "zzz") Then
            Print #1, "<<<" & qdf.Name & ">>>" & vbCrLf & qdf.SQL
        End If
    Next

    OutputQueriesSqlText = True

PROC_EXIT:
    Set qdf = Nothing
    Set dbs = Nothing
    Close 1
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputQueriesSqlText of Class aegitClass"
    OutputQueriesSqlText = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Sub KillProperly(Killfile As String)
' Ref: http://word.mvps.org/faqs/macrosvba/DeleteFiles.htm

    ' Use a call stack and global error handler
    'If gcfHandleErrors Then On Error GoTo PROC_ERR
    'PushCallStack "KillProperly"

    On Error GoTo PROC_ERR

TryAgain:
    If Len(Dir(Killfile)) > 0 Then
        SetAttr Killfile, vbNormal
        Kill Killfile
    End If

PROC_EXIT:
    'PopCallStack
    Exit Sub

PROC_ERR:
    If Err = 70 Then
        Pause (0.25)
        Resume TryAgain
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " Killfile=" & Killfile & " (" & Err.Description & ") in procedure KillProperly of Class aegitClass"
    'GlobalErrHandler
    Resume PROC_EXIT

End Sub

Private Function GetPropEnum(typeNum As Long) As String
' Ref: http://msdn.microsoft.com/en-us/library/bb242635.aspx
 
    Select Case typeNum
        Case 1
            GetPropEnum = "dbBoolean"
        Case 2
            GetPropEnum = "dbByte"
        Case 3
            GetPropEnum = "dbInteger"
        Case 4
            GetPropEnum = "dbLong"
        Case 5
            GetPropEnum = "dbCurrency"
        Case 6
            GetPropEnum = "dbSingle"
        Case 7
            GetPropEnum = "dbDouble"
        Case 8
            GetPropEnum = "dbDate"
        Case 9
            GetPropEnum = "dbBinary"
        Case 10
            GetPropEnum = "dbText"
        Case 11
            GetPropEnum = "dbLongBinary"
        Case 12
            GetPropEnum = "dbMemo"
        Case 15
            GetPropEnum = "dbGUID"
        Case 16
            GetPropEnum = "dbBigInt"
        Case 17
            GetPropEnum = "dbVarBinary"
        Case 18
            GetPropEnum = "dbChar"
        Case 19
            GetPropEnum = "dbNumeric"
        Case 20
            GetPropEnum = "dbDecimal"
        Case 21
            GetPropEnum = "dbFloat"
        Case 22
            GetPropEnum = "dbTime"
        Case 23
            GetPropEnum = "dbTimeStamp"
        Case 101
            GetPropEnum = "dbAttachment"
        Case 102
            GetPropEnum = "dbComplexByte"
        Case 103
            GetPropEnum = "dbComplexInteger"
        Case 104
            GetPropEnum = "dbComplexLong"
        Case 105
            GetPropEnum = "dbComplexSingle"
        Case 106
            GetPropEnum = "dbComplexDouble"
        Case 107
            GetPropEnum = "dbComplexGUID"
        Case 108
            GetPropEnum = "dbComplexDecimal"
        Case 109
            GetPropEnum = "dbComplexText"
    End Select

End Function

Private Function GetPrpValue(obj As Object) As String
    'On Error Resume Next
    GetPrpValue = obj.Properties("Value")
End Function
 
Private Function OutputBuiltInPropertiesText() As Boolean
' Ref: http://www.jpsoftwaretech.com/listing-built-in-access-database-properties/

    Dim dbs As DAO.Database
    Dim prps As DAO.Properties
    Dim prp As DAO.Property
    Dim varPropValue As Variant
    Dim strFile As String
    Dim strError As String

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "OutputBuiltInPropertiesText"

    strFile = aestrSourceLocation & aePrpTxtFile

    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If
 
    Set dbs = CurrentDb
    Set prps = dbs.Properties

    Debug.Print "OutputBuiltInPropertiesText"

    For Each prp In prps
        Print #1, "Name: " & prp.Name
        Print #1, "Type: " & GetPropEnum(prp.Type)
        ' Fixed for error 3251
        varPropValue = GetPrpValue(prp)
        Print #1, "Value: " & varPropValue
        Print #1, "Inherited: " & prp.Inherited & ";" & strError
        strError = ""
        Print #1, "---"
    Next prp

    OutputBuiltInPropertiesText = True

PROC_EXIT:
    Set prp = Nothing
    Set prps = Nothing
    Set dbs = Nothing
    Close 1
    PopCallStack
    Exit Function

PROC_ERR:
     Select Case Err.Number
        Case 3251
            strError = Err.Number & "," & Err.Description
            varPropValue = Null
            Resume Next
        Case Else
            'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputBuiltInPropertiesText of Class aegitClass"
            'If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & err.Number & " (" & Err.Description & ") in procedure OutputBuiltInPropertiesText of Class aegitClass"
            OutputBuiltInPropertiesText = False
            GlobalErrHandler
            Resume PROC_EXIT
    End Select

End Function
 
Private Function IsFileLocked(PathFileName As String) As Boolean
' Ref: http://accessexperts.com/blog/2012/03/06/checking-if-files-are-locked/

    'Debug.Print "IsFileLocked Entry PathFileName=" & PathFileName

    ' Use a call stack and global error handler
    'If gcfHandleErrors Then On Error GoTo PROC_ERR
    'PushCallStack "IsFileLocked"

    On Error GoTo PROC_ERR

    Dim i As Integer

    'Debug.Print , Len(Dir$(PathFileName))
    If Len(Dir$(PathFileName)) Then
        i = FreeFile()
        Open PathFileName For Random Access Read Write Lock Read Write As #i
        Lock i 'Redundant but let's be 100% sure
        Unlock i
        Close i
    Else
        'Err.Raise 53
    End If

PROC_EXIT:
    On Error GoTo 0
    'PopCallStack
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 70 'Unable to acquire exclusive lock
            MsgBox "A:Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegitClass"
            IsFileLocked = True
        Case 9
            MsgBox "B:Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegitClass" & _
                    vbCrLf & "IsFileLocked Entry PathFileName=" & PathFileName, vbCritical, "ERROR=9"
            IsFileLocked = False
            'If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegitClass"
            'GlobalErrHandler
            Resume PROC_EXIT
        Case Else
            MsgBox "C:Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegitClass"
            'If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegitClass"
            'GlobalErrHandler
            Resume PROC_EXIT
    End Select
    Resume

End Function

Private Function DocumentTheContainer(strContainerType As String, strExt As String, Optional varDebug As Variant) As Boolean
' strContainerType: Forms, Reports, Scripts (Macros), Modules

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "DocumentTheContainer"

    Dim dbs As DAO.Database
    Dim cnt As DAO.Container
    Dim doc As DAO.Document
    Dim i As Integer
    Dim intAcObjType As Integer
    Dim blnDebug As Boolean
    Dim strTheCurrentPathAndFile As String

    Set dbs = CurrentDb() ' use CurrentDb() to refresh Collections

    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of DocumentTheContainer is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of DocumentTheContainer is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    i = 0
    Set cnt = dbs.Containers(strContainerType)

    Select Case strContainerType
        Case "Forms": intAcObjType = 2   'acForm
        Case "Reports": intAcObjType = 3 'acReport
        Case "Scripts": intAcObjType = 4 'acMacro
        Case "Modules": intAcObjType = 5 'acModule
        Case Else
            MsgBox "Wrong Case Select in DocumentTheContainer"
    End Select

    If blnDebug Then Debug.Print UCase(strContainerType)

    For Each doc In cnt.Documents
        If blnDebug Then Debug.Print , doc.Name
        If Not (Left(doc.Name, 3) = "zzz" Or Left(doc.Name, 4) = "~TMP") Then
            i = i + 1
            strTheCurrentPathAndFile = aestrSourceLocation & doc.Name & "." & strExt
            'If strTheCurrentPathAndFile = "C:\ae\aezdb\src\basTranslate.bas" Then Debug.Print ">A:Here", doc.Name, strTheCurrentPathAndFile
            If IsFileLocked(strTheCurrentPathAndFile) Then
                MsgBox strTheCurrentPathAndFile & " is locked!", vbCritical, "STOP in DocumentTheContainer"
                'Stop
            End If
            'If strTheCurrentPathAndFile = "C:\ae\aezdb\src\basTranslate.bas" Then Debug.Print ">B:Here", doc.Name, strTheCurrentPathAndFile
            KillProperly (strTheCurrentPathAndFile)
            'If strTheCurrentPathAndFile = "C:\ae\aezdb\src\basTranslate.bas" Then Debug.Print ">C:Here", doc.Name, strTheCurrentPathAndFile
SaveAsText:
            Application.SaveAsText intAcObjType, doc.Name, strTheCurrentPathAndFile
            ' Convert UTF-16 to txt - fix for Access 2013
            If aeReadWriteStream(strTheCurrentPathAndFile) = True Then
                'If intAcObjType = 2 Then Pause (0.25)
                KillProperly (strTheCurrentPathAndFile)
                Name strTheCurrentPathAndFile & ".clean.txt" As strTheCurrentPathAndFile
            End If
        End If
    Next doc

    If blnDebug Then
        Debug.Print , i & " EXPORTED!"
        Debug.Print , cnt.Documents.Count & " EXISTING!"
    End If

    DocumentTheContainer = True

PROC_EXIT:
    Set doc = Nothing
    Set cnt = Nothing
    Set dbs = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    If Err = 2220 Then
        Debug.Print "Err=3220 : Resume SaveAsText - " & doc.Name & " - " & strTheCurrentPathAndFile
        Err.Clear
        Pause (0.25)
        Resume SaveAsText
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DocumentTheContainer of Class aegitClass"
    DocumentTheContainer = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Sub KillAllFiles(strLoc As String, Optional varDebug As Variant)

    Dim strFile As String
    Dim blnDebug As Boolean

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "KillAllFiles"

    Debug.Print "aeDocumentTheDatabase"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of KillAllFiles is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of KillAllFiles is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    If strLoc = "src" Then
        ' Delete all the exported src files
        strFile = Dir(aestrSourceLocation & "*.*")
        Do While strFile <> ""
            KillProperly (aestrSourceLocation & strFile)
            ' Need to specify full path again because a file was deleted
            strFile = Dir(aestrSourceLocation & "*.*")
        Loop
        strFile = Dir(aestrSourceLocation & "xml\" & "*.*")
        Do While strFile <> ""
            KillProperly (aestrSourceLocation & "xml\" & strFile)
            ' Need to specify full path again because a file was deleted
            strFile = Dir(aestrSourceLocation & "xml\" & "*.*")
        Loop
    ElseIf strLoc = "xml" Then
        ' Delete files in xml location
        If aegitSetup Then
            strFile = Dir(aestrXMLLocation & "*.*")
            Do While strFile <> ""
                KillProperly (aestrXMLLocation & strFile)
                ' Need to specify full path again because a file was deleted
                strFile = Dir(aestrXMLLocation & "*.*")
            Loop
        End If
    Else
        MsgBox "Bad strLoc", vbCritical, "STOP"
        Stop
    End If

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    If Err = 70 Then    ' Permission denied
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure KillAllFiles of Class aegitClass" _
            & vbCrLf & vbCrLf & _
            "Manually delete the files from git, compact and repair database, then try again!", vbCritical, "STOP"
        Stop
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure KillAllFiles of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure KillAllFiles of Class aegitClass"
    GlobalErrHandler
    Resume PROC_EXIT

End Sub

Private Function FolderExists(strPath As String) As Boolean
' Ref: http://allenbrowne.com/func-11.html
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Private Sub ListOrCloseAllOpenQueries(Optional strCloseAll As Variant)
' Ref: http://msdn.microsoft.com/en-us/library/office/aa210652(v=office.11).aspx

    Dim obj As AccessObject
    Dim dbs As Object
    Set dbs = Application.CurrentData

    If IsMissing(strCloseAll) Then
        ' Search for open AccessObject objects in AllQueries collection.
        For Each obj In dbs.AllQueries
            If obj.IsLoaded = True Then
                ' Print name of obj
                Debug.Print obj.Name
            End If
        Next obj
    Else
        For Each obj In dbs.AllQueries
            If obj.IsLoaded = True Then
                ' Close obj
                DoCmd.Close acQuery, obj.Name, acSaveYes
                Debug.Print "Closed query " & obj.Name
            End If
        Next obj
    End If

End Sub

Private Function aeDocumentTheDatabase(Optional varDebug As Variant) As Boolean
' Based on sample code from Arvin Meyer (MVP) June 2, 1999
' Ref: http://www.accessmvp.com/Arvin/DocDatabase.txt
' Ref: http://www.tek-tips.com/faqs.cfm?fid=6905
'====================================================================
' Author:   Peter F. Ennis
' Date:     February 8, 2011
' Comment:  Uses the undocumented [Application.SaveAsText] syntax
'           To reload use the syntax [Application.LoadFromText]
'           Add explicit references for DAO
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim dbs As DAO.Database
    Dim cnt As DAO.Container
    Dim doc As DAO.Document
    Dim qdf As DAO.QueryDef
    Dim i As Integer
    Dim Debugit As Boolean

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeDocumentTheDatabase"

    Debug.Print "aeDocumentTheDatabase"
    If IsMissing(varDebug) Then
        Debugit = False
        Debug.Print , "varDebug IS missing so blnDebug of aeDocumentTheDatabase is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debugit = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeDocumentTheDatabase is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If
    
    If aegitSourceFolder = "default" Then
        aestrSourceLocation = aegitType.SourceFolder
        aegitSetup = True
    Else
        aestrSourceLocation = aegitSourceFolder
    End If

    If aegitImportFolder = "default" Then
        aestrImportLocation = aegitType.ImportFolder
    End If

    If aegitUseImportFolder Then
        aestrImportLocation = aegitImportFolder
    End If

    If aegitXMLFolder = "default" Then
        aestrXMLLocation = aegitType.XMLFolder
    End If

    ListOrCloseAllOpenQueries

    If Debugit Then
        Debug.Print , ">==> aeDocumentTheDatabase >==>"
        Debug.Print , "Property Get SourceFolder = " & aestrSourceLocation
        Debug.Print , "aegitUseImportFolder = " & aegitUseImportFolder
        Debug.Print , "Property Get ImportFolder = " & aestrImportLocation
        Debug.Print , "Property Get XMLFolder = " & aestrXMLLocation
    End If
    'Stop

    If IsMissing(varDebug) Then
        KillAllFiles "src"
    Else
        KillAllFiles "src", varDebug
    End If

    ' NOTE: Erl(0) Error 2950 if the ouput location does not exist so test for it first.
    If FolderExists(aestrSourceLocation) Then
        If Debugit Then
            DocumentTheContainer "Forms", "frm", "WithDebugging"
            DocumentTheContainer "Reports", "rpt", "WithDebugging"
            DocumentTheContainer "Scripts", "mac", "WithDebugging"
            DocumentTheContainer "Modules", "bas", "WithDebugging"
        Else
            DocumentTheContainer "Forms", "frm"
            DocumentTheContainer "Reports", "rpt"
            DocumentTheContainer "Scripts", "mac"
            DocumentTheContainer "Modules", "bas"
        End If
    Else
        MsgBox aestrSourceLocation & " Does not exist!", vbCritical, "aegit"
        Stop
    End If
    
    ListOfContainers "ListOfContainers.txt"
    ListOfAllHiddenQueries

    If Debugit Then
        ListOfAccessApplicationOptions "Debug"
    Else
        ListOfAccessApplicationOptions
    End If

    ListOfApplicationProperties
    'Stop

    Set dbs = CurrentDb() ' use CurrentDb() to refresh Collections

    '=============
    ' QUERIES
    '=============
    i = 0
    If Debugit Then Debug.Print "QUERIES"

    ' Delete all TEMP queries ...
    For Each qdf In CurrentDb.QueryDefs
        If Left(qdf.Name, 1) = "~" Then
            CurrentDb.QueryDefs.Delete qdf.Name
            CurrentDb.QueryDefs.Refresh
        End If
    Next qdf

    For Each qdf In CurrentDb.QueryDefs
        If Debugit Then Debug.Print , qdf.Name
        If Not (Left(qdf.Name, 4) = "MSys" Or Left(qdf.Name, 4) = "~sq_" _
                        Or Left(qdf.Name, 4) = "~TMP" _
                        Or Left(qdf.Name, 3) = "zzz") Then
            i = i + 1
            Application.SaveAsText acQuery, qdf.Name, aestrSourceLocation & qdf.Name & ".qry"
            ' Convert UTF-16 to txt - fix for Access 2013
            If aeReadWriteStream(aestrSourceLocation & qdf.Name & ".qry") = True Then
                KillProperly (aestrSourceLocation & qdf.Name & ".qry")
                Name aestrSourceLocation & qdf.Name & ".qry" & ".clean.txt" As aestrSourceLocation & qdf.Name & ".qry"
            End If
        End If
    Next qdf

    If Debugit Then
        If i = 1 Then
            Debug.Print , "1 Query EXPORTED!"
        Else
            Debug.Print , i & " Queries EXPORTED!"
        End If
        
        If CurrentDb.QueryDefs.Count = 1 Then
            Debug.Print , "1 Query EXISTING!"
        Else
            Debug.Print , CurrentDb.QueryDefs.Count & " Queries EXISTING!"
        End If
    End If

    OutputQueriesSqlText
    OutputBuiltInPropertiesText
    OutputFieldLookupControlTypeList
    OutputTheSchemaFile
    OutputAllContainerProperties
    
    If Debugit Then
        OutputPrinterInfo "Debug"
    Else
        OutputPrinterInfo
    End If

    aeDocumentTheDatabase = True

PROC_EXIT:
    Set qdf = Nothing
    Set doc = Nothing
    Set cnt = Nothing
    Set dbs = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTheDatabase of Class aegitClass"
    If Debugit Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTheDatabase of Class aegitClass"
    aeDocumentTheDatabase = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function BuildTheDirectory(fso As Object, _
                                        Optional varDebug As Variant) As Boolean
'Private Function BuildTheDirectory(FSO As Scripting.FileSystemObject, _
                                        Optional varDebug As Variant) As Boolean
'*** Requires reference to "Microsoft Scripting Runtime"
'
' Ref: http://msdn.microsoft.com/en-us/library/ebkhfaaz(v=vs.85).aspx
'====================================================================
' Author:   Peter F. Ennis
' Date:     February 8, 2011
' Comment:  Add optional debug parameter
' Requires: Reference to Microsoft Scripting Runtime
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim objImportFolder As Object
    Dim blnDebug As Boolean

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "BuildTheDirectory"

    Debug.Print "BuildTheDirectory"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of BuildTheDirectory is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of BuildTheDirectory is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    If blnDebug Then Debug.Print , ">==> BuildTheDirectory >==>"

    ' Bail out if (a) the drive does not exist, or if (b) the directory already exists.

    If blnDebug Then Debug.Print , , "THE_DRIVE = " & THE_DRIVE
    If blnDebug Then Debug.Print , , "FSO.DriveExists(THE_DRIVE) = " & fso.DriveExists(THE_DRIVE)
    If Not fso.DriveExists(THE_DRIVE) Then
        If blnDebug Then Debug.Print , , "FSO.DriveExists(THE_DRIVE) = FALSE - The drive DOES NOT EXIST !!!"
        BuildTheDirectory = False
        Exit Function
    End If
    If blnDebug Then Debug.Print , , "The drive EXISTS !!!"
    If blnDebug Then Debug.Print , , "aegitUseImportFolder = " & aegitUseImportFolder
    
    If aegitImportFolder = "default" Then
        aestrImportLocation = aegitType.ImportFolder
    End If
    If aegitUseImportFolder And aegitImportFolder <> "default" Then
        aestrImportLocation = aegitImportFolder
    End If
        
    If blnDebug Then Debug.Print , , "The import directory is: " & aestrImportLocation
   
    If fso.FolderExists(aestrImportLocation) Then
        If blnDebug Then Debug.Print , , "FSO.FolderExists(aestrImportLocation) = TRUE - The directory EXISTS !!!"
        BuildTheDirectory = False
        Exit Function
    End If
    If blnDebug Then Debug.Print , , "The import directory does NOT EXIST !!!"

    If aegitUseImportFolder Then
        Set objImportFolder = fso.CreateFolder(aestrImportLocation)
        If blnDebug Then Debug.Print , , aestrImportLocation & " has been CREATED !!!"
    End If

    BuildTheDirectory = True

PROC_EXIT:
    Set objImportFolder = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure BuildTheDirectory of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure BuildTheDirectory of Class aegitClass"
    BuildTheDirectory = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function aeReadDocDatabase(blnImport As Boolean, Optional varDebug As Variant) As Boolean
' VBScript makes use of ADOX (Microsoft's Active Data Objects Extensions for Data Definition Language and Security)
' to create a query on a Microsoft Access database
' Ref: http://stackoverflow.com/questions/859530/alternative-to-application-loadfromtext-for-ms-access-queries
' Microsoft Access Stored Queries and VBscript: How to Create and Edit a Stored Database Query
' Ref: http://www.suite101.com/content/microsoft-access-stored-queries-and-vbscript-a87978#ixzz1D32Vqbso
' Using WScript
' Ref: http://www.codeforexcelandoutlook.com/vba/shell-scripting-using-vba-and-the-windows-script-host-object-model/
' vbscript get file extension
' Ref: http://www.experts-exchange.com/Programming/Languages/Visual_Basic/Q_24297896.html
'
'====================================================================
' Author:   Peter F. Ennis
' Date:     February 8, 2011
' Comment:  Add explicit references for objects, wscript, fso
' Requires: Reference to Microsoft Scripting Runtime
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim MyFile As Object
    Dim strFileType As String
    Dim strFileBaseName As String
    Dim bln As Boolean
    Dim blnDebug As Boolean

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeReadDocDatabase"

    Debug.Print "aeReadDocDatabase"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeReadDocDatabase is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeReadDocDatabase is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    If Not blnImport Then
        Debug.Print , "blnImport IS FALSE so exit aeReadDocDatabase"
        aeReadDocDatabase = False
        Exit Function
    End If
    
    Const acQuery = 1

    If aegitSourceFolder = "default" Then
        aestrSourceLocation = aegitType.SourceFolder
    Else
        aestrSourceLocation = aegitSourceFolder
    End If

    If aegitImportFolder = "default" Then
        aestrImportLocation = aegitType.ImportFolder
    End If
    If aegitUseImportFolder And aegitImportFolder <> "default" Then
        aestrImportLocation = aegitImportFolder
    End If

    If blnDebug Then
        Debug.Print ">==> aeReadDocDatabase >==>"
        Debug.Print , "aegit VERSION: " & aegitVERSION
        Debug.Print , "aegit VERSION_DATE: " & aegitVERSION_DATE
        Debug.Print , "SourceFolder = " & aestrSourceLocation
        Debug.Print , "UseImportFolder = " & aegitUseImportFolder
        Debug.Print , "ImportFolder = " & aestrImportLocation
        'Stop
    End If

    ' Create needed objects
    Dim wsh As Object  ' As Object if late-bound
    Set wsh = CreateObject("WScript.Shell")

    wsh.CurrentDirectory = aestrImportLocation
    If blnDebug Then Debug.Print , "wsh.CurrentDirectory = " & wsh.CurrentDirectory
    ' CurDir Function
    If blnDebug Then Debug.Print , "CurDir = " & CurDir

    ' Create needed objects
    Dim fso As Object
'    Dim FSO As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    If blnDebug Then
        bln = BuildTheDirectory(fso, "WithDebugging")
        Debug.Print , "<==<"
    Else
        bln = BuildTheDirectory(fso)
    End If

    If aegitUseImportFolder Then
        Dim objFolder As Object
        Set objFolder = fso.GetFolder(aegitType.ImportFolder)

        For Each MyFile In objFolder.Files
            If blnDebug Then Debug.Print "myFile = " & MyFile
            If blnDebug Then Debug.Print "myFile.Name = " & MyFile.Name
            strFileBaseName = fso.GetBaseName(MyFile.Name)
            strFileType = fso.GetExtensionName(MyFile.Name)
            If blnDebug Then Debug.Print strFileBaseName & " (" & strFileType & ")"

            If (strFileType = "frm") Then
                If Exists("FORMS", strFileBaseName) Then
                    MsgBox "Skipping: FORM " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                    If blnDebug Then Debug.Print "Skipping: FORM " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
                Else
                    Application.LoadFromText acForm, strFileBaseName, MyFile.Path
                End If
            ElseIf (strFileType = "rpt") Then
                If Exists("REPORTS", strFileBaseName) Then
                    MsgBox "Skipping: REPORT " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                    If blnDebug Then Debug.Print "Skipping: REPORT " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
                Else
                    Application.LoadFromText acReport, strFileBaseName, MyFile.Path
                End If
            ElseIf (strFileType = "bas") Then
                If Exists("MODULES", strFileBaseName) Then
                    MsgBox "Skipping: MODULE " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                    If blnDebug Then Debug.Print "Skipping: MODULE " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
                Else
                    Application.LoadFromText acModule, strFileBaseName, MyFile.Path
                End If
            ElseIf (strFileType = "mac") Then
                If Exists("MACROS", strFileBaseName) Then
                    MsgBox "Skipping: MACRO " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                    If blnDebug Then Debug.Print "Skipping: MACRO " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
                Else
                    Application.LoadFromText acMacro, strFileBaseName, MyFile.Path
                End If
            ElseIf (strFileType = "qry") Then
                If Exists("QUERIES", strFileBaseName) Then
                    MsgBox "Skipping: QUERY " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                    If blnDebug Then Debug.Print "Skipping: QUERY " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
                Else
                    Application.LoadFromText acQuery, strFileBaseName, MyFile.Path
                End If
            End If
        Next
    End If

    If blnDebug Then Debug.Print "<==<"
    
    aeReadDocDatabase = True

PROC_EXIT:
    Set MyFile = Nothing
    'Set ojbFolder = Nothing
    Set fso = Nothing
    Set wsh = Nothing
    aeReadDocDatabase = True
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadDocDatabase of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadDocDatabase of Class aegitClass"
    aeReadDocDatabase = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function aeExists(strAccObjType As String, _
                        strAccObjName As String, Optional varDebug As Variant) As Boolean
' Ref: http://vbabuff.blogspot.com/2010/03/does-access-object-exists.html
'
'====================================================================
' Author:     Peter F. Ennis
' Date:       February 18, 2011
' Comment:    Return True if the object exists
' Parameters:
'             strAccObjType: "Tables", "Queries", "Forms",
'                            "Reports", "Macros", "Modules"
'             strAccObjName: The name of the object
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim objType As Object
    Dim obj As Variant
    Dim blnDebug As Boolean
    
    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeExists"

    Debug.Print "aeExists"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeExists is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeExists is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    If blnDebug Then Debug.Print ">==> aeExists >==>"

    Select Case strAccObjType
        Case "Tables"
            Set objType = CurrentDb.TableDefs
        Case "Queries"
            Set objType = CurrentDb.QueryDefs
        Case "Forms"
            Set objType = CurrentProject.AllForms
        Case "Reports"
            Set objType = CurrentProject.AllReports
        Case "Macros"
            Set objType = CurrentProject.AllMacros
        Case "Modules"
            Set objType = CurrentProject.AllModules
        Case Else
            MsgBox "Wrong option!", vbCritical, "in procedure aeExists of Class aegitClass"
            If blnDebug Then
                Debug.Print , "strAccObjType = >" & strAccObjType & "< is  a false value"
                Debug.Print , "Option allowed is one of 'Tables', 'Queries', 'Forms', 'Reports', 'Macros', 'Modules'"
                Debug.Print "<==<"
            End If
            aeExists = False
            Set obj = Nothing
            Exit Function
    End Select

    If blnDebug Then Debug.Print , "strAccObjType = " & strAccObjType
    If blnDebug Then Debug.Print , "strAccObjName = " & strAccObjName

    For Each obj In objType
        If blnDebug Then Debug.Print , obj.Name, strAccObjName
        If obj.Name = strAccObjName Then
            If blnDebug Then
                Debug.Print , strAccObjName & " EXISTS!"
                Debug.Print "<==<"
            End If
            aeExists = True
            Set obj = Nothing
            Exit Function ' Found it!
        Else
            aeExists = False
        End If
    Next
    If blnDebug Then
        Debug.Print , strAccObjName & " DOES NOT EXIST!"
        Debug.Print "<==<"
    End If

    aeExists = True

PROC_EXIT:
    Set obj = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class aegitClass"
    aeExists = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function GetType(Value As Long) As String
' Ref: http://bytes.com/topic/access/answers/557780-getting-string-name-enum

    Select Case Value
        Case acCheckBox
            GetType = "CheckBox"
        Case acTextBox
            GetType = "TextBox"
        Case acListBox
            GetType = "ListBox"
        Case acComboBox
            GetType = "ComboBox"
        Case Else
    End Select

End Function

Private Sub OutputFieldLookupControlTypeList()
    Dim bln As Boolean
    bln = FieldLookupControlTypeList
End Sub

Private Function FieldLookupControlTypeList(Optional varDebug As Variant) As Boolean
' Ref: http://support.microsoft.com/kb/304274
' Ref: http://msdn.microsoft.com/en-us/library/office/bb225848(v=office.12).aspx
' 106 - acCheckBox, 109 - acTextBox, 110 - acListBox, 111 - acComboBox

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "FieldLookupControlTypeList"

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDefs
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    Dim lng As Long
    Dim strChkTbl As String
    Dim strChkFld As String

    ' Counters for DisplayControl types
    Static intChk As Integer
    Static intTxt As Integer
    Static intLst As Integer
    Static intCbo As Integer
    Static intAllFieldsCount As Integer
    Static intElse As Integer

    Set dbs = CurrentDb()
    Set tdf = dbs.TableDefs

    Dim fle As Integer

    fle = FreeFile()
    Open aegitSourceFolder & "\" & aeFLkCtrFile For Output As #fle

    intChk = 0
    intTxt = 0
    intLst = 0
    intCbo = 0
    intAllFieldsCount = 0
    intElse = 0

    On Error Resume Next
    For Each tbl In tdf
        If Left(tbl.Name, 4) <> "MSys" And Left(tbl.Name, 3) <> "zzz" _
            And Left(tbl.Name, 1) <> "~" Then
            'Debug.Print tbl.Name
            Print #fle, tbl.Name
            For Each fld In tbl.Fields
                intAllFieldsCount = intAllFieldsCount + 1
                lng = fld.Properties("DisplayControl").Value
                'Debug.Print , fld.Name, lng, GetType(lng)
                Print #fle, , fld.Name, lng, GetType(lng)
                Select Case lng
                    Case acCheckBox
                        intChk = intChk + 1
                        'Debug.Print intChk, ">Here"
                        strChkTbl = tbl.Name
                        strChkFld = fld.Name
                    Case acTextBox
                        intTxt = intTxt + 1
                        'Debug.Print intTxt, ">Here"
                    Case acListBox
                        intLst = intLst + 1
                        'Debug.Print intLst, ">Here"
                    Case acComboBox
                        intCbo = intCbo + 1
                        'Debug.Print intCbo, ">Here"
                    Case Else
                        intElse = intElse + 1
                        'MsgBox "lng=" & lng
                End Select
            Next fld
        End If
    Next tbl
    If Not IsMissing(varDebug) Then Debug.Print "Count of Check box = " & intChk
    If Not IsMissing(varDebug) Then Debug.Print "Count of Text box  = " & intTxt
    If Not IsMissing(varDebug) Then Debug.Print "Count of List box  = " & intLst
    If Not IsMissing(varDebug) Then Debug.Print "Count of Combo box = " & intCbo
    If Not IsMissing(varDebug) Then Debug.Print "Count of Else      = " & intElse
    If Not IsMissing(varDebug) Then Debug.Print "Count of Display Controls = " & intChk + intTxt + intLst + intCbo
    If Not IsMissing(varDebug) Then Debug.Print "Count of All Fields = " & intAllFieldsCount - intElse
    'Debug.Print "Table with check box is " & strChkTbl
    'Debug.Print "Field with check box is " & strChkFld

    Print #fle, "Count of Check box = " & intChk
    Print #fle, "Count of Text box  = " & intTxt
    Print #fle, "Count of List box  = " & intLst
    Print #fle, "Count of Combo box = " & intCbo
    Print #fle, "Count of Else      = " & intElse
    Print #fle, "Count of Display Controls = " & intChk + intTxt + intLst + intCbo
    Print #fle, "Count of All Fields = " & intAllFieldsCount - intElse
    'Print #fle, "Table with check box is " & strChkTbl
    'Print #fle, "Field with check box is " & strChkFld

    If intAllFieldsCount - intElse = intChk + intTxt + intLst + intCbo Then
        FieldLookupControlTypeList = True
    Else
        FieldLookupControlTypeList = False
    End If

PROC_EXIT:
    On Error Resume Next
    Close fle
    Set tdf = Nothing
    Set dbs = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FieldLookupControlTypeList of Class aegitClass", vbCritical, "Error"
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Public Function ListOfContainers(strTheFileName As String) As Boolean
' Ref: http://www.susandoreydesigns.com/software/AccessVBATechniques.pdf
' Ref: http://msdn.microsoft.com/en-us/library/office/bb177484(v=office.12).aspx

    Dim dbs As DAO.Database
    Dim conItem As DAO.Container
    Dim prpLoop As DAO.Property
    Dim strName As String
    Dim strOwner As String
    Dim strText As String
    Dim strFile As String
    Dim lngFileNum As Long

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "ListOfContainers"

    Set dbs = CurrentDb
    lngFileNum = FreeFile

    strFile = aestrSourceLocation & strTheFileName

    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As lngFileNum
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As lngFileNum
    End If

    With dbs
        ' Enumerate Containers collection.
        For Each conItem In .Containers
            'Debug.Print "Properties of " & conItem.Name _
                & " container"
            WriteStringToFile lngFileNum, "Properties of " & conItem.Name _
                & " container", strFile
            
            ' Enumerate Properties collection of each Container object.
            For Each prpLoop In conItem.Properties
                'Debug.Print "  " & prpLoop.Name _
                    & " = "; prpLoop
                WriteStringToFile lngFileNum, "  " & prpLoop.Name _
                    & " = " & prpLoop, strFile
            Next prpLoop
        Next conItem
        .Close
    End With

    ListOfContainers = True

PROC_EXIT:
    Set prpLoop = Nothing
    Set conItem = Nothing
    Set dbs = Nothing
    Close lngFileNum
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListOfContainers of Class aegitClass", vbCritical, "Error"
    ListOfContainers = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Public Sub OutputAllContainerProperties(Optional varDebug As Variant)

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "OutputAllContainerProperties"

    If Not IsMissing(varDebug) Then
        Debug.Print "Container information for properties of saved Databases"
        ListAllContainerProperties "Databases", varDebug
        Debug.Print "Container information for properties of saved Tables and Queries"
        ListAllContainerProperties "Tables", varDebug
        Debug.Print "Container information for properties of saved Relationships"
        ListAllContainerProperties "Relationships", varDebug
    Else
        ListAllContainerProperties "Databases"
        ListAllContainerProperties "Tables"
        ListAllContainerProperties "Relationships"
    End If

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputAllContainerProperties of Class aegitClass"
    'If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputAllContainerProperties of Class aegitClass"
    GlobalErrHandler
    Resume PROC_EXIT

End Sub

Private Function fListGUID(strTableName As String) As String
' Ref: http://stackoverflow.com/questions/8237914/how-to-get-the-guid-of-a-table-in-microsoft-access
' e.g. ?fListGUID("tblThisTableHasSomeReallyLongNameButItCouldBeMuchLonger")

    Dim i As Integer
    Dim arrGUID8() As Byte
    Dim strArrGUID8(8) As String
    Dim strGuid As String

    strGuid = ""
    arrGUID8 = CurrentDb.TableDefs(strTableName).Properties("GUID").Value
    For i = 1 To 8
        If Len(Hex(arrGUID8(i))) = 1 Then
            strArrGUID8(i) = "0" & CStr(Hex(arrGUID8(i)))
        Else
            strArrGUID8(i) = Hex(arrGUID8(i))
        End If
    Next

    For i = 1 To 8
        strGuid = strGuid & strArrGUID8(i) & "-"
    Next
    fListGUID = Left(strGuid, 23)

End Function

Private Sub ListAllContainerProperties(strContainer As String, Optional varDebug As Variant)
' Ref: http://www.dbforums.com/microsoft-access/1620765-read-ms-access-table-properties-using-vba.html
' Ref: http://msdn.microsoft.com/en-us/library/office/aa139941(v=office.10).aspx
    
    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "ListAllContainerProperties"

    Dim dbs As DAO.Database
    Dim obj As Object
    Dim prp As DAO.Property
    Dim doc As Document
    Dim fle As Integer

    Set dbs = Application.CurrentDb
    Set obj = dbs.Containers(strContainer)

    fle = FreeFile()
    Open aegitSourceFolder & "\OutputContainer" & strContainer & "Properties.txt" For Output As #fle

    ' Ref: http://stackoverflow.com/questions/16642362/how-to-get-the-following-code-to-continue-on-error
    For Each doc In obj.Documents
        If Left(doc.Name, 4) <> "MSys" And Left(doc.Name, 3) <> "zzz" _
            And Left(doc.Name, 1) <> "~" Then
            If Not IsMissing(varDebug) Then Debug.Print ">>>" & doc.Name
            Print #fle, ">>>" & doc.Name
            For Each prp In doc.Properties
                On Error Resume Next
                If prp.Name = "GUID" And strContainer = "tables" Then
                    If Not IsMissing(varDebug) Then Debug.Print , prp.Name, fListGUID(doc.Name)
                        Print #fle, , prp.Name, fListGUID(doc.Name)
                    ElseIf prp.Name = "DOL" Then
                        If Not IsMissing(varDebug) Then Debug.Print prp.Name, "Track name AutoCorrect info is ON!"
                        Print #fle, , prp.Name, "Track name AutoCorrect info is ON!"
                    ElseIf prp.Name = "NameMap" Then
                        If Not IsMissing(varDebug) Then Debug.Print , prp.Name, "Track name AutoCorrect info is ON!"
                        Print #fle, , prp.Name, "Track name AutoCorrect info is ON!"
                    Else
                        If Not IsMissing(varDebug) Then Debug.Print , prp.Name, prp.Value
                        Print #fle, , prp.Name, prp.Value
                    End If
                On Error GoTo 0
            Next
        End If
    Next

    Set obj = Nothing
    Set dbs = Nothing
    Close fle

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListAllContainerProperties of Class aegitClass"
    'If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListAllContainerProperties of Class aegitClass"
    GlobalErrHandler
    Resume PROC_EXIT

End Sub

Private Sub OutputTheTableDataAsXML(avarTableNames() As Variant, Optional varDebug As Variant)
' Ref: http://wiki.lessthandot.com/index.php/Output_Access_/_Jet_to_XML
' Ref: http://msdn.microsoft.com/en-us/library/office/aa164887(v=office.10).aspx

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "OutputTheTableDataAsXML"

    Const adOpenStatic = 3
    Const adLockOptimistic = 3
    Const adPersistXML = 1

    Dim strFileName As String

    Dim cnn As Object
    Set cnn = CurrentProject.Connection
    Dim rst As Object
    Set rst = CreateObject("ADODB.Recordset")

    If IsMissing(varDebug) Then
        KillAllFiles "xml"
    Else
        KillAllFiles "xml", varDebug
    End If

    rst.Open "Select * from " & avarTableNames(1), cnn, adOpenStatic, adLockOptimistic

    strFileName = aestrXMLLocation & avarTableNames(1) & ".xml"

    If aegitSetup Then
        MsgBox "aegitSetup=True aestrXMLLocation=" & aestrXMLLocation
        If Not rst.EOF Then
            rst.MoveFirst
            rst.Save strFileName, adPersistXML
        End If
    Else
    End If

    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputTheTableDataAsXML of Class aegitClass"
    'If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputTheTableDataAsXML of Class aegitClass"
    GlobalErrHandler
    Resume PROC_EXIT

End Sub

Public Sub OutputPrinterInfo(Optional varDebug As Variant)
' Ref: http://msdn.microsoft.com/en-us/library/office/aa139946(v=office.10).aspx
' Ref: http://answers.microsoft.com/en-us/office/forum/office_2010-access/how-do-i-change-default-printers-in-vba/d046a937-6548-4d2b-9517-7f622e2cfed2

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "PrinterInfo"

    Dim prt As Printer
    Dim prtCount As Integer
    Dim i As Integer
    Dim fle As Integer

    fle = FreeFile()
    Open aegitSourceFolder & "\OutputPrinterInfo.txt" For Output As #fle

    If Not IsMissing(varDebug) Then Debug.Print "Default Printer=" & Application.Printer.DeviceName
    Print #fle, "Default Printer=" & Application.Printer.DeviceName
    prtCount = Application.Printers.Count
    If Not IsMissing(varDebug) Then Debug.Print "Number of Printers=" & prtCount
    Print #fle, "Number of Printers=" & prtCount
    For Each prt In Printers
        If Not IsMissing(varDebug) Then Debug.Print , prt.DeviceName
        Print #fle, , prt.DeviceName
    Next prt

    If Not IsMissing(varDebug) Then
        For i = 0 To prtCount - 1
            Debug.Print "DeviceName=" & Application.Printers(i).DeviceName
            Debug.Print , "BottomMargin=" & Application.Printers(i).BottomMargin
            Debug.Print , "ColorMode=" & Application.Printers(i).ColorMode
            Debug.Print , "ColumnSpacing=" & Application.Printers(i).ColumnSpacing
            Debug.Print , "Copies=" & Application.Printers(i).Copies
            Debug.Print , "DataOnly=" & Application.Printers(i).DataOnly
            Debug.Print , "DefaultSize=" & Application.Printers(i).DefaultSize
            Debug.Print , "DriverName=" & Application.Printers(i).DriverName
            Debug.Print , "Duplex=" & Application.Printers(i).Duplex
            Debug.Print , "ItemLayout=" & Application.Printers(i).ItemLayout
            Debug.Print , "ItemsAcross=" & Application.Printers(i).ItemsAcross
            Debug.Print , "ItemSizeHeight=" & Application.Printers(i).ItemSizeHeight
            Debug.Print , "ItemSizeWidth=" & Application.Printers(i).ItemSizeWidth
            Debug.Print , "LeftMargin=" & Application.Printers(i).LeftMargin
            Debug.Print , "Orientation=" & Application.Printers(i).Orientation
            Debug.Print , "PaperBin=" & Application.Printers(i).PaperBin
            Debug.Print , "PaperSize=" & Application.Printers(i).PaperSize
            Debug.Print , "Port=" & Application.Printers(i).Port
            Debug.Print , "PrintQuality=" & Application.Printers(i).PrintQuality
            Debug.Print , "RightMargin=" & Application.Printers(i).RightMargin
            Debug.Print , "RowSpacing=" & Application.Printers(i).RowSpacing
            Debug.Print , "TopMargin=" & Application.Printers(i).TopMargin
        Next
    End If

    For i = 0 To prtCount - 1
        Print #fle, "DeviceName=" & Application.Printers(i).DeviceName
        Print #fle, , "BottomMargin=" & Application.Printers(i).BottomMargin
        Print #fle, , "ColorMode=" & Application.Printers(i).ColorMode
        Print #fle, , "ColumnSpacing=" & Application.Printers(i).ColumnSpacing
        Print #fle, , "Copies=" & Application.Printers(i).Copies
        Print #fle, , "DataOnly=" & Application.Printers(i).DataOnly
        Print #fle, , "DefaultSize=" & Application.Printers(i).DefaultSize
        Print #fle, , "DriverName=" & Application.Printers(i).DriverName
        Print #fle, , "Duplex=" & Application.Printers(i).Duplex
        Print #fle, , "ItemLayout=" & Application.Printers(i).ItemLayout
        Print #fle, , "ItemsAcross=" & Application.Printers(i).ItemsAcross
        Print #fle, , "ItemSizeHeight=" & Application.Printers(i).ItemSizeHeight
        Print #fle, , "ItemSizeWidth=" & Application.Printers(i).ItemSizeWidth
        Print #fle, , "LeftMargin=" & Application.Printers(i).LeftMargin
        Print #fle, , "Orientation=" & Application.Printers(i).Orientation
        Print #fle, , "PaperBin=" & Application.Printers(i).PaperBin
        Print #fle, , "PaperSize=" & Application.Printers(i).PaperSize
        Print #fle, , "Port=" & Application.Printers(i).Port
        Print #fle, , "PrintQuality=" & Application.Printers(i).PrintQuality
        Print #fle, , "RightMargin=" & Application.Printers(i).RightMargin
        Print #fle, , "RowSpacing=" & Application.Printers(i).RowSpacing
        Print #fle, , "TopMargin=" & Application.Printers(i).TopMargin
    Next

PROC_EXIT:
    Close fle
    PopCallStack
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case 9
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputPrinterInfo of Class aegitClass"
            GlobalErrHandler
            Resume Next
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputPrinterInfo of Class aegitClass"
            GlobalErrHandler
            Resume Next
    End Select

End Sub

'==================================================
' Global Error Handler Routines
' Ref: http://msdn.microsoft.com/en-us/library/office/ee358847(v=office.12).aspx#odc_ac2007_ta_ErrorHandlingAndDebuggingTipsForAccessVBAndVBA_WritingCodeForDebugging
'==================================================

Private Sub ResetWorkspace()
    Dim intCounter As Integer

    On Error Resume Next

    Application.MenuBar = ""
    DoCmd.SetWarnings False
    DoCmd.Hourglass False
    DoCmd.Echo True

    ' Clean up workspace by closing open forms and reports
    For intCounter = 0 To Forms.Count - 1
        DoCmd.Close acForm, Forms(intCounter).Name
    Next intCounter

    For intCounter = 0 To Reports.Count - 1
        DoCmd.Close acReport, Reports(intCounter).Name
    Next intCounter
End Sub

Private Sub GlobalErrHandler()
' Main procedure to handle errors that occur.

    Dim strError As String
    Dim lngError As Long
    Dim intErl As Integer
    Dim strMsg As String

    ' Variables to preserve error information
    strError = Err.Description
    lngError = Err.Number
    intErl = Erl

    ' Reset workspace, close open objects
    ResetWorkspace

    ' Prompt the user with information on the error:
    strMsg = "Procedure: " & CurrentProcName() & vbCrLf & _
             "Line: " & Erl & vbCrLf & _
             "Error: (" & lngError & ")" & strError & vbCrLf & _
             "Application Quit is turned OFF !!!"
    MsgBox strMsg, vbCritical, "GlobalErrHandler"

    ' Write error to file:
    WriteErrorToFile intErl, lngError, CurrentProcName(), strError

    ' Exit Access without saving any changes
    ' (you might want to change this to save all changes)

    'Application.Quit acExit
End Sub

Private Function CurrentProcName() As String
    CurrentProcName = mastrCallStack(mintStackPointer - 1)
End Function

Private Sub PushCallStack(strProcName As String)
' Add the current procedure name to the Call Stack.
' Should be called whenever a procedure is called

    On Error Resume Next

    ' Verify the stack array can handle the current array element
    If mintStackPointer > UBound(mastrCallStack) Then
    ' If array has not been defined, initialize the error handler
        If Err.Number = 9 Then
            ErrorHandlerInit
        Else
            ' Increase the size of the array to not go out of bounds
            ReDim Preserve mastrCallStack(UBound(mastrCallStack) + _
            mcintIncrementStackSize)
        End If
    End If

    On Error GoTo 0

    mastrCallStack(mintStackPointer) = strProcName

    ' Increment pointer to next element
    mintStackPointer = mintStackPointer + 1
End Sub

Private Sub ErrorHandlerInit()
    mfInErrorHandler = False
    mintStackPointer = 1
    ReDim mastrCallStack(1 To mcintIncrementStackSize)
End Sub

Private Sub PopCallStack()
' Remove a procedure name from the call stack

    If mintStackPointer <= UBound(mastrCallStack) Then
        mastrCallStack(mintStackPointer) = ""
    End If

    ' Reset pointer to previous element
    mintStackPointer = mintStackPointer - 1
End Sub

Private Sub WriteErrorToFile(intTheErl As Integer, lngTheErrorNum As Long, _
                strCurrentProcName As String, strErrorDescription As String)
    
    Dim strFilePath As String
    Dim lngFileNum As Long
    
    On Error Resume Next

    ' Write to a text file called aegitErrorLog in the MyDocuments folder
    strFilePath = CreateObject("WScript.Shell").SpecialFolders("MYDOCUMENTS") & "\aegitErrorLog.txt"

    lngFileNum = FreeFile
    Open strFilePath For Append Access Write Lock Write As lngFileNum
        Print #lngFileNum, Now(), intTheErl, lngTheErrorNum, strCurrentProcName, strErrorDescription
    Close lngFileNum

End Sub

Private Sub WriteStringToFile(lngFileNum As Long, strTheString As String, strTheAbsoluteFileName As String)
  
    On Error Resume Next

    Open strTheAbsoluteFileName For Append Access Write Lock Write As lngFileNum
        Print #lngFileNum, strTheString
    Close lngFileNum

End Sub