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
' =======================================================================
' Author:   Peter F. Ennis
' Date:     February 24, 2011
' Comment:  Create class for revision control
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
' Web:      https://github.com/peterennis/aegit
' Newspeak: http://www.collinsdictionary.com/dictionary/english/eejit
' Webspeak: https://disqus.com/home/discussion/fabiensanglardswebsite/git_source_code_review/
' =======================================================================

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Private Declare PtrSafe Function apiSetActiveWindow Lib "user32" Alias "SetActiveWindow" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Const EXCLUDE_1 As String = "aebasChangeLog_aegit_expClass"
Private Const EXCLUDE_2 As String = "aebasTEST_aegit_expClass"
Private Const EXCLUDE_3 As String = "aegit_expClass"

Private Const aegit_expVERSION As String = "1.4.9"
Private Const aegit_expVERSION_DATE As String = "August 6, 2015"
Private Const aeAPP_NAME As String = "aegit_exp"
Private Const mblnOutputPrinterInfo As Boolean = False
' If mblnUTF16 is True the form txt exported files will be UTF-16 Windows format
' If mblnUTF16 is False the BOM marker will be stripped and files will be ANSI
Private Const mblnUTF16 As Boolean = False

Private Enum SizeStringSide
    TextLeft = 1
    TextRight = 2
End Enum

Private Type myExclusions
    exclude1 As String
    exclude2 As String
    exclude3 As String
End Type

Private Type mySetupType
    SourceFolder As String
    ImportFolder As String
    UseImportFolder As Boolean
    XMLfolder As String
End Type

Private Type myExportType               ' Initialize defaults as:
    ExportAll As Boolean                ' True
    ExportCodeAndObjects As Boolean     ' True
    ExportModuleCodeOnly As Boolean     ' True
    ExportQAT As Boolean                ' True
    ExportCBID As Boolean               ' False
End Type

Private myExclude As myExclusions
Private pExclude As Boolean

Private aegitSetup As Boolean
Private aegitType As mySetupType
Private aegitExport As myExportType
Private aegitSourceFolder As String
Private aegitSourceFolderBe As String
Private aegitFrontEndApp As Boolean
Private aegitTextEncoding As String
Private aegitXMLfolder As String
Private aegitDataXML() As Variant
Private aegitExportDataToXML As Boolean
Private aestrSourceLocation As String
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
Private aestrBackEndDb1 As String
Private aestrPassword As String
Private Const aestr4 As String = "    "
Private Const aeSqlTxtFile As String = "OutputSqlCodeForQueries.txt"
Private Const aeTblTxtFile As String = "OutputTblSetupForTables.txt"
Private Const aeRefTxtFile As String = "OutputReferencesSetup.txt"
Private Const aeRelTxtFile As String = "OutputRelationsSetup.txt"
Private Const aePrpTxtFile As String = "OutputPropertiesBuiltIn.txt"
Private Const aeFLkCtrFile As String = "OutputFieldLookupControlTypeList.txt"
Private Const aeSchemaFile As String = "OutputSchemaFile.txt"
Private Const aePrnterInfo As String = "OutputPrinterInfo.txt"
Private Const aeAppOptions As String = "OutputListOfAccessApplicationOptions.txt"
Private Const aeAppListPrp As String = "OutputListOfApplicationProperties.txt"
Private Const aeAppListCnt As String = "OutputListOfContainers.txt"
Private Const aeAppCmbrIds As String = "OutputListOfCommandBarIDs.txt"
Private Const aeAppListQAT As String = "OutputQAT"  ' Will be saved with file extension .exportedUI
Private Const aeCatalogObj As String = "OutputCatalogUserCreatedObjects.txt"
'

Private Sub Class_Initialize()
' Ref: http://www.cadalyst.com/cad/autocad/programming-with-class-part-1-5050
' Ref: http://www.bigresource.com/Tracker/Track-vb-cyJ1aJEyKj/
' Ref: http://stackoverflow.com/questions/1731052/is-there-a-way-to-overload-the-constructor-initialize-procedure-for-a-class-in

    On Error GoTo PROC_ERR
    
    If Application.VBE.ActiveVBProject.Name = "aegit" Then
        Dim dbs As DAO.Database
        Set dbs = CurrentDb()
        dbs.Properties("AppTitle").Value = Application.VBE.ActiveVBProject.Name & " " & aegit_expVERSION
        Application.RefreshTitleBar
        Set dbs = Nothing
    End If
    ' Provide a default value for the SourceFolder, ImportFolder and other properties
    aegitSourceFolder = "default"
    aegitSourceFolderBe = "default"
    aegitXMLfolder = "default"
    aestrBackEndDb1 = "default"             ' default for aegit is no back end database
    ReDim Preserve aegitDataXML(0 To 0)
    If Application.VBE.ActiveVBProject.Name = "aegit" Then
        aegitDataXML(0) = "aetlkpStates"
    End If
    aegitExportDataToXML = True
    aegitType.SourceFolder = "C:\ae\aegit\aerc\src\"
    aegitType.XMLfolder = "C:\ae\aegit\aerc\src\xml\"
    aeintLTN = 11           ' Set a minimum default
    aeintFNLen = 4          ' Set a minimum default
    aeintFTLen = 4          ' Set a minimum default
    aeintFDLen = 4          ' Set a minimum default

    With aegitExport
        .ExportAll = True
        .ExportCodeAndObjects = True
        .ExportModuleCodeOnly = True
        .ExportQAT = True
        .ExportCBID = False
    End With

    pExclude = True             ' Default setting is not to export associated aegit_exp files
    aegitFrontEndApp = True     ' Default if a front end app

    Debug.Print "Class_Initialize"
    Debug.Print , "Default for aegitSourceFolder = " & aegitSourceFolder
    Debug.Print , "Default for aegitType.SourceFolder = " & aegitType.SourceFolder
    Debug.Print , "Default for aegitType.XMLfolder = " & aegitType.XMLfolder
    Debug.Print , "Default for aegitSourceFolderBe = " & aegitSourceFolderBe
    Debug.Print , "Default for aestrBackEndDb1 = " & aestrBackEndDb1
    Debug.Print , "Default for aegitFrontEndApp = " & aegitFrontEndApp
    Debug.Print , "aeintLTN = " & aeintLTN
    Debug.Print , "aeintFNLen = " & aeintFNLen
    Debug.Print , "aeintFTLen = " & aeintFTLen
    Debug.Print , "aeintFSize = " & aeintFSize
    '
    Debug.Print , "aegitExport.ExportAll = " & aegitExport.ExportAll
    Debug.Print , "aegitExport.ExportCodeAndObjects = " & aegitExport.ExportCodeAndObjects
    Debug.Print , "aegitExport.ExportCodeOnly = " & aegitExport.ExportModuleCodeOnly
    Debug.Print , "aegitExport.ExportQAT = " & aegitExport.ExportQAT
    Debug.Print , "aegitExport.ExportCBID = " & aegitExport.ExportCBID
    defineMyExclusions
    Debug.Print , "pExclude = " & pExclude
    If aeExists("Forms", "_frmPersist") Then
        If Not IsLoaded("_frmPersist") Then
            DoCmd.OpenForm "_frmPersist", acNormal, , , acFormReadOnly, acHidden
        End If
        Debug.Print , "IsLoaded _frmPersist = " & IsLoaded("_frmPersist")
    End If
    'Stop

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class aegit_expClass"
    Resume PROC_EXIT

End Sub

Private Sub Class_Terminate()

    On Error GoTo 0
    Dim strFile As String
    strFile = aegitSourceFolder & "export.ini"
    If Dir$(strFile) <> vbNullString Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
    End If
    Debug.Print
    Debug.Print "Class_Terminate"
    If aeExists("Forms", "_frmPersist") Then
        If IsLoaded("_frmPersist") Then
            DoCmd.Close acForm, "_frmPersist", acSaveNo
        End If
        Debug.Print , "IsLoaded _frmPersist = " & IsLoaded("_frmPersist")
    End If
    Debug.Print , "aegit_exp VERSION: " & aegit_expVERSION
    Debug.Print , "aegit_exp VERSION_DATE: " & aegit_expVERSION_DATE
    '
'?    If Application.VBE.ActiveVBProject.Name <> "aegit" And _
'?            aestrBackEndDb1 <> "NONE" Then
'?        OpenAllDatabases True
'?    End If

End Sub

Public Property Get SourceFolder() As String
    On Error GoTo 0
    SourceFolder = aegitSourceFolder
End Property

Public Property Let SourceFolder(ByVal strSourceFolder As String)
    ' Ref: http://www.techrepublic.com/article/build-your-skills-using-class-modules-in-an-access-database-solution/5031814
    ' Ref: http://www.utteraccess.com/wiki/index.php/Classes
    On Error GoTo 0
    aegitSourceFolder = strSourceFolder
End Property

Public Property Get TextEncoding() As String
    On Error GoTo 0
    TextEncoding = aegitTextEncoding
End Property

Public Property Let TextEncoding(ByVal strTextEncoding As String)
    On Error GoTo 0
    aegitTextEncoding = strTextEncoding
End Property

Public Property Get SourceFolderBe() As String
    On Error GoTo 0
    SourceFolder = aegitSourceFolderBe
End Property

Public Property Let SourceFolderBe(ByVal strSourceFolderBe As String)
    On Error GoTo 0
    aegitSourceFolderBe = strSourceFolderBe
End Property

Public Property Get BackEndDb1() As String
    On Error GoTo 0
    BackEndDb1 = aestrBackEndDb1
End Property

Public Property Let BackEndDb1(ByVal strBackEndDbFullPath As String)
    On Error GoTo 0
    Debug.Print , "strBackEndDbFullPath = " & strBackEndDbFullPath, "Property Let BackEndDb1"
    aestrBackEndDb1 = strBackEndDbFullPath
End Property

Public Property Get XMLfolder() As String
    On Error GoTo 0
    XMLfolder = aegitXMLfolder
End Property

Public Property Let XMLfolder(ByVal strXMLfolder As String)
    On Error GoTo 0
    aegitXMLfolder = strXMLfolder
End Property

Public Property Let ExportQAT(ByVal blnExportQAT As Boolean)
    On Error GoTo 0
    If blnExportQAT Then
        aegitExport.ExportQAT = True
    Else
        aegitExport.ExportQAT = False
    End If
End Property

Public Property Let ExportCBID(ByVal blnExportCBID As Boolean)
    On Error GoTo 0
    If blnExportCBID Then
        aegitExport.ExportCBID = True
    Else
        aegitExport.ExportCBID = False
    End If
End Property

Public Property Let TablesExportToXML(ByVal varTablesArray As Variant)
' Ref: http://stackoverflow.com/questions/2265349/how-can-i-use-an-optional-array-argument-in-a-vba-procedure
    On Error GoTo PROC_ERR

    Debug.Print , "LBound(varTablesArray) = " & LBound(varTablesArray), "varTablesArray(0) = " & varTablesArray(0)
    Debug.Print , "UBound(varTablesArray) = " & UBound(varTablesArray)
    If UBound(varTablesArray) > 0 Then
        Debug.Print , "varTablesArray(1) = " & varTablesArray(1)
    End If
    ReDim Preserve aegitDataXML(0 To UBound(varTablesArray))
    aegitDataXML = varTablesArray
    Debug.Print , "aegitDataXML(0) = " & aegitDataXML(0)
    If UBound(varTablesArray) > 0 Then
        Debug.Print , "aegitDataXML(1) = " & aegitDataXML(1)
    End If

PROC_EXIT:
    Exit Property

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TablesExportToXML"
            Stop
    End Select

End Property

Public Property Get DocumentTheDatabase(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        Debug.Print "Get DocumentTheDatabase"
        Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentTheDatabase = aeDocumentTheDatabase()
    Else
        Debug.Print "Get DocumentTheDatabase"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTheDatabase = aeDocumentTheDatabase(varDebug)
    End If
End Property

Public Property Get Exists(ByVal strAccObjType As String, _
                        ByVal strAccObjName As String, _
                        Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        Debug.Print "Get Exists"
        Debug.Print , "varDebug IS missing so no parameter is passed to aeExists"
        Debug.Print , "DEBUGGING IS OFF"
        Exists = aeExists(strAccObjType, strAccObjName)
    Else
        Debug.Print "Get Exists"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeExists"
        Debug.Print , "DEBUGGING TURNED ON"
        Exists = aeExists(strAccObjType, strAccObjName, varDebug)
    End If
End Property

Public Property Get GetReferences(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        Debug.Print "Get GetReferences"
        Debug.Print , "varDebug IS missing so no parameter is passed to aeGetReferences"
        Debug.Print , "DEBUGGING IS OFF"
        GetReferences = aeGetReferences()
    Else
        Debug.Print "Get GetReferences"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeGetReferences"
        Debug.Print , "DEBUGGING TURNED ON"
        GetReferences = aeGetReferences(varDebug)
    End If
End Property

Public Property Get DocumentRelations(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        Debug.Print "Get DocumentRelations"
        Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentRelations"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentRelations = aeDocumentRelations()
    Else
        Debug.Print "Get DocumentRelations"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentRelations"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentRelations = aeDocumentRelations(varDebug)
    End If
End Property

Public Property Get DocumentTables(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        Debug.Print "Get DocumentTables"
        Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTables"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentTables = aeDocumentTables()
    Else
        Debug.Print "Get DocumentTables"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTables"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTables = aeDocumentTables(varDebug)
    End If
End Property

Public Property Get DocumentTablesXML(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        Debug.Print "Get DocumentTablesXML"
        Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTablesXML"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentTablesXML = aeDocumentTablesXML()
    Else
        Debug.Print "Get DocumentTablesXML"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTablesXML"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTablesXML = aeDocumentTablesXML(varDebug)
    End If
End Property

Public Property Get CompactAndRepair(Optional ByVal varTrueFalse As Variant) As Boolean
' Automation for Compact and Repair

    On Error GoTo 0
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

Public Property Get ExcludeFiles(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    ExcludeFiles = pExclude
    Debug.Print , "ExcludeFiles = " & pExclude
    If IsMissing(varDebug) Then
        Debug.Print "Get ExcludeFiles"
        Debug.Print , "varDebug IS missing so no parameter is passed to ExcludeFiles"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print "Get ExcludeFiles"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to ExcludeFiles"
        Debug.Print , "DEBUGGING TURNED ON"
    End If
End Property

Public Property Let ExcludeFiles(Optional ByVal varDebug As Variant, ByVal blnExclude As Boolean)
    On Error GoTo 0
    pExclude = blnExclude
    Debug.Print , "Let ExcludeFiles = " & pExclude
End Property

Private Function IsLoaded(ByVal strFormName As String) As Boolean
 ' Returns True if the specified form is open in Form view or Datasheet view.
    
    Const conObjStateClosed = 0
    Const conDesignView = 0
    
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If
    
End Function

Private Sub OutputTableProperties(Optional ByVal varDebug As Variant)
' Ref: http://bytes.com/topic/access/answers/709190-how-export-table-structure-including-description

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim prp As DAO.Property
    Dim fldprp As DAO.Property
    Dim fld As DAO.Field

    If IsMissing(varDebug) Then
        Debug.Print "OutputTableProperties"
        Debug.Print , "varDebug IS missing so no parameter is passed to OutputTableProperties"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print "OutputTableProperties"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to OutputTableProperties"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    If Not IsMissing(varDebug) Then Debug.Print "aegitSourceFolder=" & aegitSourceFolder

    If aegitSourceFolder = "default" Then
        aegitSourceFolder = aegitType.SourceFolder
    End If

    Set dbs = CurrentDb()
    For Each tdf In CurrentDb.TableDefs

        If Not (Left$(tdf.Name, 4) = "MSys" _
                Or Left$(tdf.Name, 4) = "~TMP" _
                Or Left$(tdf.Name, 3) = "zzz") Then

            Open aegitSourceFolder & "Properties_" & tdf.Name & ".txt" For Output As #1

            On Error Resume Next
            For Each prp In tdf.Properties
                ' Ignore DateCreated, LastUpdated, GUID and NameMap output
                If prp.Name = "DateCreated" Then
                    Print #1, "|-- " & prp.Name & " >> " & "DateCreated"
                ElseIf prp.Name = "LastUpdated" Then
                    Print #1, "|-- " & prp.Name & " >> " & "LastUpdated"
                ElseIf prp.Name = "GUID" Then
                    Print #1, "|-- " & prp.Name & " >> " & "GUID"
                ElseIf prp.Name = "NameMap" Then
                    Print #1, "|-- " & prp.Name & " >> " & "NameMap"
                Else
                    Print #1, "|-- " & prp.Name & " >> " & prp.Value
                End If
            Next prp
            Print #1, "---------------------------------------------------------"
            For Each fld In tdf.Fields
                Print #1, "|-- " & fld.Name & " (Field in " & tdf.Name & ")"
                For Each fldprp In fld.Properties
                    If fldprp.Name = "GUID" Then
                        Print #1, "|------ " & fldprp.Name & " >> " & "GUID"
                    Else
                        Print #1, "|------ " & fldprp.Name & " >> " & fldprp.Value
                    End If
                Next
            Next
            Close #1
        End If
NextTdf:
    Next tdf
End Sub

Private Function LinkedTable(strTblName) As Boolean

    On Error GoTo PROC_ERR

    ' Linked table connection string is > 0
    If Len(CurrentDb.TableDefs(strTblName).Connect) > 0 Then
        ' Linked table exists, but is the link valid?
        ' The next line of code will generate Errors 3011 or 3024 if it isn't
        CurrentDb.TableDefs(strTblName).RefreshLink
        'If you get to this point, you have a valid, Linked Table
        LinkedTable = True
        'Debug.Print "LinkedTable = True"
    Else
        ' Local table connect string length = 0
        ' MsgBox "[" & strTblName & "] is a Non-Linked Table", vbInformation, "Internal Table"
        LinkedTable = False
        'Debug.Print "LinkedTable = False"
    End If

PROC_EXIT:
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 3265
            MsgBox "[" & strTblName & "] does not exist as either an Internal or Linked Table", _
                vbCritical, "Table Missing"
        Case 3011, 3024     'Linked Table does not exist or DB Path not valid
            MsgBox "[" & strTblName & "] is not a valid, Linked Table", vbCritical, "Link Not Valid"
        Case Else
            MsgBox Err.Description & Err.Number, vbExclamation, "LinkedTable Error"
    End Select
    Resume PROC_EXIT
End Function

Private Sub OpenAllDatabases(blnInit As Boolean)
' Open a handle to all databases and keep it open during the entire time the application runs.
' Params : blnInit   TRUE to initialize (call when application starts)
'                    FALSE to close (call when application ends)
' Ref    : http://stackoverflow.com/questions/29838317/issue-when-using-a-dao-handle-when-the-database-closes-unexpectedly

    Dim intX As Integer
    Dim strName As String
    Dim strMsg As String

    If aestrBackEndDb1 = "NONE" Then
        Exit Sub
    Else
        Debug.Print , "aestrBackEndDb1 = " & aestrBackEndDb1, "OpenAllDatabases"
    End If

    ' Maximum number of back end databases to link
    Const cintMaxDatabases As Integer = 1

    ' List of databases kept in a static array so we can close them later
    Static dbsOpen() As DAO.Database
 
    'MsgBox "aestrBackEndDb1 = " & aestrBackEndDb1, vbInformation, "OpenAllDatabases"
    If blnInit Then
        ReDim dbsOpen(1 To cintMaxDatabases)
        For intX = 1 To cintMaxDatabases
            ' Specify your back end databases
            Select Case intX
                Case 1:
                    strName = aestrBackEndDb1
                Case 2:
                    strName = "H:\folder\Backend2.mdb"
            End Select
        strMsg = ""
        'Debug.Print , "strName = " & strName, "OpenAllDatabases"

        On Error Resume Next
        ' Ref: https://support.microsoft.com/en-us/kb/209953
        ' If you use a Connect argument and you do not provide the Options and Read-Only arguments, you receive run-time error 3031: Not a valid password.
        ' Ref: https://msdn.microsoft.com/en-us/library/office/ff835343.aspx
        Set dbsOpen(intX) = OpenDatabase(strName) ' Shared, Read Only
        ' Example for password protected back end requires use of Let property for aestrPassword
        'Set dbsOpen(intX) = OpenDatabase(strName, False, True, "MS Access;pwd=" & aestrPassword) ' Shared, Read Only
        If Err.Number > 0 Then
            strMsg = "Trouble opening database: " & strName & vbCrLf & _
                    "Make sure the drive is available." & vbCrLf & _
                    "Error: " & Err.Description & " (" & Err.Number & ")"
        End If

        On Error GoTo 0
        If strMsg <> "" Then
            MsgBox strMsg & vbCrLf & "strName = " & strName, vbExclamation, "OpenAllDatabases"
            Exit For
        End If
        Next intX
    Else
        On Error Resume Next
        For intX = 1 To cintMaxDatabases
            dbsOpen(intX).Close
        Next intX
    End If

End Sub

Private Function Delay(ByVal mSecs As Long) As Boolean
    On Error GoTo 0
    Sleep mSecs ' delay milli seconds
End Function

Private Function defineMyExclusions() As myExclusions
    On Error GoTo 0
    myExclude.exclude1 = EXCLUDE_1
    myExclude.exclude2 = EXCLUDE_2
    myExclude.exclude3 = EXCLUDE_3
End Function

Private Function fExclude(strName As String) As Boolean
    On Error GoTo 0
    fExclude = False
    'Debug.Print "1: fExclude", strName, "myExclude.exclude1 = " & myExclude.exclude1
    If strName = myExclude.exclude1 Then
        fExclude = True
        Exit Function
    End If
    If strName = myExclude.exclude2 Then
        fExclude = True
        Exit Function
    End If
    If strName = myExclude.exclude3 Then
        fExclude = True
        Exit Function
    End If
End Function

Private Function FixHeaderXML(ByVal strPathFileName As String) As Boolean

    On Error GoTo PROC_ERR

    Dim fName As String
    Dim fName2 As String
    Dim fnr As Integer
    Dim fnr2 As Integer
    Dim tstring As String * 1
    Dim blnDone As Boolean

    FixHeaderXML = False
    blnDone = False

    fName = strPathFileName
    fName2 = strPathFileName & ".fixed.xml"
    Debug.Print fName, fName2

    fnr = FreeFile()
    Open fName For Binary Access Read As #fnr
    Get #fnr, , tstring

    fnr2 = FreeFile()
    Open fName2 For Binary Lock Read Write As #fnr2

    While Not EOF(fnr)
        Get #fnr, , tstring
        Debug.Print Asc(tstring)
        If Not blnDone Then
                If Asc(tstring) <> 62 Then          ' ">"
                    Debug.Print Asc(tstring)
                Else
                    blnDone = True
                End If
        Else
            If Asc(tstring) <> 0 Then Put #fnr2, , tstring
        End If
    Wend
    'Stop

PROC_EXIT:
    Close #fnr
    Close #fnr2
    FixHeaderXML = True
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 9
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FixHeaderXML of Class aegit_expClass" & _
                    vbCrLf & "FixHeaderXML Entry strPathFileName=" & strPathFileName, vbCritical, "FixHeaderXML ERROR=9"
            'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FixHeaderXML of Class aegit_expClass"
            Resume Next
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FixHeaderXML of Class aegit_expClass"
            Resume Next
    End Select

End Function

Private Function NoBOM(ByVal strFileName As String) As Boolean
' Ref: http://www.experts-exchange.com/Programming/Languages/Q_27478996.html
' Use the same file name for input and output

    On Error GoTo PROC_ERR

    ' Define needed constants
    Const ForReading = 1
    Const ForWriting = 2
    Const TriStateUseDefault = -2
    Const adTypeText = 2
    Dim strContent As String

    NoBOM = False
    ' Convert UTF-16 file to ANSI file
    Dim objStreamFile As Object
    Set objStreamFile = CreateObject("Adodb.Stream")
    With objStreamFile
        .Charset = "UTF-8"
        .Type = adTypeText
        .Open
        .LoadFromFile strFileName
        strContent = .ReadText
        .Close
    End With
    Set objStreamFile = Nothing
    Kill strFileName
    'Stop

    DoEvents

    ' Write out after "conversion"
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFile As Object
    'Debug.Print , "strFileName = " & strFileName
    Set objFile = objFSO.OpenTextFile(strFileName, ForWriting, True)
    strContent = Right$(strContent, Len(strContent) - 2)
    objFile.Write strContent
    objFile.Close

    Set objFile = Nothing
    NoBOM = True

PROC_EXIT:
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 9999
            Resume PROC_EXIT
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure NoBOM of Class aegit_expClass"
            Resume PROC_EXIT
    End Select
    Resume PROC_EXIT

End Function

Private Function aeReadWriteStream(ByVal strPathFileName As String) As Boolean

    On Error GoTo PROC_ERR

    Dim fName As String
    Dim fName2 As String
    Dim fnr As Integer
    Dim fnr2 As Integer
    Dim tstring As String * 1

    aeReadWriteStream = False

    ' If the file has no Byte Order Mark (BOM)
    ' Ref: http://msdn.microsoft.com/en-us/library/windows/desktop/dd374101%28v=vs.85%29.aspx
    ' then do nothing
    fName = strPathFileName
    fName2 = strPathFileName & ".clean.txt"

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
    Open fName2 For Binary Lock Read Write As #fnr2

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
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 9
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadWriteStream of Class aegit_expClass" & _
                    vbCrLf & "aeReadWriteStream Entry strPathFileName=" & strPathFileName, vbCritical, "aeReadWriteStream ERROR=9"
            'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadWriteStream of Class aegit_expClass"
            Resume Next
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadWriteStream of Class aegit_expClass"
            Resume Next
    End Select

End Function

Private Function RecordsetUpdatable(ByVal strSQL As String) As Boolean
' Ref: http://msdn.microsoft.com/en-us/library/office/ff193796(v=office.15).aspx

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim intPosition As Integer

    On Error GoTo PROC_ERR

    ' Initialize the function's return value to True.
    RecordsetUpdatable = True

    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(strSQL, dbOpenDynaset)

    ' If the entire dynaset isn't updatable, return False.
    If rst.Updatable = False Then
        RecordsetUpdatable = False
    Else
        ' If the dynaset is updatable, check if all fields in the
        ' dynaset are updatable. If one of the fields isn't updatable,
        ' return False.
        For intPosition = 0 To rst.Fields.Count - 1
            If rst.Fields(intPosition).DataUpdatable = False Then
                RecordsetUpdatable = False
                Exit For
            End If
        Next intPosition
    End If

PROC_EXIT:
    rst.Close
    dbs.Close
    Set rst = Nothing
    Set dbs = Nothing
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RecordsetUpdatable of Class aegit_expClass"
    Resume Next

End Function

Private Function IsQryHidden(ByVal strQueryName As String) As Boolean
    On Error GoTo 0
    If IsNull(strQueryName) Or strQueryName = vbNullString Then
        IsQryHidden = False
        'Debug.Print "IsQryHidden Null Test", strQueryName, IsQryHidden
    Else
        IsQryHidden = GetHiddenAttribute(acQuery, strQueryName)
        'Debug.Print "IsQryHidden Attribute Test", strQueryName, IsQryHidden
    End If
End Function

Private Sub OutputTheQAT(ByVal strTheFile As String, Optional ByVal varDebug As Variant)
' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=52635
' Set focus to the Access window

    On Error GoTo PROC_ERR

    Dim strAppTitle As String
    Dim lngHwnd As Long

    DoCmd.RunCommand acCmdAppMaximize
    Delay 500
    strAppTitle = CurrentDb.Properties("AppTitle")
    AppActivate strAppTitle
    lngHwnd = FindWindow(vbNullString, strAppTitle)
    If Not IsMissing(varDebug) Then
        Debug.Print "strAppTitle=" & strAppTitle
        Debug.Print "lngHwnd=" & lngHwnd
    End If

    ' FIXME - NOTE: Flaky and unreliable SendKeys... needs some work

    apiSetActiveWindow lngHwnd
    ' Ref: http://www.jpsoftwaretech.com/vba/shell-scripting-vba-windows-script-host-object-model/
    ' Ref: http://www.computerperformance.co.uk/ezine/ezine26.htm
    'Pause 5
    'Stop
    Dim wsc As Object
    Set wsc = CreateObject("WScript.Shell")
    Delay 500
    wsc.SendKeys "{F10}FT"
    wsc.SendKeys "{HOME}CCC"
    Delay 250
    wsc.SendKeys "%P{TAB}{ENTER}"
    Delay 250
    wsc.SendKeys "{BACKSPACE}"
    Delay 500
    wsc.SendKeys aegitSourceFolder & strTheFile
    Delay 250
    wsc.SendKeys "%S"
    wsc.SendKeys "{ESC}"
    Pause 1

    ' NOTE: the QAT Ribbon XML as <mso:cmd app="Access" dt="0" /> at the start
    ' and it messes parsing as standard XML. Remove it, rename the file as .xml,
    ' save to the xml folder and prettify for reading.

    FixHeaderXML (aegitSourceFolder & strTheFile & ".exportedUI")
    'Stop

    If Not IsMissing(varDebug) Then
        PrettyXML aegitSourceFolder & strTheFile & ".exportedUI.fixed.xml", varDebug
    Else
        PrettyXML aegitSourceFolder & strTheFile & ".exportedUI.fixed.xml"
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    If Err = 3270 Then ' Property not found
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputTheQAT of Class aegit_expClass" & vbCrLf & _
            "No AppTitle found so QAT is not being exported!", vbInformation, "OutputTheQAT"
            Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputTheQAT of Class aegit_expClass"
        Resume Next
    End If

End Sub

Private Sub OutputListOfAllHiddenQueries(Optional ByVal varDebug As Variant)
' Ref: http://www.pcreview.co.uk/forums/runtime-error-7874-a-t2922352.html

    Dim strTheSQL As String
    Dim varResult As Variant
    Dim intHidden As Integer
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset

    Set dbs = CurrentDb()
    intHidden = 0

    On Error GoTo PROC_ERR

    Const strTempTable As String = "zzzTmpTblQueries"
    ' NOTE: Use zzz* for the table name so that it will be ignored by aegit code export if it exists
    ' MSysObjects list of types - Ref: http://allenbrowne.com/func-DDL.html - Query = 5
    Const strSQL As String = "SELECT m.Name INTO " & strTempTable & " " & vbCrLf & _
                                "FROM MSysObjects AS m " & vbCrLf & _
                                "WHERE (((m.Name) Not ALike ""~%"") AND ((m.Type)=5)) " & vbCrLf & _
                                "ORDER BY m.Name;"
    If Not IsMissing(varDebug) Then Debug.Print strSQL
    If Not IsMissing(varDebug) And _
                Application.VBE.ActiveVBProject.Name = "aegit" Then
        Debug.Print "IsQryHidden('qpt_Dummy') = " & IsQryHidden("qpt_Dummy")
        Debug.Print "IsQryHidden('qry_HiddenDummy') = " & IsQryHidden("qry_HiddenDummy")
    End If
    'Stop

    DoCmd.SetWarnings False
    ' Use RunSQL for action queries - Insert list of db queries into a temp table
    DoCmd.RunSQL strSQL

e3167e3011e3078:
    Set rst = dbs.OpenRecordset(strTempTable, dbOpenTable)

    With rst
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While Not rst.EOF
                varResult = !Name
                If Not IsMissing(varDebug) Then
                    Debug.Print ">", !Name.Value, IsQryHidden(!Name)
                End If
                If IsQryHidden(!Name) Then
                    intHidden = intHidden + 1
                    .MoveNext
                Else
                    .Delete
                    ' Ref: https://msdn.microsoft.com/en-us/library/bb243799%28v=office.12%29.aspx
                    ' When you use the Delete method, the Microsoft Access database engine immediately deletes the current record
                    ' without any warning or prompting. Deleting a record does not automatically cause the next record to become the current record;
                    ' to move to the next record you must use the MoveNext method. However, keep in mind that after you have moved off the deleted record, you cannot move back to it.
                    .MoveNext
                End If
            Loop
        Else
            Debug.Print "No records!"
        End If
    End With

    If Not IsMissing(varDebug) Then
        Debug.Print "The number of hidden queries in the database is: " & intHidden, "rst.RecordCount = " & rst.RecordCount     ', "DCount(""Name"", strTempTable) = " & DCount("Name", strTempTable)
    End If

    rst.Close
    dbs.Close

    DoCmd.TransferText acExportDelim, vbNullString, strTempTable, aestrSourceLocation & "OutputListOfAllHiddenQueries.txt", False
'Stop
    CurrentDb.Execute "DROP TABLE " & strTempTable
    DoCmd.SetWarnings True

PROC_EXIT:
    Set rst = Nothing
    Set dbs = Nothing
    Exit Sub

PROC_ERR:
    If Err = 3167 Then          ' Record is deleted
        'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfAllHiddenQueries of Class aegit_expClass"
        Resume e3167e3011e3078
    ElseIf Err = 3011 Then
        'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfAllHiddenQueries of Class aegit_expClass"
        Resume e3167e3011e3078
    ElseIf Err = 3078 Then
        'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfAllHiddenQueries of Class aegit_expClass"
        Resume e3167e3011e3078
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfAllHiddenQueries of Class aegit_expClass"
    End If
    Resume PROC_EXIT

End Sub

Private Sub OutputListOfAccessApplicationOptions(Optional ByVal varDebug As Variant)

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

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim fle As Integer

    If IsMissing(varDebug) Then
        Debug.Print "OutputListOfAccessApplicationOptions"
        Debug.Print , "varDebug IS missing so no parameter is passed to OutputListOfAccessApplicationOptions"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print "OutputListOfAccessApplicationOptions"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to OutputListOfAccessApplicationOptions"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    ' Test for relative path
    Dim strTestPath As String
    strTestPath = aegitSourceFolder
    If Left(aegitSourceFolder, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid(aegitSourceFolder, 2, Len(aegitSourceFolder) - 1)
        aegitSourceFolder = strTestPath
        'Debug.Print , "aegitSourceFolder = " & aegitSourceFolder, "OutputListOfApplicationOptions"
        'Stop
    End If

    If Not IsMissing(varDebug) Then Debug.Print "aegitSourceFolder=" & aegitSourceFolder

    If aegitSourceFolder = "default" Then
        aegitSourceFolder = aegitType.SourceFolder
    End If

    fle = FreeFile()
    Open aegitSourceFolder & "\" & aeAppOptions For Output As #fle

    Print #fle, ">>>Standard Options"
    ' 2000 The following options are equivalent to the standard startup options found in the Startup Options dialog box.
    Print #fle, , "2000", "AppTitle              ", dbs.Properties!AppTitle                     ' String  The title of an application, as displayed in the title bar.
    Print #fle, , "2000", "AppIcon               ", dbs.Properties!AppIcon                      ' String  The file name and path of an application's icon.
    Print #fle, , "2000", "StartupMenuBar        ", dbs.Properties!StartUpMenuBar               ' String  Sets the default menu bar for the application.
    Print #fle, , "2000", "AllowFullMenus        ", dbs.Properties!AllowFullMenus               ' True/False  Determines if the built-in Access menu bars are displayed.
    Print #fle, , "2000", "AllowShortcutMenus    ", dbs.Properties!AllowShortcutMenus           ' True/False  Determines if the built-in Access shortcut menus are displayed.
    Print #fle, , "2000", "StartupForm           ", dbs.Properties!StartUpForm                  ' String  Sets the form or data page to show when the application is first opened.
    Print #fle, , "2000", "StartupShowDBWindow   ", dbs.Properties!StartUpShowDBWindow          ' True/False  Determines if the database window is displayed when the application is first opened.
    Print #fle, , "2000", "StartupShowStatusBar  ", dbs.Properties!StartUpShowStatusBar         ' True/False  Determines if the status bar is displayed.
    Print #fle, , "2000", "StartupShortcutMenuBar", dbs.Properties!StartUpShortcutMenuBar       ' String  Sets the shortcut menu bar to be used in all forms and reports.
    Print #fle, , "2000", "AllowBuiltInToolbars  ", dbs.Properties!AllowBuiltInToolbars         ' True/False  Determines if the built-in Access toolbars are displayed.
    Print #fle, , "2000", "AllowToolbarChanges   ", dbs.Properties!AllowToolbarChanges          ' True/False  Determined if toolbar changes can be made.
    Print #fle, ">>>Advanced Option"
    Print #fle, , "2000", "AllowSpecialKeys      ", dbs.Properties!AllowSpecialKeys             ' Option (True/False value) determines if the use of special keys is permitted. It is equivalent to the advanced startup option found in the Startup Options dialog box.
    Print #fle, ">>>Extra Options"
    ' The following options are not available from the Startup Options dialog box or any other Access user interface component, they are only available in programming code.
    Print #fle, , "2000", "AllowBypassKey        ", dbs.Properties!AllowBypassKey               ' True/False  Determines if the SHIFT key can be used to bypass the application load process.
    Print #fle, , "2000", "AllowBreakIntoCode    ", dbs.Properties!AllowBreakIntoCode           ' True/False  Determines if the CTRL+BREAK key combination can be used to stop code from running.
    Print #fle, , "2000", "HijriCalendar         ", dbs.Properties!HijriCalendar                ' True/False  Applies only to Arabic countries; determines if the application uses Hijri or Gregorian dates.
    Print #fle, ">>>View Tab"
    Print #fle, , "XP, 2003", "Show Status Bar                 ", Application.GetOption("Show Status Bar")                    ' Show, Status bar
    Print #fle, , "XP, 2003", "Show Startup Dialog Box         ", Application.GetOption("Show Startup Dialog Box")            ' Show, Startup Task Pane
    Print #fle, , "XP, 2003", "Show New Object Shortcuts       ", Application.GetOption("Show New Object Shortcuts")          ' Show, New object shortcuts
    Print #fle, , "XP, 2003", "Show Hidden Objects             ", Application.GetOption("Show Hidden Objects")                ' Show, Hidden objects
    Print #fle, , "XP, 2003", "Show System Objects             ", Application.GetOption("Show System Objects")                ' Show, System objects
    Print #fle, , "XP, 2003", "ShowWindowsInTaskbar            ", Application.GetOption("ShowWindowsInTaskbar")               ' Show, Windows in Taskbar
    Print #fle, , "XP, 2003", "Show Macro Names Column         ", Application.GetOption("Show Macro Names Column")            ' Show in Macro Design, Names column
    Print #fle, , "XP, 2003", "Show Conditions Column          ", Application.GetOption("Show Conditions Column")             ' Show in Macro Design, Conditions column
    Print #fle, , "XP, 2003", "Database Explorer Click Behavior", Application.GetOption("Database Explorer Click Behavior")   ' Click options in database window
    Print #fle, ">>>General Tab"
    Print #fle, , "XP, 2003", "Left Margin                 ", Application.GetOption("Left Margin")                                            ' Print margins, Left margin
    Print #fle, , "XP, 2003", "Right Margin                ", Application.GetOption("Right Margin")                                           ' Print margins, Right margin
    Print #fle, , "XP, 2003", "Top Margin                  ", Application.GetOption("Top Margin")                                             ' Print margins, Top margin
    Print #fle, , "XP, 2003", "Bottom Margin               ", Application.GetOption("Bottom Margin")                                          ' Print margins, Bottom margin
    Print #fle, , "XP, 2003", "Four-Digit Year Formatting  ", Application.GetOption("Four-Digit Year Formatting")                             ' Use four-year digit year formatting, This database
    Print #fle, , "XP, 2003", "Four-Digit Year Formatting All Databases", Application.GetOption("Four-Digit Year Formatting All Databases")   ' Use four-year digit year formatting, All databases  Four-Digit Year Formatting All Databases
    Print #fle, , "XP, 2003", "Track Name AutoCorrect Info ", Application.GetOption("Track Name AutoCorrect Info")                            ' Name AutoCorrect, Track name AutoCorrect info
    Print #fle, , "XP, 2003", "Perform Name AutoCorrect    ", Application.GetOption("Perform Name AutoCorrect")                               ' Name AutoCorrect, Perform name AutoCorrect
    Print #fle, , "XP, 2003", "Log Name AutoCorrect Changes", Application.GetOption("Log Name AutoCorrect Changes")                           ' Name AutoCorrect, Log name AutoCorrect changes
    Print #fle, , "XP, 2003", "Enable MRU File List        ", Application.GetOption("Enable MRU File List")                                   ' Recently used file list
    Print #fle, , "XP, 2003", "Size of MRU File List       ", "Not Tracked"     'Application.GetOption("Size of MRU File List")                                  ' Recently used file list, (number of files)
    Print #fle, , "XP, 2003", "Provide Feedback with Sound ", Application.GetOption("Provide Feedback with Sound")                            ' Provide feedback with sound
    Print #fle, , "XP, 2003", "Auto Compact                ", Application.GetOption("Auto Compact")                                           ' Compact on Close
    Print #fle, , "XP, 2003", "New Database Sort Order     ", Application.GetOption("New Database Sort Order")                                ' New database sort order
    Print #fle, , "XP, 2003", "Remove Personal Information ", Application.GetOption("Remove Personal Information")                            ' Remove personal information from this file
    Print #fle, , "XP, 2003", "Default Database Directory  ", Application.GetOption("Default Database Directory")                             ' Default database folder
    Print #fle, ">>>Edit/Find Tab"
    Print #fle, , "XP, 2003", "Default Find/Replace Behavior", Application.GetOption("Default Find/Replace Behavior")       ' Default find/replace behavior
    Print #fle, , "XP, 2003", "Confirm Record Changes       ", Application.GetOption("Confirm Record Changes")              ' Confirm, Record changes
    Print #fle, , "XP, 2003", "Confirm Document Deletions   ", Application.GetOption("Confirm Document Deletions")          ' Confirm, Document deletions
    Print #fle, , "XP, 2003", "Confirm Action Queries       ", Application.GetOption("Confirm Action Queries")              ' Confirm, Action queries
    Print #fle, , "XP, 2003", "Show Values in Indexed       ", Application.GetOption("Show Values in Indexed")              ' Show list of values in, Local indexed fields
    Print #fle, , "XP, 2003", "Show Values in Non-Indexed   ", Application.GetOption("Show Values in Non-Indexed")          ' Show list of values in, Local nonindexed fields
    Print #fle, , "XP, 2003", "Show Values in Remote        ", Application.GetOption("Show Values in Remote")               ' Show list of values in, ODBC fields
    Print #fle, , "XP, 2003", "Show Values in Snapshot      ", Application.GetOption("Show Values in Snapshot")             ' Show list of values in, Records in local snapshot
    Print #fle, , "XP, 2003", "Show Values in Server        ", Application.GetOption("Show Values in Server")               ' Show list of values in, Records at server
    Print #fle, , "XP, 2003", "Show Values Limit            ", Application.GetOption("Show Values Limit")                   ' Don't display lists where more than this number of records read
    Print #fle, ">>>Datasheet Tab"
    Print #fle, , "XP, 2003", "Default Font Color           ", Application.GetOption("Default Font Color")                  ' Default colors, Font
    Print #fle, , "XP, 2003", "Default Background Color     ", Application.GetOption("Default Background Color")            ' Default colors, Background
    Print #fle, , "XP, 2003", "Default Gridlines Color      ", Application.GetOption("Default Gridlines Color")             ' Default colors, Gridlines
    Print #fle, , "XP, 2003", "Default Gridlines Horizontal ", Application.GetOption("Default Gridlines Horizontal")        ' Default gridlines showing, Horizontal
    Print #fle, , "XP, 2003", "Default Gridlines Vertical   ", Application.GetOption("Default Gridlines Vertical")          ' Default gridlines showing, Vertical
    Print #fle, , "XP, 2003", "Default Column Width         ", Application.GetOption("Default Column Width")                ' Default column width
    Print #fle, , "XP, 2003", "Default Font Name            ", Application.GetOption("Default Font Name")                   ' Default font, Font
    Print #fle, , "XP, 2003", "Default Font Weight          ", Application.GetOption("Default Font Weight")                 ' Default font, Weight
    Print #fle, , "XP, 2003", "Default Font Size            ", Application.GetOption("Default Font Size")                   ' Default font, Size
    Print #fle, , "XP, 2003", "Default Font Underline       ", Application.GetOption("Default Font Underline")              ' Default font, Underline
    Print #fle, , "XP, 2003", "Default Font Italic          ", Application.GetOption("Default Font Italic")                 ' Default font, Italic
    Print #fle, , "XP, 2003", "Default Cell Effect          ", Application.GetOption("Default Cell Effect")                 ' Default cell effect
    Print #fle, , "XP, 2003", "Show Animations              ", Application.GetOption("Show Animations")                     ' Show animations
    Print #fle, , "    2003", "Show Smart Tags on Datasheets", Application.GetOption("Show Smart Tags on Datasheets")       ' Show Smart Tags on Datasheets
    Print #fle, ">>>Keyboard Tab"
    Print #fle, , "XP, 2003", "Move After Enter                ", Application.GetOption("Move After Enter")                   ' Move after enter
    Print #fle, , "XP, 2003", "Behavior Entering Field         ", Application.GetOption("Behavior Entering Field")            ' Behavior entering field
    Print #fle, , "XP, 2003", "Arrow Key Behavior              ", Application.GetOption("Arrow Key Behavior")                 ' Arrow key behavior
    Print #fle, , "XP, 2003", "Cursor Stops at First/Last Field", Application.GetOption("Cursor Stops at First/Last Field")   ' Cursor stops at first/last field
    Print #fle, , "XP, 2003", "Ime Autocommit                  ", Application.GetOption("Ime Autocommit")                     ' Auto commit
    Print #fle, , "XP, 2003", "Datasheet Ime Control           ", Application.GetOption("Datasheet Ime Control")              ' Datasheet IME control
    Print #fle, ">>>Tables/Queries Tab"
    Print #fle, , "XP, 2003", "Default Text Field Size             ", Application.GetOption("Default Text Field Size")              ' Table design, Default field sizes - Text
    Print #fle, , "XP, 2003", "Default Number Field Size           ", Application.GetOption("Default Number Field Size")            ' Table design, Default field sizes - Number
    Print #fle, , "XP, 2003", "Default Field Type                  ", Application.GetOption("Default Field Type")                   ' Table design, Default field type
    Print #fle, , "XP, 2003", "AutoIndex on Import/Create          ", Application.GetOption("AutoIndex on Import/Create")           ' Table design, AutoIndex on Import/Create
    Print #fle, , "XP, 2003", "Show Table Names                    ", Application.GetOption("Show Table Names")                     ' Query design, Show table names
    Print #fle, , "XP, 2003", "Output All Fields                   ", Application.GetOption("Output All Fields")                    ' Query design, Output all fields
    Print #fle, , "XP, 2003", "Enable AutoJoin                     ", Application.GetOption("Enable AutoJoin")                      ' Query design, Enable AutoJoin
    Print #fle, , "XP, 2003", "Run Permissions                     ", Application.GetOption("Run Permissions")                      ' Query design, Run permissions
    Print #fle, , "XP, 2003", "ANSI Query Mode                     ", Application.GetOption("ANSI Query Mode")                      ' Query design, SQL Server Compatible Syntax (ANSI 92) - This database
    Print #fle, , "XP, 2003", "ANSI Query Mode Default             ", Application.GetOption("ANSI Query Mode Default")              ' Query design, SQL Server Compatible Syntax (ANSI 92) - Default for new databases
    Print #fle, , "    2003", "Query Design Font Name              ", Application.GetOption("Query Design Font Name")               ' Query design, Query design font, Font
    Print #fle, , "    2003", "Query Design Font Size              ", Application.GetOption("Query Design Font Size")               ' Query design, Query design font, Size
    Print #fle, , "    2003", "Show Property Update Options buttons", Application.GetOption("Show Property Update Options buttons") ' Show Property Update Options buttons
    Print #fle, ">>>Forms/Reports Tab"
    Print #fle, , "XP, 2003", "Selection Behavior         ", Application.GetOption("Selection Behavior")              ' Selection behavior
    Print #fle, , "XP, 2003", "Form Template              ", Application.GetOption("Form Template")                   ' Form template
    Print #fle, , "XP, 2003", "Report Template            ", Application.GetOption("Report Template")                 ' Report template
    Print #fle, , "XP, 2003", "Always Use Event Procedures", Application.GetOption("Always Use Event Procedures")     ' Always use event procedures
    Print #fle, , "    2003", "Show Smart Tags on Forms   ", Application.GetOption("Show Smart Tags on Forms")        ' Show Smart Tags on Forms
    Print #fle, , "    2003", "Themed Form Controls       ", Application.GetOption("Themed Form Controls")            ' Show Windows Themed Controls on Forms
    Print #fle, ">>>Advanced Tab"
    Print #fle, , "XP, 2003", "Ignore DDE Requests            ", Application.GetOption("Ignore DDE Requests")             ' DDE operations, Ignore DDE requests
    Print #fle, , "XP, 2003", "Enable DDE Refresh             ", Application.GetOption("Enable DDE Refresh")              ' DDE operations, Enable DDE refresh
    Print #fle, , "XP, 2003", "Default File Format            ", Application.GetOption("Default File Format")             ' Default File Format
    Print #fle, , "XP      ", "Row Limit                      ", Application.GetOption("Row Limit")                       ' Client-server settings, Default max records
    Print #fle, , "XP, 2003", "Default Open Mode for Databases", Application.GetOption("Default Open Mode for Databases") ' Default open mode
    Print #fle, , "XP, 2003", "Command-Line Arguments         ", Application.GetOption("Command-Line Arguments")          ' Command-line arguments
    Print #fle, , "XP, 2003", "OLE/DDE Timeout (sec)          ", Application.GetOption("OLE/DDE Timeout (sec)")           ' OLE/DDE timeout
    Print #fle, , "XP, 2003", "Default Record Locking         ", Application.GetOption("Default Record Locking")          ' Default record locking
    Print #fle, , "XP, 2003", "Refresh Interval (sec)         ", Application.GetOption("Refresh Interval (sec)")          ' Refresh interval
    Print #fle, , "XP, 2003", "Number of Update Retries       ", Application.GetOption("Number of Update Retries")        ' Number of update retries
    Print #fle, , "XP, 2003", "ODBC Refresh Interval (sec)    ", Application.GetOption("ODBC Refresh Interval (sec)")     ' ODBC fresh interval
    Print #fle, , "XP, 2003", "Update Retry Interval (msec)   ", Application.GetOption("Update Retry Interval (msec)")    ' Update retry interval
    Print #fle, , "XP, 2003", "Use Row Level Locking          ", Application.GetOption("Use Row Level Locking")           ' Open databases using record-level locking
    Print #fle, , "XP      ", "Save Login and Password        ", Application.GetOption("Save Login and Password")         ' Save login and password
    Print #fle, ">>>Pages Tab"
    Print #fle, , "XP, 2003", "Section Indent             ", Application.GetOption("Section Indent")                      ' Default Designer Properties, Section Indent
    Print #fle, , "XP, 2003", "Alternate Row Color        ", Application.GetOption("Alternate Row Color")                 ' Default Designer Properties, Alternative Row Color
    Print #fle, , "XP, 2003", "Caption Section Style      ", Application.GetOption("Caption Section Style")               ' Default Designer Properties, Caption Section Style
    Print #fle, , "XP, 2003", "Footer Section Style       ", Application.GetOption("Footer Section Style")                ' Default Designer Properties, Footer Section Style
    Print #fle, , "XP, 2003", "Use Default Page Folder    ", Application.GetOption("Use Default Page Folder")             ' Default Database/Project Properties, Use Default Page Folder
    Print #fle, , "XP, 2003", "Default Page Folder        ", Application.GetOption("Default Page Folder")                 ' Default Database/Project Properties, Default Page Folder
    Print #fle, , "XP, 2003", "Use Default Connection File", Application.GetOption("Use Default Connection File")         ' Default Database/Project Properties, Use Default Connection File
    Print #fle, , "XP, 2003", "Default Connection File    ", Application.GetOption("Default Connection File")             ' Default Database/Project Properties, Default Connection File
    Print #fle, ">>>Spelling Tab"
    Print #fle, , "XP, 2003", "Spelling dictionary language               ", Application.GetOption("Spelling dictionary language")                 ' Dictionary Language
    Print #fle, , "XP, 2003", "Spelling add words to                      ", Application.GetOption("Spelling add words to")                        ' Add words to
    Print #fle, , "XP, 2003", "Spelling suggest from main dictionary only ", Application.GetOption("Spelling suggest from main dictionary only")   ' Suggest from main dictionary only
    Print #fle, , "XP, 2003", "Spelling ignore words in UPPERCASE         ", Application.GetOption("Spelling ignore words in UPPERCASE")           ' Ignore words in UPPERCASE
    Print #fle, , "XP, 2003", "Spelling ignore words with number          ", Application.GetOption("Spelling ignore words with number")            ' Ignore words with numbers
    Print #fle, , "XP, 2003", "Spelling ignore Internet and file addresses", Application.GetOption("Spelling ignore Internet and file addresses")  ' Ignore Internet and file addresses
    Print #fle, , "XP, 2003", "Spelling use German post-reform rules      ", Application.GetOption("Spelling use German post-reform rules")        ' Language-specific, German: Use post-reform rules
    Print #fle, , "XP, 2003", "Spelling combine aux verb/adj              ", Application.GetOption("Spelling combine aux verb/adj")                ' Language-specific, Korean: Combine aux verb/adj.
    Print #fle, , "XP, 2003", "Spelling use auto-change list              ", Application.GetOption("Spelling use auto-change list")                ' Language-specific, Korean: Use auto-change list
    Print #fle, , "XP, 2003", "Spelling process compound nouns            ", Application.GetOption("Spelling process compound nouns")              ' Language-specific, Korean: Process compound nouns
    Print #fle, , "XP, 2003", "Spelling Hebrew modes                      ", Application.GetOption("Spelling Hebrew modes")                        ' Language-specific, Hebrew modes
    Print #fle, , "XP, 2003", "Spelling Arabic modes                      ", Application.GetOption("Spelling Arabic modes")                        ' Language-specific, Arabic modes
    Print #fle, ">>>International Tab"
    Print #fle, , "    2003", "Default direction ", Application.GetOption("Default direction")       ' Right-to-Left, Default direction
    Print #fle, , "    2003", "General alignment ", Application.GetOption("General alignment")       ' Right-to-Left, General alignment
    Print #fle, , "    2003", "Cursor movement   ", Application.GetOption("Cursor movement")         ' Right-to-Left, Cursor movement
    Print #fle, , "    2003", "Use Hijri Calendar", Application.GetOption("Use Hijri Calendar")      ' Use Hijri Calendar
    Print #fle, ">>>Error Checking Tab"
    Print #fle, , "    2003", "Enable Error Checking                        ", Application.GetOption("Enable Error Checking")                          ' Settings, Enable error checking
    Print #fle, , "    2003", "Error Checking Indicator Color               ", Application.GetOption("Error Checking Indicator Color")                 ' Settings, Error indicator color
    Print #fle, , "    2003", "Unassociated Label and Control Error Checking", Application.GetOption("Unassociated Label and Control Error Checking")  ' Form/Report Design Rules, Unassociated label and control
    Print #fle, , "    2003", "Keyboard Shortcut Errors Error Checking      ", Application.GetOption("Keyboard Shortcut Errors Error Checking")        ' Form/Report Design Rules, Keyboard shortcut errors
    Print #fle, , "    2003", "Invalid Control Properties Error Checking    ", Application.GetOption("Invalid Control Properties Error Checking")      ' Form/Report Design Rules, Invalid control properties
    Print #fle, , "    2003", "Common Report Errors Error Checking          ", Application.GetOption("Common Report Errors Error Checking")            ' Form/Report Design Rules, Common report errors
    Print #fle, ">>>Popular Tab"
    Print #fle, "   >>>Creating databases section"
    Print #fle, , "2007, 2010, 2013", "Default File Format       ", Application.GetOption("Default File Format")            ' Default file format
    Print #fle, , "2007, 2010, 2013", "Default Database Directory", Application.GetOption("Default Database Directory")     ' Default database folder
    Print #fle, , "2007, 2010, 2013", "New Database Sort Order   ", Application.GetOption("New Database Sort Order")        ' New database sort order
    Print #fle, ">>>Current Database Tab"
    Print #fle, "   >>>Application Options section"
    Print #fle, , "2007, 2010, 2013", "Auto Compact                   ", Application.GetOption("Auto Compact")                      ' Compact on Close
    Print #fle, , "2007, 2010, 2013", "Remove Personal Information    ", Application.GetOption("Remove Personal Information")       ' Remove personal information from file properties on save
    Print #fle, , "2007, 2010, 2013", "Themed Form Controls           ", Application.GetOption("Themed Form Controls")              ' Use Windows-themed Controls on Forms
    Print #fle, , "2007, 2010, 2013", "DesignWithData                 ", Application.GetOption("DesignWithData")                    ' Enable Layout View for this database
    Print #fle, , "2007, 2010, 2013", "CheckTruncatedNumFields        ", Application.GetOption("CheckTruncatedNumFields")           ' Check for truncated number fields
    Print #fle, , "2007, 2010, 2013", "Picture Property Storage Format", Application.GetOption("Picture Property Storage Format")   ' Picture Property Storage Format
    Print #fle, "   >>>Name AutoCorrect Options section"
    Print #fle, , "2007, 2010, 2013", "Track Name AutoCorrect Info ", Application.GetOption("Track Name AutoCorrect Info")   ' Track name AutoCorrect info
    Print #fle, , "2007, 2010, 2013", "Perform Name AutoCorrect    ", Application.GetOption("Perform Name AutoCorrect")      ' Perform name AutoCorrect
    Print #fle, , "2007, 2010, 2013", "Log Name AutoCorrect Changes", Application.GetOption("Log Name AutoCorrect Changes")  ' Log name AutoCorrect changes
    Print #fle, "   >>>Filter Lookup options for <Database Name> Database section"
    Print #fle, , "2007, 2010, 2013", "Show Values in Indexed    ", Application.GetOption("Show Values in Indexed")         ' Show list of values in, Local indexed fields
    Print #fle, , "2007, 2010, 2013", "Show Values in Non-Indexed", Application.GetOption("Show Values in Non-Indexed")     ' Show list of values in, Local nonindexed fields
    Print #fle, , "2007, 2010, 2013", "Show Values in Remote     ", Application.GetOption("Show Values in Remote")          ' Show list of values in, ODBC fields
    Print #fle, , "2007, 2010, 2013", "Show Values in Snapshot   ", Application.GetOption("Show Values in Snapshot")        ' Show list of values in, Records in local snapshot
    Print #fle, , "2007, 2010, 2013", "Show Values in Server     ", Application.GetOption("Show Values in Server")          ' Show list of values in, Records at server
    Print #fle, , "2007, 2010, 2013", "Show Values Limit         ", Application.GetOption("Show Values Limit")              ' Don't display lists where more than this number of records read
    Print #fle, ">>>Datasheet Tab"
    Print #fle, "   >>>Default colors section"
    Print #fle, , "2007, 2010, 2013", "Default Font Color      ", Application.GetOption("Default Font Color")               ' Font color
    Print #fle, , "2007, 2010, 2013", "Default Background Color", Application.GetOption("Default Background Color")         ' Background color
    Print #fle, , "2007, 2010, 2013", "_64                     ", Application.GetOption("_64")                              ' Alternate background color
    Print #fle, , "2007, 2010, 2013", "Default Gridlines Color ", Application.GetOption("Default Gridlines Color")          ' Gridlines color
    Print #fle, "   >>>Gridlines and cell effects section"
    Print #fle, , "2007, 2010, 2013", "Default Gridlines Horizontal", Application.GetOption("Default Gridlines Horizontal") ' Default gridlines showing, Horizontal
    Print #fle, , "2007, 2010, 2013", "Default Gridlines Vertical  ", Application.GetOption("Default Gridlines Vertical")   ' Default gridlines showing, Vertical
    Print #fle, , "2007, 2010, 2013", "Default Cell Effect         ", Application.GetOption("Default Cell Effect")          ' Default cell effect
    Print #fle, , "2007, 2010, 2013", "Default Column Width        ", Application.GetOption("Default Column Width")         ' Default column width
    Print #fle, "   >>>Default font section"
    Print #fle, , "2007, 2010, 2013", "Default Font Name     ", Application.GetOption("Default Font Name")                  ' Font
    Print #fle, , "2007, 2010, 2013", "Default Font Size     ", Application.GetOption("Default Font Size")                  ' Size
    Print #fle, , "2007, 2010, 2013", "Default Font Weight   ", Application.GetOption("Default Font Weight")                ' Weight
    Print #fle, , "2007, 2010, 2013", "Default Font Underline", Application.GetOption("Default Font Underline")             ' Underline
    Print #fle, , "2007, 2010, 2013", "Default Font Italic   ", Application.GetOption("Default Font Italic")                ' Italic
    Print #fle, ">>>Object Designers Tab"
    Print #fle, "   >>>Table design section"
    Print #fle, , "2007, 2010, 2013", "Default Text Field Size             ", Application.GetOption("Default Text Field Size")              ' Default text field size
    Print #fle, , "2007, 2010, 2013", "Default Number Field Size           ", Application.GetOption("Default Number Field Size")            ' Default number field size
    Print #fle, , "2007, 2010, 2013", "Default Field Type                  ", Application.GetOption("Default Field Type")                   ' Default field type
    Print #fle, , "2007, 2010, 2013", "AutoIndex on Import/Create          ", Application.GetOption("AutoIndex on Import/Create")           ' AutoIndex on Import/Create
    Print #fle, , "2007, 2010, 2013", "Show Property Update Options Buttons", Application.GetOption("Show Property Update Options Buttons") ' Show Property Update Option Buttons
    Print #fle, "   >>>Query design section"
    Print #fle, , "2007, 2010, 2013", "Show Table Names       ", Application.GetOption("Show Table Names")                  ' Show table names
    Print #fle, , "2007, 2010, 2013", "Output All Fields      ", Application.GetOption("Output All Fields")                 ' Output all fields
    Print #fle, , "2007, 2010, 2013", "Enable AutoJoin        ", Application.GetOption("Enable AutoJoin")                   ' Enable AutoJoin
    Print #fle, , "2007, 2010, 2013", "ANSI Query Mode        ", Application.GetOption("ANSI Query Mode")                   ' SQL Server Compatible Syntax (ANSI 92), This database
    Print #fle, , "2007, 2010, 2013", "ANSI Query Mode Default", Application.GetOption("ANSI Query Mode Default")           ' SQL Server Compatible Syntax (ANSI 92), Default for new databases
    Print #fle, , "2007, 2010, 2013", "Query Design Font Name ", Application.GetOption("Query Design Font Name")            ' Query design font, Font
    Print #fle, , "2007, 2010, 2013", "Query Design Font Size ", Application.GetOption("Query Design Font Size")            ' Query design font, Size
    Print #fle, "   >>>Forms/Reports section"
    Print #fle, , "2007, 2010, 2013", "Selection Behavior         ", Application.GetOption("Selection Behavior")            ' Selection behavior
    Print #fle, , "2007, 2010, 2013", "Form Template              ", Application.GetOption("Form Template")                 ' Form template
    Print #fle, , "2007, 2010, 2013", "Report Template            ", Application.GetOption("Report Template")               ' Report template
    Print #fle, , "2007, 2010, 2013", "Always Use Event Procedures", Application.GetOption("Always Use Event Procedures")   ' Always use event procedures
    Print #fle, "   >>>Error checking section"
    Print #fle, , "2007, 2010, 2013", "Enable Error Checking                        ", Application.GetOption("Enable Error Checking")                           ' Enable error checking
    Print #fle, , "2007, 2010, 2013", "Error Checking Indicator Color               ", Application.GetOption("Error Checking Indicator Color")                  ' Error indicator color
    Print #fle, , "2007, 2010, 2013", "Unassociated Label and Control Error Checking", Application.GetOption("Unassociated Label and Control Error Checking")   ' Check for unassociated label and control
    Print #fle, , "2007, 2010, 2013", "New Unassociated Labels Error Checking       ", Application.GetOption("New Unassociated Labels Error Checking")          ' Check for new unassociated labels
    Print #fle, , "2007, 2010, 2013", "Keyboard Shortcut Errors Error Checking      ", Application.GetOption("Keyboard Shortcut Errors Error Checking")         ' Check for keyboard shortcut errors
    Print #fle, , "2007, 2010, 2013", "Invalid Control Properties Error Checking    ", Application.GetOption("Invalid Control Properties Error Checking")       ' Check for invalid control properties
    Print #fle, , "2007, 2010, 2013", "Common Report Errors Error Checking          ", Application.GetOption("Common Report Errors Error Checking")             ' Check for common report errors
    Print #fle, ">>>Proofing Tab"
    Print #fle, "   >>>When correcting spelling in Microsoft Office programs section"
    Print #fle, , "2007, 2010, 2013", "Spelling ignore words in UPPERCASE         ", Application.GetOption("Spelling ignore words in UPPERCASE")            ' Ignore words in UPPERCASE
    Print #fle, , "2007, 2010, 2013", "Spelling ignore words with number          ", Application.GetOption("Spelling ignore words with number")             ' Ignore words that contain numbers
    Print #fle, , "2007, 2010, 2013", "Spelling ignore Internet and file addresses", Application.GetOption("Spelling ignore Internet and file addresses")   ' Ignore Internet and file addresses
    Print #fle, , "2007, 2010, 2013", "Spelling suggest from main dictionary only ", Application.GetOption("Spelling suggest from main dictionary only")    ' Suggest from main dictionary only
    Print #fle, , "2007, 2010, 2013", "Spelling dictionary language               ", Application.GetOption("Spelling dictionary language")                  ' Dictionary Language
    Print #fle, ">>>Advanced Tab"
    Print #fle, "   >>>Editing section"
    Print #fle, , "2007, 2010, 2013", "Move After Enter                ", Application.GetOption("Move After Enter")                     ' Move after enter
    Print #fle, , "2007, 2010, 2013", "Behavior Entering Field         ", Application.GetOption("Behavior Entering Field")              ' Behavior entering field
    Print #fle, , "2007, 2010, 2013", "Arrow Key Behavior              ", Application.GetOption("Arrow Key Behavior")                   ' Arrow key behavior
    Print #fle, , "2007, 2010, 2013", "Cursor Stops at First/Last Field", Application.GetOption("Cursor Stops at First/Last Field")     ' Cursor stops at first/last field
    Print #fle, , "2007, 2010, 2013", "Default Find/Replace Behavior   ", Application.GetOption("Default Find/Replace Behavior")        ' Default find/replace behavior
    Print #fle, , "2007, 2010, 2013", "Confirm Record Changes          ", Application.GetOption("Confirm Record Changes")               ' Confirm, Record changes
    Print #fle, , "2007, 2010, 2013", "Confirm Document Deletions      ", Application.GetOption("Confirm Document Deletions")           ' Confirm, Document deletions
    Print #fle, , "2007, 2010, 2013", "Confirm Action Queries          ", Application.GetOption("Confirm Action Queries")               ' Confirm, Action queries
    Print #fle, , "2007, 2010, 2013", "Default Direction               ", Application.GetOption("Default Direction")                    ' Default direction
    Print #fle, , "2007, 2010, 2013", "General Alignment               ", Application.GetOption("General Alignment")                    ' General alignment
    Print #fle, , "2007, 2010, 2013", "Cursor Movement                 ", Application.GetOption("Cursor Movement")                      ' Cursor movement
    Print #fle, , "2007, 2010, 2013", "Datasheet Ime Control           ", Application.GetOption("Datasheet Ime Control")                ' Datasheet IME control
    Print #fle, , "2007, 2010, 2013", "Use Hijri Calendar              ", Application.GetOption("Use Hijri Calendar")                   ' Use Hijri Calendar
    Print #fle, "   >>>Display section"
    Print #fle, , "2007, 2010, 2013", "Size of MRU File List               ", "Not Tracked"     'Application.GetOption("Size of MRU File List")                ' Show this number of Recent Documents
    Print #fle, , "2007, 2010, 2013", "Show Status Bar                     ", Application.GetOption("Show Status Bar")                      ' Status bar
    Print #fle, , "2007, 2010, 2013", "Show Animations                     ", Application.GetOption("Show Animations")                      ' Show animations
    Print #fle, , "2007, 2010, 2013", "Show Smart Tags on Datasheets       ", Application.GetOption("Show Smart Tags on Datasheets")        ' Show Smart Tags on Datasheets
    Print #fle, , "2007, 2010, 2013", "Show Smart Tags on Forms and Reports", Application.GetOption("Show Smart Tags on Forms and Reports") ' Show Smart Tags on Forms and Reports
    Print #fle, , "2007, 2010, 2013", "Show Macro Names Column             ", Application.GetOption("Show Macro Names Column")              ' Show in Macro Design, Names column
    Print #fle, , "2007, 2010, 2013", "Show Conditions Column              ", Application.GetOption("Show Conditions Column")               ' Show in Macro Design, Conditions column
    Print #fle, "   >>>Printing section"
    Print #fle, , "2007, 2010, 2013", "Left Margin  ", Application.GetOption("Left Margin")         ' Left margin
    Print #fle, , "2007, 2010, 2013", "Right Margin ", Application.GetOption("Right Margin")        ' Right margin
    Print #fle, , "2007, 2010, 2013", "Top Margin   ", Application.GetOption("Top Margin")          ' Top margin
    Print #fle, , "2007, 2010, 2013", "Bottom Margin", Application.GetOption("Bottom Margin")       ' Bottom margin
    Print #fle, "   >>>General section"
    Print #fle, , "2007, 2010, 2013", "Provide Feedback with Sound             ", Application.GetOption("Provide Feedback with Sound")                  ' Provide feedback with sound
    Print #fle, , "2007, 2010, 2013", "Four-Digit Year Formatting              ", Application.GetOption("Four-Digit Year Formatting")                   ' Use four-year digit year formatting, This database
    Print #fle, , "2007, 2010, 2013", "Four-Digit Year Formatting All Databases", Application.GetOption("Four-Digit Year Formatting All Databases")     ' Use four-year digit year formatting, All databases
    Print #fle, "   >>>Advanced section"
    Print #fle, , "2007, 2010, 2013", "Open Last Used Database When Access Starts", Application.GetOption("Open Last Used Database When Access Starts")     ' Open last used database when Access starts
    Print #fle, , "2007, 2010, 2013", "Default Open Mode for Databases           ", Application.GetOption("Default Open Mode for Databases")                ' Default open mode
    Print #fle, , "2007, 2010, 2013", "Default Record Locking                    ", Application.GetOption("Default Record Locking")                         ' Default record locking
    Print #fle, , "2007, 2010, 2013", "Use Row Level Locking                     ", Application.GetOption("Use Row Level Locking")                          ' Open databases by using record-level locking
    Print #fle, , "2007, 2010, 2013", "OLE/DDE Timeout (sec)                     ", Application.GetOption("OLE/DDE Timeout (sec)")                          ' OLE/DDE timeout (sec)
    Print #fle, , "2007, 2010, 2013", "Refresh Interval (sec)                    ", Application.GetOption("Refresh Interval (sec)")                         ' Refresh interval (sec)
    Print #fle, , "2007, 2010, 2013", "Number of Update Retries                  ", Application.GetOption("Number of Update Retries")                       ' Number of update retries
    Print #fle, , "2007, 2010, 2013", "ODBC Refresh Interval (sec)               ", Application.GetOption("ODBC Refresh Interval (sec)")                    ' ODBC refresh interval (sec)
    Print #fle, , "2007, 2010, 2013", "Update Retry Interval (msec)              ", Application.GetOption("Update Retry Interval (msec)")                   ' Update retry interval (msec)
    Print #fle, , "2007, 2010, 2013", "Ignore DDE Requests                       ", Application.GetOption("Ignore DDE Requests")                            ' DDE operations, Ignore DDE requests
    Print #fle, , "2007, 2010, 2013", "Enable DDE Refresh                        ", Application.GetOption("Enable DDE Refresh")                             ' DDE operations, Enable DDE refresh
    Print #fle, , "2007, 2010, 2013", "Command-Line Arguments                    ", Application.GetOption("Command-Line Arguments")                         ' Command-line arguments

PROC_EXIT:
    Set dbs = Nothing
    Close fle
    Exit Sub

PROC_ERR:
    If Err = 2091 Then          ' '...' is an invalid name.
        If Not IsMissing(varDebug) Then Debug.Print "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfAccessApplicationOptions of Class aegit_expClass"
        Print #fle, "!" & Err.Description
        Err.Clear
    ElseIf Err = 3270 Then      ' Property not found.
        Err.Clear
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfAccessApplicationOptions of Class aegit_expClass"
    End If
    Resume Next

End Sub

Private Sub OutputListOfApplicationProperties()
' Ref: http://www.granite.ab.ca/access/settingstartupoptions.htm

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim fle As Integer
    Dim strError As String

    fle = FreeFile()
    Open aegitSourceFolder & "\" & aeAppListPrp For Output As #fle

    Dim i As Integer
    Dim strPropName As String
    Dim varPropValue As Variant
    Dim varPropType As Variant
    Dim varPropInherited As Variant

    With dbs
        For i = 0 To (.Properties.Count - 1)
            strError = vbNullString
            strPropName = .Properties(i).Name
'''            varPropValue = Null
            ' Fixed for error 3251
            varPropValue = .Properties(i).Value
            varPropType = .Properties(i).Type
            varPropInherited = .Properties(i).Inherited
            Print #fle, strPropName & ": " & varPropValue & ", " & _
                varPropType & ", " & varPropInherited & ";" & strError
        Next i
    End With

PROC_EXIT:
    Set dbs = Nothing
    Close fle
    Exit Sub

PROC_ERR:
    If Err = 3251 Then
        strError = " " & Err.Number & ", '" & Err.Description & "'"
        varPropValue = Null
        Resume Next
        'Debug.Print "Erl=" & Erl & " Error " & Err.Number & " strPropName=" & strPropName & " (" & Err.Description & ") in procedure OutputListOfApplicationProperties of Class aegit_expClass"
        'Print #fle, "!" & Err.Description, strPropName
        'Err.Clear
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfApplicationProperties of Class aegit_expClass"
    End If
    Resume Next

End Sub

Private Function Pause(ByVal NumberOfSeconds As Variant) As Boolean
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Pause of Class aegit_expClass"
    Resume PROC_EXIT

End Function

Private Sub WaitSeconds(ByVal intSeconds As Integer)
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure WaitSeconds of Class aegit_expClass"
    Resume PROC_EXIT
End Sub

Private Function aeGetReferences(Optional ByVal varDebug As Variant) As Boolean
' Ref: http://vbadud.blogspot.com/2008/04/get-references-of-vba-project.html
' Ref: http://www.pcreview.co.uk/forums/Type-property-reference-object-vbulletin-project-t3793816.html
' Ref: http://www.cpearson.com/excel/missingreferences.aspx
' Ref: http://allenbrowne.com/ser-38.html
' Ref: http://access.mvps.org/access/modules/mdl0022.htm (References Wizard)
' Ref: http://www.accessmvp.com/djsteele/AccessReferenceErrors.html
' ====================================================================
' Author:   Peter F. Ennis
' Date:     November 28, 2012
' Comment:  Added and adapted from aeladdin (tm) code
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
' ====================================================================

    Dim i As Integer
    Dim RefName As String
    Dim RefDesc As String
    Dim blnRefBroken As Boolean
    Dim strFile As String

    Dim vbaProj As Object
    Set vbaProj = Application.VBE.ActiveVBProject

    On Error GoTo PROC_ERR

    Debug.Print "aeGetReferences"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to aeGetReferences"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeGetReferences"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    strFile = aestrSourceLocation & aeRefTxtFile
    
    If Dir$(strFile) <> vbNullString Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    If Not IsMissing(varDebug) Then
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

        If Not IsMissing(varDebug) Then Debug.Print , , vbaProj.References(i).Name, vbaProj.References(i).Description
        If Not IsMissing(varDebug) Then Debug.Print , , , vbaProj.References(i).FullPath
        If Not IsMissing(varDebug) Then Debug.Print , , , vbaProj.References(i).GUID

        Print #1, , , vbaProj.References(i).Name, vbaProj.References(i).Description
        Print #1, , , , vbaProj.References(i).FullPath
        Print #1, , , , vbaProj.References(i).GUID

        ' Returns a Boolean value indicating whether or not the Reference object points to a valid reference in the registry. Read-only.
        If Application.VBE.ActiveVBProject.References(i).IsBroken = True Then
              blnRefBroken = True
              If Not IsMissing(varDebug) Then Debug.Print , , vbaProj.References(i).Name, "blnRefBroken=" & blnRefBroken
              Print #1, , , vbaProj.References(i).Name, "blnRefBroken=" & blnRefBroken
        End If
    Next
    If Not IsMissing(varDebug) Then Debug.Print , "<*_*>"
    If Not IsMissing(varDebug) Then Debug.Print "<==<"

    Print #1, , "<*_*>"
    Print #1, "<==<"

    aeGetReferences = True

PROC_EXIT:
    Set vbaProj = Nothing
    Close 1
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeGetReferences of Class aegit_expClass"
    If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeGetReferences of Class aegit_expClass"
    aeGetReferences = False
    Resume PROC_EXIT

End Function

Private Function LongestTableName() As Integer
' ====================================================================
' Author:   Peter F. Ennis
' Date:     November 30, 2012
' Comment:  Return the length of the longest table name
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
' ====================================================================

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim intTNLen As Integer

    On Error GoTo PROC_ERR

    intTNLen = 0
    Set dbs = CurrentDb()
    For Each tdf In CurrentDb.TableDefs
        If Not (Left$(tdf.Name, 4) = "MSys" _
                Or Left$(tdf.Name, 4) = "~TMP" _
                Or Left$(tdf.Name, 3) = "zzz") Then
            If Len(tdf.Name) > intTNLen Then
                intTNLen = Len(tdf.Name)
            End If
        End If
    Next tdf

    LongestTableName = intTNLen

PROC_EXIT:
    Set tdf = Nothing
    Set dbs = Nothing
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestTableName of Class aegit_expClass"
    LongestTableName = 0
    Resume PROC_EXIT

End Function

Private Function LongestFieldPropsName() As Boolean
' =======================================================================
' Author:   Peter F. Ennis
' Date:     December 5, 2012
' Comment:  Return length of field properties for text output alignment
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
' =======================================================================

    Dim dbs As DAO.Database
    Dim tblDef As DAO.TableDef
    Dim fld As DAO.Field

    On Error GoTo PROC_ERR

    aeintFNLen = 0
    aeintFTLen = 0
    aeintFDLen = 0

    Set dbs = CurrentDb()

    For Each tblDef In CurrentDb.TableDefs
        If Not (Left$(tblDef.Name, 4) = "MSys" _
                Or Left$(tblDef.Name, 4) = "~TMP" _
                Or Left$(tblDef.Name, 3) = "zzz") Then
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
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestFieldPropsName of Class aegit_expClass"
    LongestFieldPropsName = False
    Resume PROC_EXIT

End Function

Private Function SizeString(ByVal Text As String, ByVal Length As Long, _
    Optional ByVal TextSide As SizeStringSide = TextLeft, _
    Optional ByVal PadChar As String = " ") As String
' Ref: http://www.cpearson.com/excel/sizestring.htm
' Enum SizeStringSide is used by SizeString to indicate whether the
' supplied text appears on the left or right side of result string.
' =========================================================================
' SizeString
' This procedure creates a string of a specified length. Text is the original string
' to include, and Length is the length of the result string. TextSide indicates whether
' Text should appear on the left (in which case the result is padded on the right with
' PadChar) or on the right (in which case the string is padded on the left). When padding on
' either the left or right, padding is done using the PadChar. character. If PadChar is omitted,
' a space is used. If PadChar is longer than one character, the left-most character of PadChar
' is used. If PadChar is an empty string, a space is used. If TextSide is neither
' TextLeft or TextRight, the procedure uses TextLeft.
' =========================================================================

    On Error GoTo 0
    Dim sPadChar As String

    If Len(Text) >= Length Then
        ' if the source string is longer than the specified length, return the
        ' Length left characters
        SizeString = Left$(Text, Length)
        Exit Function
    End If

    If Len(PadChar) = 0 Then
        ' PadChar is an empty string. use a space.
        sPadChar = " "
    Else
        ' use only the first character of PadChar
        sPadChar = Left$(PadChar, 1)
    End If

    If (TextSide <> TextLeft) And (TextSide <> TextRight) Then
        ' if TextSide was neither TextLeft nor TextRight, use TextLeft.
        TextSide = TextLeft
    End If

    If TextSide = TextLeft Then
        ' if the text goes on the left, fill out the right with spaces
        SizeString = Text & String$(Length - Len(Text), sPadChar)
    Else
        ' otherwise fill on the left and put the Text on the right
        SizeString = String$(Length - Len(Text), sPadChar) & Text
    End If

End Function

Private Function GetLinkedTableCurrentPath(ByVal strTblName As String) As String
' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=198057
' =========================================================================
' Procedure : GetLinkedTableCurrentPath
' Purpose   : Return Current Path of a Linked Table in Access and do not show password
' Author    : Peter F. Ennis
' Updated   : All notes moved to change log
' History   : See comment details, basChangeLog, commit messages on github
' =========================================================================

    On Error GoTo PROC_ERR

    Dim strConnect As String
    Dim intStrConnectLen As Integer
    Dim intEqualPos As Integer
    Dim intDatabasePos As Integer
    Dim strMidLink As String

    If Len(CurrentDb.TableDefs(strTblName).Connect) > 0 Then
        ' Linked table exists, but is the link valid?
        ' The next line of code will generate Errors 3011 or 3024 if it isn't
        CurrentDb.TableDefs(strTblName).RefreshLink
        ' If you get to this point, you have a valid, Linked Table
        strConnect = CurrentDb.TableDefs(strTblName).Connect
        intStrConnectLen = Len(strConnect)
        intDatabasePos = InStr(1, strConnect, "Database=") + 8
        strMidLink = Mid$(strConnect, intDatabasePos + 1, Len(strConnect) - intDatabasePos)
        'MsgBox "strTblName = " & strTblName & vbCrLf & _
            "strConnect = " & strConnect & vbCrLf & _
            "intStrConnectLen = " & intStrConnectLen & vbCrLf & _
            "intDatabasePos = " & intDatabasePos & " : " & Left$(strConnect, intDatabasePos) & vbCrLf & _
            "strMidLink = " & Mid$(strConnect, intDatabasePos + 1, Len(strConnect) - intDatabasePos) & vbCrLf _
            , vbInformation, "GetLinkedTableCurrentPath"
        GetLinkedTableCurrentPath = strMidLink
    Else
        GetLinkedTableCurrentPath = "Local Table=>" & strTblName
    End If

PROC_EXIT:
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 3265
            MsgBox "(" & strTblName & ") does not exist as either an Internal or Linked Table", _
                vbCritical, "Table Missing"
        Case 3011, 3024                 ' Linked Table does not exist or DB Path not valid
            MsgBox "(" & strTblName & ") is not a valid, Linked Table", vbCritical, "Link Not Valid"
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetLinkedTableCurrentPath of Class aegit_expClass"
    End Select
    Resume PROC_EXIT
End Function

Private Function FileLocked(ByVal strFileName As String) As Boolean
' Ref: http://support.microsoft.com/kb/209189

    On Error GoTo PROC_ERR

    Dim fle As Long
    fle = FreeFile()
    On Error Resume Next
    ' If the file is already opened by another process,
    ' and the specified type of access is not allowed,
    ' the Open operation fails and an error occurs.
    Open strFileName For Binary Access Read Write Lock Read Write As #fle
    Close fle
    ' If an error occurs, the document is currently open.
    If Err.Number <> 0 Then
        ' Display the error number and description.
        MsgBox ">>> Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FileLocked of Class aegit_expClass"
        FileLocked = True
        Err.Clear
    End If

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FileLocked of Class aegit_expClass"
    Resume PROC_EXIT

End Function

Private Function TableInfo(ByVal strTableName As String, Optional ByVal varDebug As Variant) As Boolean
' Ref: http://allenbrowne.com/func-06.html
' =============================================================================
' Purpose:  Display the field names, types, sizes and descriptions for a table
' Argument: Name of a table in the current database
' Updates:  Peter F. Ennis
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
' =============================================================================

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim sLen As Long
    Dim strLinkedTablePath As String

    On Error GoTo PROC_ERR

    strLinkedTablePath = vbNullString

    If IsMissing(varDebug) Then
        'Debug.Print , "varDebug IS missing so no parameter is passed to TableInfo"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        'Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to TableInfo"
        'Debug.Print , "DEBUGGING TURNED ON"
    End If

    Set dbs = CurrentDb()
    Set tdf = dbs.TableDefs(strTableName)
    sLen = Len("TABLE: ") + Len(strTableName)

    strLinkedTablePath = GetLinkedTableCurrentPath(strTableName)
    'MsgBox strLinkedTablePath & " " & Left$(strLinkedTablePath, 13), vbInformation, "TableInfo"

    aeintFDLen = LongestTableDescription(tdf.Name)

    If aeintFDLen < Len("DESCRIPTION") Then aeintFDLen = Len("DESCRIPTION")

    If Not IsMissing(varDebug) Then
        Debug.Print SizeString("-", sLen, TextLeft, "-")
        Debug.Print SizeString("TABLE: " & strTableName, sLen, TextLeft, " ")
        Debug.Print SizeString("-", sLen, TextLeft, "-")
        If Left(strLinkedTablePath, 13) <> "Local Table=>" Then
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
    If Left(strLinkedTablePath, 13) <> "Local Table=>" Then
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

    For Each fld In tdf.Fields
        If Not IsMissing(varDebug) Then
        'If Not IsMissing(varDebug) And aeintFDLen <> 11 Then
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
    If Not IsMissing(varDebug) Then Debug.Print
    'If Not IsMissing(varDebug) And aeintFDLen <> 11 Then Debug.Print
    Print #1, vbCrLf

    TableInfo = True

PROC_EXIT:
    Set fld = Nothing
    Set tdf = Nothing
    Set dbs = Nothing
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TableInfo of Class aegit_expClass"
    If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TableInfo of Class aegit_expClass"
    TableInfo = False
    Resume PROC_EXIT

End Function

Private Function GetDescrip(ByVal obj As Object) As String
    On Error Resume Next
    GetDescrip = obj.Properties("Description")
End Function

Private Function LongestTableDescription(ByVal strTblName As String) As Integer
' ?LongestTableDescription("tblCaseManager")

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim strLFD As String

    On Error GoTo PROC_ERR

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
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestTableDescription of Class aegit_expClass"
    LongestTableDescription = -1
    Resume PROC_EXIT

End Function

Private Function FieldTypeName(ByVal fld As DAO.Field) As String
' Ref: http://allenbrowne.com/func-06.html
' Purpose: Converts the numeric results of DAO Field.Type to text
    On Error GoTo 0
    Dim strReturn As String    ' Name to return

    Select Case CLng(fld.Type) ' fld.Type is Integer, but constants are Long.
        Case dbBoolean
            strReturn = "Yes/No"                        '  1
        Case dbByte
            strReturn = "Byte"                          '  2
        Case dbInteger
            strReturn = "Integer"                       '  3
        Case dbLong                                     '  4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency
            strReturn = "Currency"                      '  5
        Case dbSingle
            strReturn = "Single"                        '  6
        Case dbDouble
            strReturn = "Double"                        '  7
        Case dbDate
            strReturn = "Date/Time"                     '  8
        Case dbBinary
            strReturn = "Binary"                        '  9 (no interface)
        Case dbText                                     ' 10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        ' (no interface)
            End If
        Case dbLongBinary
            strReturn = "OLE Object"                    ' 11
        Case dbMemo                                     ' 12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID
            strReturn = "GUID"                          ' 15

        ' Attached tables only: cannot create these in JET.
        Case dbBigInt
            strReturn = "Big Integer"                   ' 16
        Case dbVarBinary
            strReturn = "VarBinary"                     ' 17
        Case dbChar
            strReturn = "Char"                          ' 18
        Case dbNumeric
            strReturn = "Numeric"                       ' 19
        Case dbDecimal
            strReturn = "Decimal"                       ' 20
        Case dbFloat
            strReturn = "Float"                         ' 21
        Case dbTime
            strReturn = "Time"                          ' 22
        Case dbTimeStamp
            strReturn = "Time Stamp"                    ' 23

        ' Constants for complex types don't work
        ' prior to Access 2007 and later.
        Case 101&
            strReturn = "Attachment"                    ' dbAttachment
        Case 102&
            strReturn = "Complex Byte"                  ' dbComplexByte
        Case 103&
            strReturn = "Complex Integer"               ' dbComplexInteger
        Case 104&
            strReturn = "Complex Long"                  ' dbComplexLong
        Case 105&
            strReturn = "Complex Single"                ' dbComplexSingle
        Case 106&
            strReturn = "Complex Double"                ' dbComplexDouble
        Case 107&
            strReturn = "Complex GUID"                  ' dbComplexGUID
        Case 108&
            strReturn = "Complex Decimal"               ' dbComplexDecimal
        Case 109&
            strReturn = "Complex Text"                  ' dbComplexText
        Case Else
            strReturn = "Field Type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn

End Function

Private Function aeDocumentTables(Optional ByVal varDebug As Variant) As Boolean
' Ref: http://www.tek-tips.com/faqs.cfm?fid=6905
' Document the tables, fields, and relationships
' Tables, field type, primary keys, foreign keys, indexes
' Relationships in the database with table, foreign table, primary keys, foreign keys
' Ref: http://allenbrowne.com/func-06.html

    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim blnResult As Boolean
    Dim intFailCount As Integer
    Dim strFile As String

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
    aestrLFN = vbNullString
    If aeintFNLen < 11 Then aeintFNLen = 11     ' Minimum required by design
    aeintFDLen = 0

    Debug.Print "aeDocumentTables"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTables"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTables"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    strFile = aestrSourceLocation & aeTblTxtFile
    'Debug.Print "aeDocumentTables", "strFile = " & strFile
    'Stop

    If Dir$(strFile) <> vbNullString Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    For Each tdf In CurrentDb.TableDefs
        If Not (Left$(tdf.Name, 4) = "MSys" _
                Or Left$(tdf.Name, 4) = "~TMP" _
                Or Left$(tdf.Name, 3) = "zzz") Then
            If Not IsMissing(varDebug) Then
                blnResult = TableInfo(tdf.Name, "WithDebugging")
                If Not blnResult Then intFailCount = intFailCount + 1
                If Not IsMissing(varDebug) And aeintFDLen <> 11 Then Debug.Print "aeintFDLen=" & aeintFDLen
            Else
                blnResult = TableInfo(tdf.Name)
                If Not blnResult Then intFailCount = intFailCount + 1
            End If
            aeintFDLen = 0
        End If
    Next tdf

    If Not IsMissing(varDebug) Then
        Debug.Print "intFailCount = " & intFailCount
    End If

    aeDocumentTables = True

PROC_EXIT:
    Set fld = Nothing
    Set tdf = Nothing
    Close 1
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTables of Class aegit_expClass"
    If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTables of Class aegit_expClass"
    aeDocumentTables = False
    Resume PROC_EXIT

End Function

Private Function aeDocumentTablesXML(Optional ByVal varDebug As Variant) As Boolean
' Ref: http://stackoverflow.com/questions/4867727/how-to-use-ms-access-saveastext-with-queries-specifically-stored-procedures

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim tbl As DAO.TableDef
    Dim strObjName As String

    Set dbs = CurrentDb

    Dim intFailCount As Integer

    ' Test for relative path
    Dim strTestPath As String
    strTestPath = aestrXMLLocation
    If Left(aestrXMLLocation, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid(aestrXMLLocation, 2, Len(aestrXMLLocation) - 1)
        aestrXMLLocation = strTestPath
        'Debug.Print , "aestrXMLLocation = " & aestrXMLLocation, "aeDocumentTablesXML"
        'Stop
    End If
    strTestPath = aegitXMLfolder
    If Left(aegitXMLfolder, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid(aegitXMLfolder, 2, Len(aegitXMLfolder) - 1)
        aegitXMLfolder = strTestPath
        'Debug.Print , "aegitXMLfolder = " & aegitXMLfolder, "aeDocumentTablesXML"
        'Stop
    End If

    'Debug.Print , "aegitXMLfolder = " & aegitXMLfolder, "aeDocumentTablesXML"
    If aegitXMLfolder = "default" Then
        aestrXMLLocation = aegitType.XMLfolder
    Else
        aestrXMLLocation = aegitXMLfolder
    End If
    'Debug.Print , "aestrXMLLocation = " & aestrXMLLocation, "aeDocumentTablesXML"

    If Not FolderExists(aestrXMLLocation) Then
        MsgBox aestrXMLLocation & " does not exist!", vbCritical, aeAPP_NAME
        Stop
    End If

    If Not IsNull(aegitDataXML(0)) And aegitDataXML(0) <> vbNullString Then
        If aegitExportDataToXML Then
            'MsgBox "aegitDataXML(0)=" & aegitDataXML(0), vbInformation, "aeDocumentTablesXML"
            If Not IsMissing(varDebug) Then
                OutputTheTableDataAsXML aegitDataXML(), varDebug
            Else
                OutputTheTableDataAsXML aegitDataXML()
            End If
        End If
    End If

    intFailCount = 0
    Debug.Print "aeDocumentTablesXML"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTablesXML"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTablesXML"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    If Not IsMissing(varDebug) Then Debug.Print ">List of tables exported as XML to " & aestrXMLLocation
    For Each tbl In dbs.TableDefs
        If Not LinkedTable(tbl.Name) And Not (tbl.Name Like "MSys*") Then
            strObjName = tbl.Name
            Application.ExportXML acExportTable, strObjName, , _
                        aestrXMLLocation & "tables_" & strObjName & ".xsd"
            If Not IsMissing(varDebug) Then
                Debug.Print , "- " & strObjName & ".xsd"
                PrettyXML aestrXMLLocation & "tables_" & strObjName & ".xsd", varDebug
            Else
                PrettyXML aestrXMLLocation & "tables_" & strObjName & ".xsd"
            End If
        End If
    Next

    If intFailCount > 0 Then
        aeDocumentTablesXML = False
    Else
        aeDocumentTablesXML = True
    End If

    If Not IsMissing(varDebug) Then
        Debug.Print "intFailCount = " & intFailCount
        Debug.Print "aeDocumentTablesXML = " & aeDocumentTablesXML
    End If

PROC_EXIT:
    Set tbl = Nothing
    Set dbs = Nothing
    Close 1
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTablesXML of Class aegit_expClass"
    If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTablesXML of Class aegit_expClass"
    aeDocumentTablesXML = False
    Resume PROC_EXIT

End Function

Private Sub OutputTheSchemaFile() ' CreateDbScript()
' Remou - Ref: http://stackoverflow.com/questions/698839/how-to-extract-the-schema-of-an-access-mdb-database/9910716#9910716

    On Error GoTo 0
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
        If Not (Left$(tdf.Name, 4) = "MSys" _
                Or Left$(tdf.Name, 4) = "~TMP" _
                Or Left$(tdf.Name, 3) = "zzz") Then

            strLinkedTablePath = GetLinkedTableCurrentPath(tdf.Name)
            If Left(strLinkedTablePath, 13) <> "Local Table=>" Then
                f.WriteLine vbCrLf & "'OriginalLink=>" & strLinkedTablePath
            Else
                f.WriteLine vbCrLf & "'Local Table"
            End If

            strSQL = "strSQL=""CREATE TABLE [" & tdf.Name & "] ("
            strFlds = vbNullString

            For Each fld In tdf.Fields

                strFlds = strFlds & ",[" & fld.Name & "] "

            ' Constants for complex types don't work prior to Access 2007 and later.

                Select Case fld.Type
                    Case dbText
                        ' No look-up fields
                        strFlds = strFlds & "Text (" & fld.size & ")"
                    Case 109&                                   ' dbComplexText
                        strFlds = strFlds & "Text (" & fld.size & ")"
                    Case dbMemo
                        If (fld.Attributes And dbHyperlinkField) = 0& Then
                            strFlds = strFlds & "Memo"
                        Else
                            strFlds = strFlds & "Hyperlink"
                        End If
                    Case dbByte
                        strFlds = strFlds & "Byte"
                    Case 102&                                   ' dbComplexByte
                        strFlds = strFlds & "Complex Byte"
                    Case dbInteger
                        strFlds = strFlds & "Integer"
                    Case 103&                                   ' dbComplexInteger
                        strFlds = strFlds & "Complex Integer"
                    Case dbLong
                        If (fld.Attributes And dbAutoIncrField) = 0& Then
                            strFlds = strFlds & "Long"
                        Else
                            strFlds = strFlds & "Counter"
                        End If
                    Case 104&                                   ' dbComplexLong
                        strFlds = strFlds & "Complex Long"
                    Case dbSingle
                        strFlds = strFlds & "Single"
                    Case 105&                                   ' dbComplexSingle
                        strFlds = strFlds & "Complex Single"
                    Case dbDouble
                        strFlds = strFlds & "Double"
                    Case 106&                                   ' dbComplexDouble
                        strFlds = strFlds & "Complex Double"
                    Case dbGUID
                        strFlds = strFlds & "GUID"
                        'strFlds = strFlds & "Replica"
                    Case 107&                                   ' dbComplexGUID
                        strFlds = strFlds & "Complex GUID"
                    Case dbDecimal
                        strFlds = strFlds & "Decimal"
                    Case 108&                                   ' dbComplexDecimal
                        strFlds = strFlds & "Complex Decimal"
                    Case dbDate
                        strFlds = strFlds & "DateTime"
                    Case dbCurrency
                        strFlds = strFlds & "Currency"
                    Case dbBoolean
                        strFlds = strFlds & "YesNo"
                    Case dbLongBinary
                        strFlds = strFlds & "OLE Object"
                    Case 101&                                   ' dbAttachment
                        strFlds = strFlds & "Attachment"
                    Case dbBinary
                        strFlds = strFlds & "Binary"
                    Case Else
                        MsgBox "Unknown fld.Type=" & fld.Type & " in procedure OutputTheSchemaFile of aegit_expClass", vbCritical, aeAPP_NAME
                        Debug.Print "Unknown fld.Type=" & fld.Type & " in procedure OutputTheSchemaFile of aegit_expClass" & vbCrLf & _
                                "tdf.Name=" & tdf.Name & " strFlds=" & strFlds
                End Select

            Next

            strSQL = strSQL & Mid$(strFlds, 2) & " )""" & vbCrLf & "Currentdb.Execute strSQL"
            f.WriteLine vbCrLf & strSQL

            ' Indexes
            For Each ndx In tdf.Indexes

                If ndx.Unique Then
                    strSQL = "strSQL=""CREATE UNIQUE INDEX "
                Else
                    strSQL = "strSQL=""CREATE INDEX "
                End If

                strSQL = strSQL & "[" & ndx.Name & "] ON [" & tdf.Name & "] ("
                strFlds = vbNullString

                For Each fld In tdf.Fields
                    strFlds = ",[" & fld.Name & "]"
                Next

                strSQL = strSQL & Mid$(strFlds, 2) & ") "
                strCn = vbNullString

                If ndx.Primary Then
                    strCn = " PRIMARY"
                End If

                If ndx.Required Then
                    strCn = strCn & " DISALLOW NULL"
                End If

                If ndx.IgnoreNulls Then
                    strCn = strCn & " IGNORE NULL"
                End If

                If Trim$(strCn) <> vbNullString Then
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

Private Function isPK(ByVal tdf As DAO.TableDef, ByVal strField As String) As Boolean
    On Error GoTo 0
    Dim Idx As DAO.Index
    Dim fld As DAO.Field
    For Each Idx In tdf.Indexes
        If Idx.Primary Then
            For Each fld In Idx.Fields
                If strField = fld.Name Then
                    isPK = True
                    Exit Function
                End If
            Next fld
        End If
    Next Idx
End Function

Private Function isIndex(ByVal tdf As DAO.TableDef, ByVal strField As String) As Boolean
    On Error GoTo 0
    Dim Idx As DAO.Index
    Dim fld As DAO.Field
    For Each Idx In tdf.Indexes
        For Each fld In Idx.Fields
            If strField = fld.Name Then
                isIndex = True
                Exit Function
            End If
        Next fld
    Next Idx
End Function

Private Function isFK(ByVal tdf As DAO.TableDef, ByVal strField As String) As Boolean
    On Error GoTo 0
    Dim Idx As DAO.Index
    Dim fld As DAO.Field
    For Each Idx In tdf.Indexes
        If Idx.Foreign Then
            For Each fld In Idx.Fields
                If strField = fld.Name Then
                    isFK = True
                    Exit Function
                End If
            Next fld
        End If
    Next Idx
End Function

Private Function aeDocumentRelations(Optional ByVal varDebug As Variant) As Boolean
' Ref: http://www.tek-tips.com/faqs.cfm?fid=6905
  
    Dim strDocument As String
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim Idx As DAO.Index
    Dim prop As DAO.Property
    Dim strFile As String

    On Error GoTo PROC_ERR

    Debug.Print "aeDocumentRelations"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentRelations"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentRelations"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    strFile = aestrSourceLocation & aeRelTxtFile
    
    If Dir$(strFile) <> vbNullString Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    For Each rel In CurrentDb.Relations
        If Not (Left$(rel.Name, 4) = "MSys" _
                        Or Left$(rel.Name, 4) = "~TMP" _
                        Or Left$(rel.Name, 3) = "zzz") Then
            strDocument = strDocument & vbCrLf & "Name: " & rel.Name & vbCrLf
            strDocument = strDocument & "  " & "Table: " & rel.Table & vbCrLf
            strDocument = strDocument & "  " & "Foreign Table: " & rel.ForeignTable & vbCrLf
            For Each fld In rel.Fields
                strDocument = strDocument & "  PK: " & fld.Name & "   FK:" & fld.ForeignName
                strDocument = strDocument & vbCrLf
            Next fld
        End If
    Next rel
    If Not IsMissing(varDebug) Then Debug.Print strDocument
    Print #1, strDocument
    
    aeDocumentRelations = True

PROC_EXIT:
    Set prop = Nothing
    Set Idx = Nothing
    Set fld = Nothing
    Set rel = Nothing
    Close 1
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentRelations of Class aegit_expClass"
    If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentRelations of Class aegit_expClass"
    aeDocumentRelations = False
    Resume PROC_EXIT

End Function

Private Function OutputQueriesSqlText() As Boolean
' Ref: http://www.pcreview.co.uk/forums/export-sql-saved-query-into-text-file-t2775525.html
' ====================================================================
' Author:   Peter F. Ennis
' Date:     December 3, 2012
' Comment:  Output the sql code of all queries to a text file
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
' ====================================================================

    Dim dbs As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strFile As String

    On Error GoTo PROC_ERR

    strFile = aestrSourceLocation & aeSqlTxtFile
    
    If Dir$(strFile) <> vbNullString Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    Set dbs = CurrentDb
    For Each qdf In dbs.QueryDefs
        If Not (Left$(qdf.Name, 4) = "MSys" Or Left$(qdf.Name, 4) = "~sq_" _
                        Or Left$(qdf.Name, 4) = "~TMP" _
                        Or Left$(qdf.Name, 3) = "zzz") Then
            Print #1, "<<<" & qdf.Name & ">>>" & vbCrLf & qdf.SQL
        End If
    Next

    OutputQueriesSqlText = True

PROC_EXIT:
    Set qdf = Nothing
    Set dbs = Nothing
    Close 1
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputQueriesSqlText of Class aegit_expClass"
    OutputQueriesSqlText = False
    Resume PROC_EXIT

End Function

Private Sub KillProperly(ByVal Killfile As String)
' Ref: http://word.mvps.org/faqs/macrosvba/DeleteFiles.htm

    On Error GoTo PROC_ERR

TryAgain:
    If Len(Dir$(Killfile)) > 0 Then
        SetAttr Killfile, vbNormal
        Kill Killfile
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    If Err = 70 Or Err = 75 Then
        Pause (0.25)
        Resume TryAgain
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " Killfile=" & Killfile & " (" & Err.Description & ") in procedure KillProperly of Class aegit_expClass"
    Resume PROC_EXIT

End Sub

Private Function GetPropEnum(ByVal typeNum As Long, Optional ByVal varDebug As Variant) As String
' Ref: http://msdn.microsoft.com/en-us/library/bb242635.aspx

    On Error GoTo PROC_ERR

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
        Case Else
            'MsgBox "Unknown typeNum:" & typeNum, vbInformation, aeAPP_NAME
            GetPropEnum = "Unknown typeNum"
            Debug.Print "Unknown typeNum:" & typeNum & " in procedure GetPropEnum of aegit_expClass"
    End Select

PROC_EXIT:
    Exit Function

PROC_ERR:
     Select Case Err.Number
'        Case 3251
'            strError = " " & Err.Number & ", '" & Err.Description & "'"
'            varPropValue = Null
'            Resume Next
        Case Else
            'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetPropEnum of Class aegit_expClass"
            If Not IsMissing(varDebug) Then Debug.Print "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetPropEnum of Class aegit_expClass"
            GetPropEnum = CStr(typeNum)
            Resume PROC_EXIT
    End Select

End Function

Private Function GetPrpValue(ByVal obj As Object) As String
    'On Error Resume Next
    On Error GoTo 0
    GetPrpValue = obj.Properties("Value")
End Function
 
Private Function OutputBuiltInPropertiesText(Optional ByVal varDebug As Variant) As Boolean
' Ref: http://www.jpsoftwaretech.com/listing-built-in-access-database-properties/

    Dim dbs As DAO.Database
    Dim prps As DAO.Properties
    Dim prp As DAO.Property
    Dim varPropValue As Variant
    Dim strFile As String
    Dim strError As String

    On Error GoTo PROC_ERR

    strFile = aestrSourceLocation & aePrpTxtFile

    If Dir$(strFile) <> vbNullString Then
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
        strError = vbNullString
        Print #1, "Name: " & prp.Name
        If Not IsMissing(varDebug) Then
            Print #1, "Type: " & GetPropEnum(prp.Type, varDebug)
        Else
            Print #1, "Type: " & GetPropEnum(prp.Type)
        End If
        ' Fixed for error 3251
        varPropValue = GetPrpValue(prp)
        Print #1, "Value: " & varPropValue
        Print #1, "Inherited: " & prp.Inherited & ";" & strError
        Print #1, "---"
    Next prp

    OutputBuiltInPropertiesText = True

PROC_EXIT:
    Set prp = Nothing
    Set prps = Nothing
    Set dbs = Nothing
    Close 1
    Exit Function

PROC_ERR:
     Select Case Err.Number
        Case 3251
            strError = " " & Err.Number & ", '" & Err.Description & "'"
            varPropValue = Null
            Resume Next
        Case Else
            'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputBuiltInPropertiesText of Class aegit_expClass"
            If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputBuiltInPropertiesText of Class aegit_expClass"
            OutputBuiltInPropertiesText = False
            Resume PROC_EXIT
    End Select

End Function
 
Private Function IsFileLocked(ByVal PathFileName As String) As Boolean
' Ref: http://accessexperts.com/blog/2012/03/06/checking-if-files-are-locked/

    On Error GoTo PROC_ERR

    Dim i As Integer

    If Len(Dir$(PathFileName)) Then
        i = FreeFile()
        Open PathFileName For Random Access Read Write Lock Read Write As #i
        Lock i      ' Redundant but let's be 100% sure
        Unlock i
        Close i
    Else
        ' Err.Raise 53
    End If

PROC_EXIT:
    On Error GoTo 0
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 70 ' Unable to acquire exclusive lock
            MsgBox "A:Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegit_expClass"
            IsFileLocked = True
        Case 9
            MsgBox "B:Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegit_expClass" & _
                    vbCrLf & "IsFileLocked Entry PathFileName=" & PathFileName, vbCritical, "ERROR=9"
            IsFileLocked = False
            'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegit_expClass"
            Resume PROC_EXIT
        Case Else
            MsgBox "C:Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegit_expClass"
            'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegit_expClass"
            Resume PROC_EXIT
    End Select
    Resume

End Function

Private Function DocumentTheContainer(ByVal strContainerType As String, ByVal strExt As String, Optional ByVal varDebug As Variant) As Boolean
' strContainerType: Forms, Reports, Scripts (Macros), Modules

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim cnt As DAO.Container
    Dim doc As DAO.Document
    Dim i As Integer
    Dim intAcObjType As Integer
    Dim strTheCurrentPathAndFile As String

    Set dbs = CurrentDb() ' use CurrentDb() to refresh Collections

    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to DocumentTheContainer"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to DocumentTheContainer"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    i = 0
    Set cnt = dbs.Containers(strContainerType)

    Select Case strContainerType
        Case "Forms"
            intAcObjType = 2    ' acForm
        Case "Reports"
            intAcObjType = 3    ' acReport
        Case "Scripts"
            intAcObjType = 4    ' acMacro
        Case "Modules"
            intAcObjType = 5    ' acModule
        Case Else
            MsgBox "Wrong Case Select in DocumentTheContainer", vbCritical, aeAPP_NAME
    End Select

    If Not IsMissing(varDebug) Then Debug.Print UCase$(strContainerType)

    For Each doc In cnt.Documents
        'Debug.Print "A", doc.Name, strContainerType
        If Not IsMissing(varDebug) Then Debug.Print , doc.Name
        If Not (Left$(doc.Name, 3) = "zzz" Or Left$(doc.Name, 4) = "~TMP") Then
            i = i + 1
            strTheCurrentPathAndFile = aestrSourceLocation & doc.Name & "." & strExt
            If IsFileLocked(strTheCurrentPathAndFile) Then
                MsgBox strTheCurrentPathAndFile & " is locked!", vbCritical, "STOP in DocumentTheContainer"
            End If
            KillProperly (strTheCurrentPathAndFile)
SaveAsText:
            If intAcObjType = 5 Then
                'Debug.Print "5:", doc.Name, fExclude(doc.Name)
                If fExclude(doc.Name) And pExclude Then
                    Debug.Print , "=> Excluded: " & doc.Name
                    GoTo NextDoc
                Else
                    Application.SaveAsText intAcObjType, doc.Name, strTheCurrentPathAndFile
                End If
            Else
                Application.SaveAsText intAcObjType, doc.Name, strTheCurrentPathAndFile
            End If
            If mblnUTF16 Then
                Application.SaveAsText intAcObjType, doc.Name, strTheCurrentPathAndFile
            Else
                Application.SaveAsText intAcObjType, doc.Name, strTheCurrentPathAndFile
                If intAcObjType = 2 Then
                    ' Convert UTF-16 to txt - fix for Access 2013
                    If NoBOM(strTheCurrentPathAndFile) Then
                        ' Conversion done
                    Else
                        ' Fallback to old method
                        If aeReadWriteStream(strTheCurrentPathAndFile) = True Then
                            'If intAcObjType = 2 Then Pause (0.25)
                            KillProperly (strTheCurrentPathAndFile)
                            Name strTheCurrentPathAndFile & ".clean.txt" As strTheCurrentPathAndFile
                        End If
                    End If
                End If
            End If
        End If
        '
        ' Ouput frm as txt
        If Not (Left$(doc.Name, 3) = "zzz" Or Left$(doc.Name, 4) = "~TMP") _
                Or Not fExclude(doc.Name) Then
            If strContainerType = "Forms" Then
                If Not IsMissing(varDebug) Then
                    CreateFormReportTextFile strTheCurrentPathAndFile, strTheCurrentPathAndFile & ".txt", varDebug
                Else
                    CreateFormReportTextFile strTheCurrentPathAndFile, strTheCurrentPathAndFile & ".txt"
                End If
            End If
        End If
NextDoc:
    Next doc

    If Not IsMissing(varDebug) Then
        Debug.Print , i & " EXPORTED!"
        Debug.Print , cnt.Documents.Count & " EXISTING!"
    End If

    DocumentTheContainer = True

PROC_EXIT:
    Set doc = Nothing
    Set cnt = Nothing
    Set dbs = Nothing
    Exit Function

PROC_ERR:
    If Err = 2220 Then  ' Run-time error 2220 Microsoft Access can't open the file
        Debug.Print "Err=2220 : Resume SaveAsText - " & doc.Name & " - " & strTheCurrentPathAndFile
        Err.Clear
        Pause (0.25)
        Resume SaveAsText
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DocumentTheContainer of Class aegit_expClass"
    DocumentTheContainer = False
    Resume PROC_EXIT

End Function

Private Sub KillAllFiles(ByVal strLoc As String, Optional ByVal varDebug As Variant)

    Dim strFile As String

    On Error GoTo PROC_ERR

    'Debug.Print "KillAllFiles", "strLoc = " & strLoc
    'Stop
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to KillAllFiles"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to KillAllFiles"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    ' Test for relative path
    Dim strTestPath As String
    strTestPath = aestrSourceLocation
    If Left(aestrSourceLocation, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid(aestrSourceLocation, 2, Len(aestrSourceLocation) - 1)
        aestrSourceLocation = strTestPath
        'Debug.Print , "aestrSourceLocation = " & aestrSourceLocation, "aeDocumentTheDatabase"
        'Stop
    End If

    If aegitFrontEndApp And strLoc = "src" Then
        ' Delete all the exported src files
        strFile = Dir$(aestrSourceLocation & "*.*")
        Do While strFile <> vbNullString
            KillProperly (aestrSourceLocation & strFile)
            ' Need to specify full path again because a file was deleted
            strFile = Dir$(aestrSourceLocation & "*.*")
        Loop
        strFile = Dir$(aestrSourceLocation & "xml\" & "*.*")
        Do While strFile <> vbNullString
            KillProperly (aestrSourceLocation & "xml\" & strFile)
            ' Need to specify full path again because a file was deleted
            strFile = Dir$(aestrSourceLocation & "xml\" & "*.*")
        Loop
    ElseIf Not aegitFrontEndApp And strLoc = "srcbe" Then
        ' Test for relative path
        'Dim strTestPath As String
        strTestPath = aegitSourceFolderBe
        If Left(aegitSourceFolderBe, 1) = "." Then
            strTestPath = CurrentProject.Path & Mid(aegitSourceFolderBe, 2, Len(aegitSourceFolderBe) - 1)
            aegitSourceFolderBe = strTestPath
        End If
        ' Delete all the exported srcbe files
        Debug.Print "KillAllFiles"
        Debug.Print , "aestrBackEndDb1 = " & aestrBackEndDb1
        Debug.Print , "aestrSourceLocation = " & aestrSourceLocation
        Debug.Print , "aegitSourceFolderBe = " & aegitSourceFolderBe
        '
        strFile = Dir$(aegitSourceFolderBe & "*.*")
        Do While strFile <> vbNullString
            KillProperly (aegitSourceFolderBe & strFile)
            ' Need to specify full path again because a file was deleted
            strFile = Dir$(aegitSourceFolderBe & "*.*")
        Loop
        strFile = Dir$(aegitSourceFolderBe & "xml\" & "*.*")
        Do While strFile <> vbNullString
            KillProperly (aegitSourceFolderBe & "xml\" & strFile)
            ' Need to specify full path again because a file was deleted
            strFile = Dir$(aegitSourceFolderBe & "xml\" & "*.*")
        Loop
        'Stop
    Else
        MsgBox "Bad strLoc", vbCritical, "STOP " & aeAPP_NAME
        Stop
    End If

    If aegitSetup Then
        strFile = Dir$(aestrXMLLocation & "*.*")
        Do While strFile <> vbNullString
            KillProperly (aestrXMLLocation & strFile)
            ' Need to specify full path again because a file was deleted
            strFile = Dir$(aestrXMLLocation & "*.*")
        Loop
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    If Err = 70 Then    ' Permission denied
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure KillAllFiles of Class aegit_expClass" _
            & vbCrLf & vbCrLf & _
            "Manually delete the files from git, compact and repair database, then try again!", vbCritical, "STOP"
        Stop
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure KillAllFiles of Class aegit_expClass"
    If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure KillAllFiles of Class aegit_expClass"
    Resume PROC_EXIT

End Sub

Private Function FolderExists(ByVal strPath As String) As Boolean
' Ref: http://allenbrowne.com/func-11.html
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Private Sub ListOrCloseAllOpenQueries(Optional ByVal strCloseAll As Variant)
' Ref: http://msdn.microsoft.com/en-us/library/office/aa210652(v=office.11).aspx

    On Error GoTo 0
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

Private Function aeDocumentTheDatabase(Optional ByVal varDebug As Variant) As Boolean
' Based on sample code from Arvin Meyer (MVP) June 2, 1999
' Ref: http://www.accessmvp.com/Arvin/DocDatabase.txt
' Ref: http://www.tek-tips.com/faqs.cfm?fid=6905
'=======================================================================
' Author:   Peter F. Ennis
' Date:     February 8, 2011
' Comment:  Uses the undocumented [Application.SaveAsText] syntax
'           To reload use the syntax [Application.LoadFromText]
'           Add explicit references for DAO
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'=======================================================================

    Dim dbs As DAO.Database
    Dim cnt As DAO.Container
    Dim doc As DAO.Document
    Dim qdf As DAO.QueryDef
    Dim i As Integer
    Dim strqdfName As String

    On Error GoTo PROC_ERR

    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    If aestrBackEndDb1 <> "default" Then
        OpenAllDatabases True
    End If

    If aegitSourceFolder = "default" Then
        aestrSourceLocation = aegitType.SourceFolder
        aegitSetup = True
    Else
        aestrSourceLocation = aegitSourceFolder
    End If

    If aegitXMLfolder = "default" Then
        aestrXMLLocation = aegitType.XMLfolder
    Else
        aestrXMLLocation = aegitXMLfolder
    End If

    If Not IsMissing(varDebug) Then
        Debug.Print , "Value for aestrSourceLocation = " & aestrSourceLocation
        Debug.Print , "Value for aegitXMLfolder = " & aegitXMLfolder
        Debug.Print , "Value for aestrXMLLocation = " & aestrXMLLocation
    End If

    ListOrCloseAllOpenQueries

    If Not IsMissing(varDebug) Then
        Debug.Print , ">==> aeDocumentTheDatabase >==>"
        Debug.Print , "Property Get SourceFolder = " & aestrSourceLocation
        Debug.Print , "Property Get XMLfolder = " & aestrXMLLocation
        Debug.Print , "Property Get BackEndDb1 = " & aestrBackEndDb1
    End If

    If aestrSourceLocation = vbNullString Then
        MsgBox "aestrSourceLocation is not set!", vbCritical, aeAPP_NAME
        Stop
    End If

    If aestrXMLLocation = vbNullString Then
        MsgBox "aestrXMLLocation is not set!", vbCritical, aeAPP_NAME
        Stop
    End If

    If aestrBackEndDb1 = vbNullString Then
        MsgBox "aestrBackEndDb1 is not set!", vbCritical, aeAPP_NAME
        Stop
    End If

    If IsMissing(varDebug) Then
        If aegitFrontEndApp Then KillAllFiles "src"
        If Not aegitFrontEndApp And aestrBackEndDb1 <> "default" Then KillAllFiles "srcbe"
    Else
        If aegitFrontEndApp Then KillAllFiles "src", varDebug
        If Not aegitFrontEndApp And aestrBackEndDb1 <> "default" Then KillAllFiles "srcbe", varDebug
    End If
    'Stop

    ' ===================================
    '    FORMS REPORTS SCRIPTS MODULES
    ' ===================================
    ' NOTE: Erl(0) Error 2950 if the ouput location does not exist so test for it first.

    'Debug.Print , "aestrSourceLocation = " & aestrSourceLocation, "aeDocumentTheDatabase"
    'Stop
    '
    ' Test for relative path
    Dim strTestPath As String
    strTestPath = aestrSourceLocation
    If Left(aestrSourceLocation, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid(aestrSourceLocation, 2, Len(aestrSourceLocation) - 1)
        aestrSourceLocation = strTestPath
        'Debug.Print , "aestrSourceLocation = " & aestrSourceLocation, "aeDocumentTheDatabase"
        'Stop
    End If

    If FolderExists(aestrSourceLocation) Then
        If Not IsMissing(varDebug) Then
            DocumentTheContainer "Forms", "frm", varDebug
            DocumentTheContainer "Reports", "rpt", varDebug
            DocumentTheContainer "Scripts", "mac", varDebug
            DocumentTheContainer "Modules", "bas", varDebug
        Else
            DocumentTheContainer "Forms", "frm"
            DocumentTheContainer "Reports", "rpt"
            DocumentTheContainer "Scripts", "mac"
            DocumentTheContainer "Modules", "bas"
        End If
    Else
        MsgBox aestrSourceLocation & " Does not exist!", vbCritical, aeAPP_NAME
        Stop
    End If

    ' =============
    '    QUERIES
    ' =============
    Set dbs = CurrentDb() ' Use CurrentDb() to refresh Collections
    i = 0
    If Not IsMissing(varDebug) Then Debug.Print "QUERIES"

    ' Delete all TEMP queries ...
    For Each qdf In CurrentDb.QueryDefs
        strqdfName = qdf.Name
        If Left$(strqdfName, 1) = "~" Then
            CurrentDb.QueryDefs.Delete strqdfName
            CurrentDb.QueryDefs.Refresh
        End If
    Next qdf
    If Not IsMissing(varDebug) Then Debug.Print , "Temp queries deleted"

    For Each qdf In CurrentDb.QueryDefs
        strqdfName = qdf.Name
        If Not IsMissing(varDebug) Then Debug.Print , strqdfName
        If Not (Left$(strqdfName, 4) = "MSys" Or Left$(strqdfName, 4) = "~sq_" _
                        Or Left$(strqdfName, 4) = "~TMP" _
                        Or Left$(strqdfName, 3) = "zzz") Then
            i = i + 1
            Application.SaveAsText acQuery, strqdfName, aestrSourceLocation & strqdfName & ".qry"
            ' Convert UTF-16 to txt - fix for Access 2013
            If aeReadWriteStream(aestrSourceLocation & strqdfName & ".qry") = True Then
                KillProperly (aestrSourceLocation & strqdfName & ".qry")
                Name aestrSourceLocation & strqdfName & ".qry" & ".clean.txt" As aestrSourceLocation & strqdfName & ".qry"
            End If
        End If
    Next qdf

    If Not IsMissing(varDebug) Then
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

    ' =============
    '    OUTPUTS
    ' =============
    If Not IsMissing(varDebug) Then
        OutputListOfContainers aeAppListCnt, varDebug
        OutputListOfAccessApplicationOptions varDebug
        If aegitExport.ExportCBID Then
            OutputListOfCommandBarIDs aeAppCmbrIds, varDebug
        End If
        OutputTableDataMacros varDebug
        OutputPrinterInfo "Debug"
        If aeExists("Tables", "aetlkpStates", varDebug) Then
            OutputTableDataAsFormattedText "aetlkpStates", varDebug
        End If
        If aeExists("Tables", "USysRibbons", varDebug) Then
            OutputTableDataAsFormattedText "USysRibbons", varDebug
        End If
        OutputCatalogUserCreatedObjects varDebug
        OutputListOfAllHiddenQueries varDebug
        OutputBuiltInPropertiesText varDebug
        OutputAllContainerProperties varDebug
        OutputTableProperties varDebug
    Else
        OutputListOfContainers aeAppListCnt
        OutputListOfAccessApplicationOptions
        If aegitExport.ExportCBID Then
            OutputListOfCommandBarIDs aeAppCmbrIds
        End If
        OutputTableDataMacros
        OutputPrinterInfo
        If aeExists("Tables", "aetlkpStates") Then
            OutputTableDataAsFormattedText "aetlkpStates"
        End If
        If aeExists("Tables", "USysRibbons") Then
            OutputTableDataAsFormattedText "USysRibbons"
        End If
        OutputCatalogUserCreatedObjects
        OutputListOfAllHiddenQueries
        OutputBuiltInPropertiesText
        OutputAllContainerProperties
        OutputTableProperties
    End If

    OutputListOfApplicationProperties
    OutputQueriesSqlText
    OutputFieldLookupControlTypeList
    OutputTheSchemaFile

    If aegitExport.ExportQAT Then
        If Not IsMissing(varDebug) Then
            OutputTheQAT aeAppListQAT, varDebug
        Else
            OutputTheQAT aeAppListQAT
        End If
    End If

    aeDocumentTheDatabase = True

PROC_EXIT:
    Set qdf = Nothing
    Set doc = Nothing
    Set cnt = Nothing
    Set dbs = Nothing
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTheDatabase of Class aegit_expClass"
    If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTheDatabase of Class aegit_expClass"
    aeDocumentTheDatabase = False
    Resume PROC_EXIT

End Function

Private Function aeExists(ByVal strAccObjType As String, _
                        ByVal strAccObjName As String, Optional ByVal varDebug As Variant) As Boolean
' Ref: http://vbabuff.blogspot.com/2010/03/does-access-object-exists.html
'
' =======================================================================
' Author:     Peter F. Ennis
' Date:       February 18, 2011
' Comment:    Return True if the object exists
' Parameters:
'             strAccObjType: "Tables", "Queries", "Forms",
'                            "Reports", "Macros", "Modules"
'             strAccObjName: The name of the object
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
' =======================================================================

    Dim objType As Object
    Dim obj As Variant
    
    On Error GoTo PROC_ERR

    Debug.Print "aeExists", strAccObjType, strAccObjName
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to aeExists"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeExists"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    If Not IsMissing(varDebug) Then Debug.Print ">==> aeExists >==>"

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
            MsgBox "Wrong option!", vbCritical, "in procedure aeExists of Class aegit_expClass"
            If Not IsMissing(varDebug) Then
                Debug.Print , "strAccObjType = >" & strAccObjType & "< is  a false value"
                Debug.Print , "Option allowed is one of 'Tables', 'Queries', 'Forms', 'Reports', 'Macros', 'Modules'"
                Debug.Print "<==<"
            End If
            aeExists = False
            Set obj = Nothing
            Exit Function
    End Select

    If Not IsMissing(varDebug) Then Debug.Print , "strAccObjType = " & strAccObjType
    If Not IsMissing(varDebug) Then Debug.Print , "strAccObjName = " & strAccObjName

    For Each obj In objType
        If Not IsMissing(varDebug) Then Debug.Print , obj.Name, strAccObjName
        If obj.Name = strAccObjName Then
            If Not IsMissing(varDebug) Then
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
    If Not IsMissing(varDebug) And aeExists = False Then
        Debug.Print , strAccObjName & " DOES NOT EXIST!"
        Debug.Print "<==<"
    End If

PROC_EXIT:
    Set obj = Nothing
    Exit Function

PROC_ERR:
    If Err = 3011 Then
        aeExists = False
        Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class aegit_expClass"
        If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class aegit_expClass"
        aeExists = False
    End If
    Resume PROC_EXIT

End Function

Private Function GetType(ByVal Value As Long) As String
' Ref: http://bytes.com/topic/access/answers/557780-getting-string-name-enum

    On Error GoTo 0
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
    On Error GoTo 0
    Dim bln As Boolean
    bln = FieldLookupControlTypeList
End Sub

Private Function FieldLookupControlTypeList(Optional ByVal varDebug As Variant) As Boolean
' Ref: http://support.microsoft.com/kb/304274
' Ref: http://msdn.microsoft.com/en-us/library/office/bb225848(v=office.12).aspx
' 106 - acCheckBox, 109 - acTextBox, 110 - acListBox, 111 - acComboBox

    On Error GoTo PROC_ERR

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
        If Left$(tbl.Name, 4) <> "MSys" And Left$(tbl.Name, 3) <> "zzz" _
            And Left$(tbl.Name, 1) <> "~" Then
            Print #fle, tbl.Name
            For Each fld In tbl.Fields
                intAllFieldsCount = intAllFieldsCount + 1
                lng = fld.Properties("DisplayControl").Value
                Print #fle, , fld.Name, lng, GetType(lng)
                Select Case lng
                    Case acCheckBox
                        intChk = intChk + 1
                        strChkTbl = tbl.Name
                        strChkFld = fld.Name
                    Case acTextBox
                        intTxt = intTxt + 1
                    Case acListBox
                        intLst = intLst + 1
                    Case acComboBox
                        intCbo = intCbo + 1
                    Case Else
                        intElse = intElse + 1
                End Select
            Next fld
        End If
    Next tbl

    If Not IsMissing(varDebug) Then
        Debug.Print "Count of Check box = " & intChk
        Debug.Print "Count of Text box  = " & intTxt
        Debug.Print "Count of List box  = " & intLst
        Debug.Print "Count of Combo box = " & intCbo
        Debug.Print "Count of Else      = " & intElse
        Debug.Print "Count of Display Controls = " & intChk + intTxt + intLst + intCbo
        Debug.Print "Count of All Fields = " & intAllFieldsCount - intElse
        'Debug.Print "Table with check box is " & strChkTbl
        'Debug.Print "Field with check box is " & strChkFld
    End If

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
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FieldLookupControlTypeList of Class aegit_expClass", vbCritical, "Error"
    Resume PROC_EXIT

End Function

Private Sub OutputListOfCommandBarIDs(ByVal strOutputFile As String, Optional ByVal varDebug As Variant)
' Programming Office Commandbars - get the ID of a CommandBarControl
' Ref: http://blogs.msdn.com/b/guowu/archive/2004/09/06/225963.aspx
' Ref: http://www.vbforums.com/showthread.php?392954-How-do-I-Find-control-IDs-in-Visual-Basic-for-Applications-for-office-2003

    On Error GoTo PROC_ERR

    Dim CBR As Object       ' CommandBar
    Set CBR = Application.CommandBars
    Dim CBTN As Object      ' CommandBarButton
    Set CBTN = Application.CommandBars.FindControls
    Dim fle As Integer
    Dim lng As Long
    Dim strPathFileName As String
    Dim strExtension As String

    strPathFileName = aegitSourceFolder & strOutputFile
    strExtension = ".sorted.txt"

    fle = FreeFile()
    Open strPathFileName For Output As #fle

    On Error Resume Next

    For Each CBR In Application.CommandBars
        For Each CBTN In CBR.Controls
            If Not IsMissing(varDebug) Then Debug.Print CBR.Name & ": " & CBTN.id & " - " & CBTN.Caption
            Print #fle, CBR.Name & ": " & CBTN.id & " - " & CBTN.Caption
        Next
    Next
    Close fle

    ' Sort the file
    If Not IsMissing(varDebug) Then
        Debug.Print "strPathFileName=" & strPathFileName
        Debug.Print "strExtension=" & strExtension
    End If

    lng = MySortIt(strPathFileName, strExtension, "Unicode")
    'Stop

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfCommandBarIDs of Class aegit_expClass", vbCritical, "Error"
    Resume PROC_EXIT

End Sub

Private Function MySortIt(ByVal strFPName As String, ByVal strExtension As String, _
                            Optional ByVal varUnicode As Variant) As Long
' Ref: http://support.microsoft.com/kb/150700
' Ref: http://www.xtremevbtalk.com/showthread.php?t=291063
' Ref: http://www.ozgrid.com/forum/showthread.php?t=167349

    On Error GoTo PROC_ERR

    Dim strVar As Variant
    Dim lngLine As Long
    Dim theCount As Long

    Dim arrayIn As Object
    Set arrayIn = CreateObject("System.Collections.ArrayList")
    Dim arrayOut() As Variant

    Close #1
    Close #2
    Open strFPName For Input As #1
    Open strFPName & strExtension For Output As #2

    With arrayIn
        Do Until EOF(1)
            Line Input #1, strVar
            .Add Trim$(CStr(strVar))
            .Add Trim$(strVar)
        Loop
        .Sort
        theCount = .Count
        'Debug.Print .Count
        arrayOut = arrayIn.ToArray
        For lngLine = LBound(arrayOut) To UBound(arrayOut)
            If IsMissing(varUnicode) Then
                Print #2, arrayIn(lngLine)
            End If
        Next
    End With

    If Not IsMissing(varUnicode) Then
        OutputMyUnicode strFPName & strExtension, arrayOut()
    End If

    MySortIt = theCount

    Close #1
    Close #2
    Set arrayIn = Nothing
    'Debug.Print "Done!"

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure MySortIt of Class aegit_expClass", vbCritical, "Error"
    Resume PROC_EXIT

End Function

Public Sub OutputMyUnicode(ByRef strPathFileName As String, _
                            ByVal arrUnicode As Variant)
' Ref: http://www.experts-exchange.com/Database/MS_Access/Q_26282187.html
' Ref: http://accessblog.net/2007/06/how-to-write-out-unicode-text-files-in.html

    On Error GoTo PROC_ERR

    Dim i As Integer
    Dim MyStream As Object
    Set MyStream = CreateObject("ADODB.Stream")
    ' `It is summer in Geneva`, said Yu Zhou.
    ' strUnicode = "«" & Chr(160) & "C'est l'été à Genève" & Chr(160) & "»," _
            & " said " & ChrW(20446) & ChrW(-32225) & "."
    ' Ref: http://msdn.microsoft.com/en-us/library/windows/desktop/ms675277(v=vs.85).aspx
    '
    Debug.Print "strPathFileName=" & strPathFileName
    Debug.Print "arrUnicode(0)=" & arrUnicode(0)
    Debug.Print "arrUnicode(1)=" & arrUnicode(1)
    arrUnicode(2) = "«" & Chr$(160) & "C'est l'été à Genève" & Chr$(160) & "»," _
                  & " said " & ChrW(20446) & ChrW(-32225) & "."
    Debug.Print "arrUnicode(2)=" & arrUnicode(2)
    Dim mystrPathFileName As String
    With MyStream
        .Type = 2    ' adTypeText
        .Charset = "Unicode"
        .Open
        For i = LBound(arrUnicode) To UBound(arrUnicode)
            .WriteText arrUnicode(i) ' The foreign unicode text
            'Debug.Print , i, "arrUnicode(i)=" & arrUnicode(i)
        Next i
        Debug.Print "aestrSourceLocation=" & aestrSourceLocation
        mystrPathFileName = aestrSourceLocation & "TEST_OutputListOfCommandBarIDs.txt"
        .SaveToFile mystrPathFileName, 2            ' adSaveCreateOverWrite
        '.SaveToFile "C:\TEMP\TestItFile.txt", 2            ' adSaveCreateOverWrite
        .Close
    End With
    'Stop

    Set MyStream = Nothing

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputMyUnicode of Class aegit_expClass", vbCritical, "Error"
    Resume PROC_EXIT

End Sub

Private Function OutputListOfContainers(ByVal strTheFileName As String, Optional ByVal varDebug As Variant) As Boolean
' Ref: http://www.susandoreydesigns.com/software/AccessVBATechniques.pdf
' Ref: http://msdn.microsoft.com/en-us/library/office/bb177484(v=office.12).aspx

    Dim dbs As DAO.Database
    Dim conItem As DAO.Container
    Dim prpLoop As DAO.Property
    Dim strFile As String
    Dim lngFileNum As Long

    On Error GoTo PROC_ERR

    OutputListOfContainers = True

    Debug.Print "OutputListOfContainers"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to OutputListOfContainers"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to OutputListOfContainers"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    Set dbs = CurrentDb
    lngFileNum = FreeFile()

    strFile = aestrSourceLocation & strTheFileName

    If Dir$(strFile) <> vbNullString Then
        ' The file exists
        If Not FileLocked(strFile) Then
            KillProperly (strFile)
        End If
        Open strFile For Append As lngFileNum
    Else
        If Not FileLocked(strFile) Then
            Open strFile For Append As lngFileNum
        End If
    End If

    With dbs
        ' Enumerate Containers collection.
        For Each conItem In .Containers
            If Not IsMissing(varDebug) Then
                Debug.Print "Properties of " & conItem.Name & " container", lngFileNum, strFile
                WriteStringToFile lngFileNum, "Properties of " & conItem.Name & " container", strFile, varDebug
            Else
                WriteStringToFile lngFileNum, "Properties of " & conItem.Name & " container", strFile
            End If

            ' Enumerate Properties collection of each Container object.
            For Each prpLoop In conItem.Properties
                If Not IsMissing(varDebug) Then
                    Debug.Print , lngFileNum, prpLoop.Name & " = " & prpLoop
                    WriteStringToFile lngFileNum, "  " & prpLoop.Name & " = " & prpLoop, strFile, varDebug
                Else
                    WriteStringToFile lngFileNum, "  " & prpLoop.Name & " = " & prpLoop, strFile
                End If
            Next prpLoop
        Next conItem
        .Close
    End With

PROC_EXIT:
    Close lngFileNum
    Set prpLoop = Nothing
    Set conItem = Nothing
    Set dbs = Nothing
'Stop
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 3358   ' Cannot open the Microsoft Access database workgroup information file
            'Debug.Print "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfContainers of Class aegit_expClass"
            Resume Next
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfContainers of Class aegit_expClass"
            Resume Next
    End Select
    OutputListOfContainers = False
    Resume Next

End Function

Public Sub OutputAllContainerProperties(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    If Not IsMissing(varDebug) Then
        Debug.Print "Container information for properties of saved Databases"
        ListAllContainerProperties "Databases", varDebug
        Debug.Print "Container information for properties of saved Tables and Queries"
        ListAllContainerProperties "Tables", varDebug
        'Stop
        Debug.Print "Container information for properties of saved Relationships"
        ListAllContainerProperties "Relationships", varDebug
    Else
        ListAllContainerProperties "Databases"
        ListAllContainerProperties "Tables"
        ListAllContainerProperties "Relationships"
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputAllContainerProperties of Class aegit_expClass"
    'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputAllContainerProperties of Class aegit_expClass"
    Resume PROC_EXIT

End Sub

Private Function fListGUID(ByVal strTableName As String) As String
' Ref: http://stackoverflow.com/questions/8237914/how-to-get-the-guid-of-a-table-in-microsoft-access
' e.g. ?fListGUID("tblThisTableHasSomeReallyLongNameButItCouldBeMuchLonger")

    On Error GoTo 0
    Dim i As Integer
    Dim arrGUID8() As Byte
    Dim strArrGUID8(8) As String
    Dim strGuid As String

    strGuid = vbNullString
    arrGUID8 = CurrentDb.TableDefs(strTableName).Properties("GUID").Value
    For i = 1 To 8
        If Len(Hex$(arrGUID8(i))) = 1 Then
            strArrGUID8(i) = "0" & Hex$(arrGUID8(i))
        Else
            strArrGUID8(i) = Hex$(arrGUID8(i))
        End If
    Next

    For i = 1 To 8
        strGuid = strGuid & strArrGUID8(i) & "-"
    Next
    fListGUID = Left$(strGuid, 23)

End Function

Private Sub ListAllContainerProperties(ByVal strContainer As String, Optional ByVal varDebug As Variant)
' Ref: http://www.dbforums.com/microsoft-access/1620765-read-ms-access-table-properties-using-vba.html
' Ref: http://msdn.microsoft.com/en-us/library/office/aa139941(v=office.10).aspx
    
    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim obj As Object
    Dim prp As DAO.Property
    Dim doc As DAO.Document
    Dim fle As Integer

    Set dbs = Application.CurrentDb
    Set obj = dbs.Containers(strContainer)

    fle = FreeFile()
    Open aegitSourceFolder & "\OutputContainer" & strContainer & "Properties.txt" For Output As #fle

    ' Ref: http://stackoverflow.com/questions/16642362/how-to-get-the-following-code-to-continue-on-error
    For Each doc In obj.Documents
        If Left$(doc.Name, 4) <> "MSys" And Left$(doc.Name, 3) <> "zzz" _
            And Left$(doc.Name, 1) <> "~" Then
            If Not IsMissing(varDebug) Then Debug.Print ">>>" & doc.Name
            Print #fle, ">>>" & doc.Name
            For Each prp In doc.Properties
                On Error Resume Next
                If prp.Name = "GUID" And strContainer = "tables" Then
                    Print #fle, , prp.Name, fListGUID(doc.Name)
                    If Not IsMissing(varDebug) Then Debug.Print , prp.Name, fListGUID(doc.Name)
                ElseIf prp.Name = "DOL" Then
                    Print #fle, , prp.Name, "Track name AutoCorrect info is ON!"
                    If Not IsMissing(varDebug) Then Debug.Print prp.Name, "Track name AutoCorrect info is ON!"
                ElseIf prp.Name = "NameMap" Then
                    Print #fle, , prp.Name, "Track name AutoCorrect info is ON!"
                    If Not IsMissing(varDebug) Then Debug.Print , prp.Name, "Track name AutoCorrect info is ON!"
                Else
                    If prp.Name = "DateCreated" Then
                        Print #fle, , prp.Name, "DateCreated"
                    ElseIf prp.Name = "LastUpdated" Then
                        Print #fle, , prp.Name, "LastUpdated"
                    Else
                        Print #fle, , prp.Name, prp.Value
                    End If
                    If Not IsMissing(varDebug) Then
                        Debug.Print , prp.Name, prp.Value
                        If prp.Name = "DateCreated" Then
                            Debug.Print , "=>", prp.Name, "DateCreated"
                        ElseIf prp.Name = "LastUpdated" Then
                            Debug.Print , "=>", prp.Name, "LastUpdated"
                        End If
                    End If
                End If
                On Error GoTo 0
            Next
        End If
    Next

    Set obj = Nothing
    Set dbs = Nothing
    Close fle

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListAllContainerProperties of Class aegit_expClass"
    'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListAllContainerProperties of Class aegit_expClass"
    Resume PROC_EXIT

End Sub

Public Sub PrettyXML(ByVal strPathFileName As String, Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    ' Beautify XML in VBA with MSXML6 only
    ' Ref: http://social.msdn.microsoft.com/Forums/en-US/409601d4-ca95-448a-aafc-aa0ee1ad67cd/beautify-xml-in-vba-with-msxml6-only?forum=xmlandnetfx
    Dim objXMLStyleSheet As Object
    Dim strXMLStyleSheet As String
    Dim objXMLDOMDoc As Object

    Dim fle As Integer
    fle = FreeFile()

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
        Debug.Print "PrettyXML: Some Error..."
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

    Dim strXMLResDoc As Variant
    Set strXMLResDoc = CreateObject("Msxml2.DOMDocument.6.0")

    objXMLDOMDoc.transformNodeToObject objXMLStyleSheet, strXMLResDoc
    strXMLResDoc = strXMLResDoc.XML
    strXMLResDoc = Replace(strXMLResDoc, vbTab, Chr$(32) & Chr$(32), , , vbBinaryCompare)
    If Not IsMissing(varDebug) Then
        Debug.Print , "Pretty XML Sample Output"
        Debug.Print strXMLResDoc
    End If

    ' Test for relative path
    Dim strTestPath As String
    strTestPath = strPathFileName
    If Left(strPathFileName, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid(strPathFileName, 2, Len(strPathFileName) - 1)
        strPathFileName = strTestPath
        'Debug.Print , "strPathFileName = " & strPathFileName, "PrettyXML"
        'Stop
    End If

    ' Rewrite the file as pretty xml
    'Debug.Print "PrettyXML strPathFileName = " & strPathFileName
    Open strPathFileName For Output As #fle
    Print #fle, strXMLResDoc
    Close #fle

    Set objXMLDOMDoc = Nothing
    Set objXMLStyleSheet = Nothing

PROC_EXIT:
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case 9999
            'Debug.Print "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrettyXML of Class aegit_expClass"
            Resume Next
        Case Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrettyXML of Class aegit_expClass"
        'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrettyXML of Class aegit_expClass"
    End Select
    Resume PROC_EXIT

End Sub

Private Sub OutputTheTableDataAsXML(ByRef avarTableNames() As Variant, Optional ByVal varDebug As Variant)
' Ref: http://wiki.lessthandot.com/index.php/Output_Access_/_Jet_to_XML
' Ref: http://msdn.microsoft.com/en-us/library/office/aa164887(v=office.10).aspx

    Dim i As Integer

    On Error GoTo PROC_ERR

    Const adOpenStatic As Integer = 3
    Const adLockOptimistic As Integer = 3
    Const adPersistXML As Integer = 1

    Dim strFileName As String
    Dim strSQL As String

    Dim cnn As Object
    Set cnn = CurrentProject.Connection
    Dim rst As Object
    Set rst = CreateObject("ADODB.Recordset")

    For i = 0 To UBound(avarTableNames)
        If aeExists("Tables", avarTableNames(i)) Then
            strSQL = "Select * from " & avarTableNames(i)
            'MsgBox strSQL, vbInformation, "OutputTheTableDataAsXML"
            If Not IsMissing(varDebug) Then Debug.Print i, "avarTableNames", avarTableNames(i)
            rst.Open strSQL, cnn, adOpenStatic, adLockOptimistic

            strFileName = aestrXMLLocation & avarTableNames(i) & ".xml"

            If aegitSetup Then
                If Not IsMissing(varDebug) Then Debug.Print "aegitSetup=True aestrXMLLocation=" & aestrXMLLocation
            Else
                If Not IsMissing(varDebug) Then Debug.Print "aegitSetup=False aestrXMLLocation=" & aestrXMLLocation
            End If

            If Not rst.EOF Then
                rst.MoveFirst
                rst.Save strFileName, adPersistXML
            End If

            If Not IsMissing(varDebug) Then
                PrettyXML strFileName, varDebug
            Else
                PrettyXML strFileName
            End If
            rst.Close
        End If
    Next

    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing

PROC_EXIT:
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case 58     ' File already exists
            'Debug.Print "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrettyXML of Class aegit_expClass"
            Resume PROC_EXIT
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ")" & vbCrLf & _
            "strFileName = " & strFileName & vbCrLf & "in procedure OutputTheTableDataAsXML of Class aegit_expClass", vbExclamation
            'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputTheTableDataAsXML of Class aegit_expClass"
    End Select
    Resume PROC_EXIT

End Sub

Public Sub OutputPrinterInfo(Optional ByVal varDebug As Variant)
' Ref: http://msdn.microsoft.com/en-us/library/office/aa139946(v=office.10).aspx
' Ref: http://answers.microsoft.com/en-us/office/forum/office_2010-access/how-do-i-change-default-printers-in-vba/d046a937-6548-4d2b-9517-7f622e2cfed2

    On Error GoTo PROC_ERR

    Dim prt As Printer
    Dim prtCount As Integer
    Dim i As Integer
    Dim fle As Integer

    If Not mblnOutputPrinterInfo Then Exit Sub
    
    fle = FreeFile()
    Open aegitSourceFolder & "\" & aePrnterInfo For Output As #fle

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
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case 9
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputPrinterInfo of Class aegit_expClass"
            Resume Next
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputPrinterInfo of Class aegit_expClass"
            Resume Next
    End Select

End Sub

Private Sub OutputTableDataMacros(Optional ByVal varDebug As Variant)
' Ref: http://stackoverflow.com/questions/9206153/how-to-export-access-2010-data-macros
' ====================================================================
' Author:   Peter F. Ennis
' Date:     February 16, 2014
' ====================================================================

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strFile As String

    On Error GoTo PROC_ERR

    Set dbs = CurrentDb()
    For Each tdf In CurrentDb.TableDefs
        If Not LinkedTable(tdf.Name) Or _
                Not (Left$(tdf.Name, 4) = "MSys" _
                Or Left$(tdf.Name, 4) = "~TMP" _
                Or Left$(tdf.Name, 3) = "zzz") Then
            strFile = aestrXMLLocation & "tables_" & tdf.Name & "_DataMacro.xml"
            'Debug.Print "OutputTableDataMacros: aestrXMLLocation = " & aestrXMLLocation
            'Debug.Print "OutputTableDataMacros: strFile = " & strFile
            SaveAsText acTableDataMacro, tdf.Name, strFile
2220:
            If Not IsMissing(varDebug) Then
                Debug.Print "OutputTableDataMacros:", tdf.Name, aestrXMLLocation, strFile
                PrettyXML strFile, varDebug
            Else
                PrettyXML strFile
            End If
        End If
NextTdf:
    Next tdf

PROC_EXIT:
    Set tdf = Nothing
    Set dbs = Nothing
    Exit Sub

PROC_ERR:
    If Err = 2950 Then ' Reserved Error
        Resume NextTdf
    ElseIf Err = 2220 Then
        Resume 2220
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputTableDataMacros of Class aegit_expClass"
    Resume PROC_EXIT

End Sub

Private Sub OutputTableDataAsFormattedText(ByVal strTblName As String, Optional ByVal varDebug As Variant)
' Ref: http://bytes.com/topic/access/answers/856136-access-2007-vba-select-external-data-ribbon

    On Error GoTo 0
    Dim strPathFileName As String
    strPathFileName = aestrSourceLocation & strTblName & "_FormattedData.txt"
    If Not IsMissing(varDebug) Then
        Debug.Print , strPathFileName
    Else
    End If
'    AcFormat can be one of these AcFormat constants.
'    acFormatASP
'    acFormatDAP
'    acFormatHTML
'    acFormatIIS
'    acFormatRTF
'    acFormatSNP
'    acFormatTXT
'    acFormatXLS
    DoCmd.OutputTo acOutputTable, strTblName, acFormatTXT, aestrSourceLocation & strTblName & "_FormattedData.txt"
    DoCmd.TransferText acExportDelim, , strTblName, aestrSourceLocation & strTblName & "_TransferText.txt", True

End Sub

Private Sub CreateFormReportTextFile(ByVal strFileIn As String, ByVal strFileOut As String, Optional ByVal varDebug As Variant)
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
    Dim strIn As String
    Dim i As Integer

    fleIn = FreeFile()
    Open strFileIn For Input As #fleIn

    fleOut = FreeFile()
    Open strFileOut For Output As #fleOut

    If Not IsMissing(varDebug) Then Debug.Print "fleIn=" & fleIn, "fleOut=" & fleOut

    i = 0
    Do While Not EOF(fleIn)
        i = i + 1
        Line Input #fleIn, strIn
        If Left$(strIn, Len("Checksum =")) = "Checksum =" Then
            Exit Do
        Else
            If Not IsMissing(varDebug) Then Debug.Print i, strIn
            Print #fleOut, strIn
        End If
    Loop
    Do While Not EOF(fleIn)
        i = i + 1
        Line Input #fleIn, strIn
NextIteration:
        If FoundKeywordInLine(strIn) Then
            If Not IsMissing(varDebug) Then Debug.Print i & ">", strIn
            Print #fleOut, strIn
            Do While Not EOF(fleIn)
                i = i + 1
                Line Input #fleIn, strIn
                If Not FoundKeywordInLine(strIn, "End") Then
                    'Debug.Print "Not Found!!!", i
                    GoTo SearchForEnd
                Else
                    If Not IsMissing(varDebug) Then Debug.Print i & ">", "Found End!!!"
                    Print #fleOut, strIn
                    i = i + 1
                    Line Input #fleIn, strIn
                    If Not IsMissing(varDebug) Then Debug.Print i & ":", strIn
                    'Stop
                    GoTo NextIteration
                End If
SearchForEnd:
            Loop
        Else
            Print #fleOut, strIn
            If Not IsMissing(varDebug) Then Debug.Print i, strIn
        End If
    Loop

    Close fleIn
    Close fleOut

End Sub

Private Function FoundKeywordInLine(ByVal strLine As String, Optional ByVal varEnd As Variant) As Boolean

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

Public Sub OutputCatalogUserCreatedObjects(Optional ByVal varDebug As Variant)
' Ref: http://blogannath.blogspot.com/2010/03/microsoft-access-tips-tricks-working.html#ixzz3WCBJcxwc
' Ref: http://stackoverflow.com/questions/5286620/saving-a-query-via-access-vba-code

    On Error GoTo PROC_ERR
    
    Dim strSQL As String
    Const MY_QUERY_NAME = "zzzqryCatalogUserCreatedObjects"
    
    Dim strPathFileName As String
    strPathFileName = aestrSourceLocation & aeCatalogObj
    If Not IsMissing(varDebug) Then
        Debug.Print , strPathFileName
    Else
    End If

    'strSQL = strSQL & vbCrLf & "MSysObjects.Name, MSysObjects.DateCreate, MSysObjects.DateUpdate "
    
'    strSQL = "SELECT IIf(type = 1,""Table"", IIf(type = 6, ""Linked Table"", "
'    strSQL = strSQL & vbCrLf & "IIf(type = 5,""Query"", IIf(type = -32768,""Form"", "
'    strSQL = strSQL & vbCrLf & "IIf(type = -32764,""Report"", IIf(type=-32766,""Module"", "
'    strSQL = strSQL & vbCrLf & "IIf(type = -32761,""Module"", ""Unknown""))))))) as [Object Type], "
'    strSQL = strSQL & vbCrLf & "MSysObjects.Name, MSysObjects.DateCreate "
'    strSQL = strSQL & vbCrLf & "FROM MSysObjects "
'    strSQL = strSQL & vbCrLf & "WHERE Type IN (1, 5, 6, -32768, -32764, -32766, -32761) "
'    strSQL = strSQL & vbCrLf & "AND Left(Name, 4) <> ""MSys"" AND Left(Name, 1) <> ""~"" "
'    strSQL = strSQL & vbCrLf & "ORDER BY IIf(type=1,""Table"",IIf(type=6,""Linked Table"",IIf(type=5,""Query"",IIf(type=-32768,""Form"",IIf(type=-32764,""Report"",IIf(type=-32766,""Module"",IIf(type=-32761,""Module"",""Unknown""))))))), MSysObjects.Name;"

    ' Ref: https://support.office.com/en-za/article/FormatDateTime-Function-aef62949-f957-4ba4-94ff-ace14be4f1ca
    ' Format DateCreate as short date, vbShortDate = 2
    'SELECT IIf(type=1,"Table",IIf(type=6,"Linked Table",IIf(type=5,"Query",IIf(type=-32768,"Form",IIf(type=-32764,"Report",IIf(type=-32766,"Module",IIf(type=-32761,"Module","Unknown"))))))) AS [Object Type], MSysObjects.Name, FormatDateTime([DateCreate],2) AS DateCreated
    'FROM MSysObjects
    'WHERE (((MSysObjects.[Type]) In (1,5,6,-32768,-32764,-32766,-32761)) AND ((Left([Name],4))<>"MSys") AND ((Left([Name],1))<>"~"))
    'ORDER BY IIf(type=1,"Table",IIf(type=6,"Linked Table",IIf(type=5,"Query",IIf(type=-32768,"Form",IIf(type=-32764,"Report",IIf(type=-32766,"Module",IIf(type=-32761,"Module","Unknown"))))))), MSysObjects.Name;

    strSQL = "SELECT IIf(type = 1,""Table"", IIf(type = 6, ""Linked Table"", "
    strSQL = strSQL & vbCrLf & "IIf(type = 5,""Query"", IIf(type = -32768,""Form"", "
    strSQL = strSQL & vbCrLf & "IIf(type = -32764,""Report"", IIf(type=-32766,""Module"", "
    strSQL = strSQL & vbCrLf & "IIf(type = -32761,""Module"", ""Unknown""))))))) as [Object Type], "
    strSQL = strSQL & vbCrLf & "MSysObjects.Name, ""DateCreated"" AS DateCreated "
    'strSQL = strSQL & vbCrLf & "MSysObjects.Name, FormatDateTime([DateCreate],2) AS DateCreated "
    strSQL = strSQL & vbCrLf & "FROM MSysObjects "
    strSQL = strSQL & vbCrLf & "WHERE Type IN (1, 5, 6, -32768, -32764, -32766, -32761) "
    strSQL = strSQL & vbCrLf & "AND Left(Name, 4) <> ""MSys"" AND Left(Name, 1) <> ""~"" "
    strSQL = strSQL & vbCrLf & "ORDER BY IIf(type=1,""Table"",IIf(type=6,""Linked Table"",IIf(type=5,""Query"",IIf(type=-32768,""Form"",IIf(type=-32764,""Report"",IIf(type=-32766,""Module"",IIf(type=-32761,""Module"",""Unknown""))))))), MSysObjects.Name;"

    'Debug.Print strSQL

    ' Using a query name and sql string, if the query does not exist, ...
    If IsNull(DLookup("Name", "MsysObjects", "Name='" & MY_QUERY_NAME & "'")) Then
        ' create it ...
        CurrentDb.CreateQueryDef MY_QUERY_NAME, strSQL
    Else
        ' other wise, update the sql
        CurrentDb.QueryDefs(MY_QUERY_NAME).SQL = strSQL
    End If

    'DoCmd.OpenQuery MY_QUERY_NAME
e3167:
    DoCmd.TransferText acExportDelim, , MY_QUERY_NAME, strPathFileName

PROC_EXIT:
    Exit Sub

PROC_ERR:
    If Err = 3167 Then          ' Record is deleted
        'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputCatalogUserCreatedObjects of Class aegit_expClass"
        Resume e3167
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputCatalogUserCreatedObjects of Class aegit_expClass"
    End If
    'Stop
    Resume PROC_EXIT

End Sub

' ==================================================
' Global Error Handler Routines
' Ref: http://msdn.microsoft.com/en-us/library/office/ee358847(v=office.12).aspx#odc_ac2007_ta_ErrorHandlingAndDebuggingTipsForAccessVBAndVBA_WritingCodeForDebugging
' ==================================================

Private Sub ResetWorkspace()
    Dim intCounter As Integer

    On Error Resume Next

    Application.MenuBar = vbNullString
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

Private Sub WriteStringToFile(ByVal lngFileNum As Long, ByVal strTheString As String, _
                                ByVal strTheAbsoluteFileName As String, Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    If Not IsMissing(varDebug) Then
        Debug.Print "WriteStringToFile"
        Debug.Print , lngFileNum, strTheString, strTheAbsoluteFileName, varDebug
    End If

ERR55:
    Close lngFileNum
    Open strTheAbsoluteFileName For Append Access Write Lock Write As lngFileNum
    Print #lngFileNum, strTheString
    Close lngFileNum

PROC_EXIT:
    Exit Sub

PROC_ERR:
    If Err = 55 Then ' File already open
        'Debug.Print "Here@Err=55", strTheString, lngFileNum, strTheAbsoluteFileName
        If Not IsMissing(varDebug) Then Debug.Print "Err=55", strTheString, lngFileNum, strTheAbsoluteFileName
        Err.Clear
        Resume ERR55
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure WriteStringToFile of Class aegit_expClass"
        Resume Next
    End If

End Sub