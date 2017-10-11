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

Private Const aegit_expVERSION As String = "2.0.0"
Private Const aegit_expVERSION_DATE As String = "October 10, 2017"
'Private Const aeAPP_NAME As String = "aegit_exp"
Private Const mblnOutputPrinterInfo As Boolean = False
' If mblnUTF16 is True the form txt exported files will be UTF-16 Windows format
' If mblnUTF16 is False the BOM marker will be stripped and files will be ANSI
Private Const mblnUTF16 As Boolean = False

Private mstrToParse As String
Private Const mTABLE As String = "CREATE TABLE ["
Private Const mPRIMARYKEY As String = "CREATE UNIQUE INDEX ["       'PrimaryKey] ON ["
Private Const mINDEX As String = "CREATE INDEX ["

Private Enum SizeStringSide
    TextLeft = 1
    TextRight = 2
End Enum

Private Type myExclusions
    excludeOne As String
    excludeTwo As String
    excludeThree As String
End Type

Private Type mySetupType
    SourceFolder As String
    SourceFolderBe As String
    XMLFolder As String
    XMLFolderBe As String
    XMLDataFolder As String
    XMLDataFolderBe As String
    ImportFolder As String
    UseImportFolder As Boolean
End Type

Private Type myExportType                           ' Initialize defaults as:
    ExportAll As Boolean                            ' True
    ExportCodeAndObjects As Boolean                 ' True
    ExportModuleCodeOnly As Boolean                 ' True
    EXPERIMENTAL_ExportQAT As Boolean               ' False
    EXPERIMENTAL_ExportCBID As Boolean              ' False
    ExportNoODBCTablesInfo As Boolean               ' True, default does not export info about ODBC linked tables
    EXPERIMENTAL_ExportCreateDbScript As Boolean    ' False,
End Type

Private myExclude As myExclusions
Private pExclude As Boolean
Private mblnIgnore As Boolean
Private mblnResult As Boolean
Private mstrTheSourceLocation As String

Private aegitSetup As Boolean
Private aegitType As mySetupType
Private aegitExport As myExportType
Private aegitFrontEndApp As Boolean
Private aegitSourceFolder As String
Private aegitSourceFolderBe As String
Private aegitXMLFolder As String
Private aegitXMLFolderBe As String
Private aegitXMLDataFolder As String
Private aegitXMLDataFolderBe As String
Private aegitTextEncoding As String
Private aegitDataXML() As Variant
Private aegitExportDataToXML As Boolean
Private aestrSourceLocation As String
Private aestrSourceLocationBe As String
Private aestrXMLLocation As String
Private aestrXMLLocationBe As String
Private aestrXMLDataLocation As String
Private aestrXMLDataLocationBe As String
Private aestrLFN As String                      ' Longest Field Name
Private aestrLFNTN As String
Private aeintFNLen As Long
Private aestrLFT As String                      ' Longest Field Type
Private aeintFTLen As Long                      ' Field Type Length
Private Const aeintFSize As Long = 4
Private aeintFDLen As Long
Private aestrLFD As String
Private aestrBackEndDbOne As String
Private aeListOfTables() As Variant
'
Private Const DebugPrintInitialize As Boolean = False
'Private aestrPassword As String
Private Const aestr4 As String = "    "
Private Const aeAppCmbrIds As String = "OutputListOfCommandBarIDs.txt"
Private Const aeAppHiddQry As String = "OutputListOfAllHiddenQueries.txt"
Private Const aeAppListCnt As String = "OutputListOfContainers.txt"
Private Const aeAppListFrm As String = "OutputListOfForms.txt"
Private Const aeAppListMac As String = "OutputListOfMacros.txt"
Private Const aeAppListMod As String = "OutputListOfModules.txt"
Private Const aeAppListPrp As String = "OutputListOfApplicationProperties.txt"
Private Const aeAppListQAT As String = "OutputQAT"  ' Will be saved with file extension .exportedUI
Private Const aeAppListRpt As String = "OutputListOfReports.txt"
Private Const aeAppListTbl As String = "OutputListOfTables.txt"
Private Const aeAppOptions As String = "OutputListOfAccessApplicationOptions.txt"
Private Const aeCatalogObj As String = "OutputCatalogUserCreatedObjects.txt"
Private Const aeFLkCtrFile As String = "OutputFieldLookupControlTypeList.txt"
Private Const aeIndexLists As String = "OutputListOfIndexes.txt"
Private Const aeLoveSchema As String = "OutputLovefieldSchema.txt"
Private Const aePrnterInfo As String = "OutputPrinterInfo.txt"
Private Const aePrpTxtFile As String = "OutputPropertiesBuiltIn.txt"
Private Const aeRefTxtFile As String = "OutputReferencesSetup.txt"
Private Const aeRelTxtFile As String = "OutputRelationsSetup.txt"
Private Const aeSchemaFile As String = "OutputSchemaFile.txt"
Private Const aeSqlTxtFile As String = "OutputSqlCodeForQueries.txt"
Private Const aeTblTxtFile As String = "OutputFieldsForTables.txt"
'
' aeLogger begin variables
Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
Private Type aeLogger
    blnNoTrace As Boolean
    blnNoEnd As Boolean
    blnNoPrint As Boolean
    blnNoTimer As Boolean
End Type

Private lngIndent As Long
Private mlngStartTime As Long
Private mlngEndTime As Long
Private aeLog As aeLogger
' aeLogger end variables

Private Sub Class_Initialize()
    ' Ref: http://www.cadalyst.com/cad/autocad/programming-with-class-part-1-5050
    ' Ref: http://www.bigresource.com/Tracker/Track-vb-cyJ1aJEyKj/
    ' Ref: http://stackoverflow.com/questions/1731052/is-there-a-way-to-overload-the-constructor-initialize-procedure-for-a-class-in

    On Error GoTo PROC_ERR
    
    With aeLog
        .blnNoEnd = False
        .blnNoPrint = False
        .blnNoTimer = False
        .blnNoTrace = False
    End With

    aeBeginLogging "Class_Initialize"
    Application.SetOption "Show Hidden Objects", True

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
    aegitXMLFolder = "default"
    aegitXMLFolderBe = "default"
    aegitXMLDataFolder = "default"
    aegitXMLDataFolderBe = "default"
    aestrBackEndDbOne = "default"             ' default for aegit is no back end database
    ReDim Preserve aegitDataXML(0 To 0)
    If Application.VBE.ActiveVBProject.Name = "aegit" Then
        aegitDataXML(0) = "aetlkpStates"
    End If
    aegitExportDataToXML = True
    aegitType.SourceFolder = "C:\ae\aegit\aerc\src\"
    aegitType.SourceFolderBe = "C:\ae\aegit\aerc\srcbe\"
    aegitType.XMLFolder = "C:\ae\aegit\aerc\src\xml\"
    aegitType.XMLFolderBe = "C:\ae\aegit\aerc\srcbe\xml\"
    aegitType.XMLDataFolder = "C:\ae\aegit\aerc\src\xmldata\"
    aegitType.XMLDataFolderBe = "C:\ae\aegit\aerc\srcbe\xmldata\"

    Const aeintLTN As Integer = 11           ' Set a minimum default
    aeintFNLen = 4          ' Set a minimum default
    aeintFTLen = 4          ' Set a minimum default
    aeintFDLen = 4          ' Set a minimum default

    With aegitExport
        .ExportAll = True
        .ExportCodeAndObjects = True
        .ExportModuleCodeOnly = True
        .EXPERIMENTAL_ExportQAT = False
        .EXPERIMENTAL_ExportCBID = False
        .ExportNoODBCTablesInfo = True
    End With

    pExclude = True             ' Default setting is not to export associated aegit_exp files
    aegitFrontEndApp = True     ' Default is a front end app

    Debug.Print "Class_Initialize"
    If DebugPrintInitialize Then
        Debug.Print , "Default for aegitFrontEndApp = " & aegitFrontEndApp
        Debug.Print , "Default for aegitType.SourceFolder = " & aegitType.SourceFolder
        Debug.Print , "Default for aegitType.SourceFolderBe = " & aegitType.SourceFolderBe
        Debug.Print , "Default for aegitType.XMLFolder = " & aegitType.XMLFolder
        Debug.Print , "Default for aegitType.XMLFolderBe = " & aegitType.XMLFolderBe
        Debug.Print , "Default for aegitType.XMLDataFolder = " & aegitType.XMLDataFolder
        Debug.Print , "Default for aegitType.XMLDataFolderBe = " & aegitType.XMLDataFolderBe
        Debug.Print , "--------------------------------------------------"
        Debug.Print , "Default for aegitSourceFolder = " & aegitSourceFolder
        Debug.Print , "Default for aegitSourceFolderBe = " & aegitSourceFolderBe
        Debug.Print , "--------------------------------------------------"
        Debug.Print , "Default for aestrSourceLocation = " & aestrSourceLocation
        Debug.Print , "Default for aestrSourceLocationBe = " & aestrSourceLocationBe
        Debug.Print , "Default for aestrXMLLocation = " & aestrXMLLocation
        Debug.Print , "Default for aestrXMLLocationBe = " & aestrXMLLocationBe
        Debug.Print , "Default for aestrXMLDataLocation = " & aestrXMLDataLocation
        Debug.Print , "Default for aestrXMLDataLocationBe = " & aestrXMLDataLocationBe
        Debug.Print , "Default for aestrBackEndDbOne = " & aestrBackEndDbOne
        Debug.Print , "aeintLTN = " & aeintLTN
        Debug.Print , "aeintFNLen = " & aeintFNLen
        Debug.Print , "aeintFTLen = " & aeintFTLen
        Debug.Print , "aeintFSize = " & aeintFSize
        '
        Debug.Print , "aegitExport.ExportAll = " & aegitExport.ExportAll
        Debug.Print , "aegitExport.ExportCodeAndObjects = " & aegitExport.ExportCodeAndObjects
        Debug.Print , "aegitExport.ExportCodeOnly = " & aegitExport.ExportModuleCodeOnly
        Debug.Print , "aegitExport.EXPERIMENTAL_ExportQAT = " & aegitExport.EXPERIMENTAL_ExportQAT
        Debug.Print , "aegitExport.EXPERIMENTAL_ExportCBID = " & aegitExport.EXPERIMENTAL_ExportCBID
        Debug.Print , "aegitExport.ExportNoODBCTablesInfo = " & aegitExport.ExportNoODBCTablesInfo
        DefineMyExclusions
        Debug.Print , "pExclude = " & pExclude
    End If

    If aeExists("Forms", "_frmPersist") Then
        If Not IsLoaded("_frmPersist") Then
            DoCmd.OpenForm "_frmPersist", acNormal, , , acFormReadOnly, acHidden
        End If
        Debug.Print , "IsLoaded _frmPersist = " & IsLoaded("_frmPersist")
    End If
    'Stop

PROC_EXIT:
    aeEndLogging "Class_Initialize"
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class aegit_expClass", vbCritical, "ERROR"
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

End Sub

Public Property Get BackEndDbOne() As String
    On Error GoTo 0
    Debug.Print "Property Get BackEndDbOne"
    BackEndDbOne = aestrBackEndDbOne
End Property

Public Property Let BackEndDbOne(ByVal strBackEndDbFullPath As String)
    On Error GoTo 0
    Debug.Print "Property Let BackEndDbOne"
    aestrBackEndDbOne = strBackEndDbFullPath
    Debug.Print , "aestrBackEndDbOne = " & aestrBackEndDbOne
End Property

Public Property Get CompactAndRepair(Optional ByVal varTrueFalse As Variant) As Boolean
    ' Automation for Compact and Repair

    On Error GoTo 0
    Debug.Print "Property Get CompactAndRepair"
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

Public Property Get DocumentRelations(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        'Debug.Print "Property Get DocumentRelations"
        'Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentRelations"
        'Debug.Print , "DEBUGGING IS OFF"
        DocumentRelations = aeDocumentRelations()
    Else
        Debug.Print "Property Get DocumentRelations"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentRelations"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentRelations = aeDocumentRelations(varDebug)
    End If
End Property

Public Property Get DocumentTables(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        'Debug.Print "Property Get DocumentTables"
        'Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTables"
        'Debug.Print , "DEBUGGING IS OFF"
        DocumentTables = aeDocumentTables()
    Else
        Debug.Print "Property Get DocumentTables"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTables"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTables = aeDocumentTables(varDebug)
    End If
End Property

Public Property Get DocumentTablesXML(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        'Debug.Print "Property Get DocumentTablesXML"
        'Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTablesXML"
        'Debug.Print , "DEBUGGING IS OFF"
        DocumentTablesXML = aeDocumentTablesXML()
    Else
        Debug.Print "Property Get DocumentTablesXML"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTablesXML"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTablesXML = aeDocumentTablesXML(varDebug)
    End If
End Property

Public Property Get DocumentTheDatabase(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        'Debug.Print "Property Get DocumentTheDatabase"
        'Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTheDatabase"
        'Debug.Print , "DEBUGGING IS OFF"
        DocumentTheDatabase = aeDocumentTheDatabase()
    Else
        Debug.Print "Property Get DocumentTheDatabase"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTheDatabase = aeDocumentTheDatabase(varDebug)
    End If
End Property

Public Property Get ExcludeFiles(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    ExcludeFiles = pExclude
    Debug.Print , "ExcludeFiles = " & pExclude
    If IsMissing(varDebug) Then
        'Debug.Print "Property Get ExcludeFiles"
        'Debug.Print , "varDebug IS missing so no parameter is passed to Get ExcludeFiles"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print "Property Get ExcludeFiles"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to Get ExcludeFiles"
        Debug.Print , "DEBUGGING TURNED ON"
    End If
End Property

Public Property Let ExcludeFiles(Optional ByVal varDebug As Variant, ByVal blnExclude As Boolean)
    On Error GoTo 0
    Debug.Print "Property Let ExcludeFiles"
    pExclude = blnExclude
    Debug.Print , "Let ExcludeFiles = " & pExclude
    If IsMissing(varDebug) Then
        'Debug.Print "Property Let ExcludeFiles"
        'Debug.Print , "varDebug IS missing so no parameter is passed to Let ExcludeFiles"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print "Property Let ExcludeFiles"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to Let ExcludeFiles"
        Debug.Print , "DEBUGGING TURNED ON"
    End If
End Property

Public Property Get Exists(ByVal strAccObjType As String, _
    ByVal strAccObjName As String, _
    Optional ByVal varDebug As Variant) As Boolean

    On Error GoTo 0
    If IsMissing(varDebug) Then
        'Debug.Print "Property Get Exists"
        'Debug.Print , "varDebug IS missing so no parameter is passed to aeExists"
        'Debug.Print , "DEBUGGING IS OFF"
        Exists = aeExists(strAccObjType, strAccObjName)
    Else
        Debug.Print "Property Get Exists"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeExists"
        Debug.Print , "DEBUGGING TURNED ON"
        Exists = aeExists(strAccObjType, strAccObjName, varDebug)
    End If
End Property

Public Property Let EXPERIMENTAL_ExportCBID(ByVal IsEXPERIMENTAL_ExportCBID As Boolean)
    On Error GoTo 0
    Debug.Print "Property Let EXPERIMENTAL_ExportCBID"
    If IsEXPERIMENTAL_ExportCBID Then
        aegitExport.EXPERIMENTAL_ExportCBID = True
    Else
        aegitExport.EXPERIMENTAL_ExportCBID = False
    End If
End Property

Public Property Let ExportNoODBCTablesInfo(ByVal ExportNoODBCTablesInfo As Boolean)
    On Error GoTo 0
    Debug.Print "Property Let ExportNoODBCTablesInfo"
    If ExportNoODBCTablesInfo Then
        aegitExport.ExportNoODBCTablesInfo = True
    Else
        aegitExport.ExportNoODBCTablesInfo = False
    End If
End Property

Public Property Let EXPERIMENTAL_ExportQAT(ByVal IsEXPERIMENTAL_ExportQAT As Boolean)
    On Error GoTo 0
    Debug.Print "Property Let EXPERIMENTAL_ExportQAT"
    If IsEXPERIMENTAL_ExportQAT Then
        aegitExport.EXPERIMENTAL_ExportQAT = True
    Else
        aegitExport.EXPERIMENTAL_ExportQAT = False
    End If
End Property

Public Property Get FrontEndApp() As Boolean
    On Error GoTo 0
    Debug.Print "Property Get FrontEndApp"
    FrontEndApp = aegitFrontEndApp
    'Debug.Print , "FrontEndApp = " & FrontEndApp
End Property

Public Property Let FrontEndApp(ByVal IsFrontEndApp As Boolean)
    On Error GoTo 0
    'Debug.Print "Property Let FrontEndApp"
    aegitFrontEndApp = IsFrontEndApp
    'Debug.Print , "aegitFrontEndApp = " & aegitFrontEndApp
End Property

Public Property Get GetReferences(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        'Debug.Print "Property Get GetReferences"
        'Debug.Print , "varDebug IS missing so no parameter is passed to aeGetReferences"
        'Debug.Print , "DEBUGGING IS OFF"
        GetReferences = aeGetReferences()
    Else
        Debug.Print "Property Get GetReferences"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeGetReferences"
        Debug.Print , "DEBUGGING TURNED ON"
        GetReferences = aeGetReferences(varDebug)
    End If
End Property

Public Property Get IsPrimaryKey(ByVal strTableName As String, ByVal strField As String) As Boolean
    On Error GoTo 0
    Debug.Print "Property Get IsPrimaryKey"
    Dim dbs As DAO.Database
    Set dbs = CurrentDb()
    Dim tdf As DAO.TableDef
    Set tdf = dbs.TableDefs(strTableName)
    mblnResult = IsPK(tdf, strField)
    IsPrimaryKey = mblnResult
    Set tdf = Nothing
End Property

Public Property Get SchemaFile(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    Debug.Print "Property Get SchemaFile"
    If IsMissing(varDebug) Then
        OutputTheSchemaFile
    Else
        Debug.Print "Get SchemaFile"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to SchemaFile"
        Debug.Print , "DEBUGGING TURNED ON"
        OutputTheSchemaFile "varDebug"
    End If
    SchemaFile = True
End Property

Public Property Let EXPERIMENTAL_ExportCreateDbScript(ByVal IsEXPERIMENTAL_ExportCreateDbScript As Boolean)
    On Error GoTo 0
    Debug.Print "Property Let EXPERIMENTAL_ExportCreateDbScript"
    If IsEXPERIMENTAL_ExportCreateDbScript Then
        aegitExport.EXPERIMENTAL_ExportCreateDbScript = True
    Else
        aegitExport.EXPERIMENTAL_ExportCreateDbScript = False
    End If
End Property

Public Property Get SourceFolder() As String
    On Error GoTo 0
    Debug.Print "Property Get SourceFolder"
    SourceFolder = aegitSourceFolder
End Property

Public Property Let SourceFolder(ByVal strSourceFolder As String)
    ' Ref: http://www.techrepublic.com/article/build-your-skills-using-class-modules-in-an-access-database-solution/5031814
    ' Ref: http://www.utteraccess.com/wiki/index.php/Classes
    On Error GoTo 0
    Debug.Print "Property Let SourceFolder"
    aegitSourceFolder = strSourceFolder
End Property

Public Property Get SourceFolderBe() As String
    On Error GoTo 0
    Debug.Print "Property Get SourceFolderBe"
    SourceFolderBe = aegitSourceFolderBe      'aestrSourceLocationBe
    'Debug.Print , "SourceFolderBe = " & SourceFolderBe
End Property

Public Property Let SourceFolderBe(ByVal strSourceFolderBe As String)
    On Error GoTo 0
    Debug.Print "Property Let SourceFolderBe"
    aegitSourceFolderBe = strSourceFolderBe
    'aestrSourceLocationBe = strSourceFolderBe
    'Debug.Print , "aestrSourceLocationBe = " & aestrSourceLocationBe
End Property

Public Property Let TablesExportToXML(ByVal varTablesArray As Variant)
    ' Ref: http://stackoverflow.com/questions/2265349/how-can-i-use-an-optional-array-argument-in-a-vba-procedure
    On Error GoTo PROC_ERR
    Debug.Print "Property Let TablesExportToXML"
    'Debug.Print , "LBound(varTablesArray) = " & LBound(varTablesArray), "varTablesArray(0) = " & varTablesArray(0)
    'Debug.Print , "UBound(varTablesArray) = " & UBound(varTablesArray)
    'If UBound(varTablesArray) > 0 Then
    '    Debug.Print , "varTablesArray(1) = " & varTablesArray(1)
    'End If
    ReDim Preserve aegitDataXML(0 To UBound(varTablesArray))
    aegitDataXML = varTablesArray
    'Debug.Print , "aegitDataXML(0) = " & aegitDataXML(0)
    If UBound(varTablesArray) > 0 Then
        Debug.Print , "aegitDataXML(1) = " & aegitDataXML(1)
    End If

PROC_EXIT:
    Exit Property

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TablesExportToXML", vbCritical, "ERROR"
            Stop
    End Select

End Property

Public Property Get TextEncoding() As String
    On Error GoTo 0
    Debug.Print "Property Get TextEncoding"
    TextEncoding = aegitTextEncoding
End Property

Public Property Let TextEncoding(ByVal strTextEncoding As String)
    On Error GoTo 0
    Debug.Print "Property Let TextEncoding"
    aegitTextEncoding = strTextEncoding
End Property

Public Property Get XMLDataFolder() As String
    On Error GoTo 0
    Debug.Print "Property Get XMLDataFolder"
    XMLDataFolder = aegitXMLDataFolder
    'Debug.Print , "XMLDataFolder = " & XMLDataFolder
End Property

Public Property Let XMLDataFolder(ByVal strXMLDataFolder As String)
    On Error GoTo 0
    Debug.Print "Property Let XMLDataFolder"
    aegitXMLDataFolder = strXMLDataFolder
    'Debug.Print , "aegitXMLDataFolder = " & aegitXMLDataFolder
End Property

Public Property Get XMLDataFolderBe() As String
    On Error GoTo 0
    Debug.Print "Property Get XMLDataFolderBe"
    XMLDataFolderBe = aegitXMLDataFolderBe
    'Debug.Print , "XMLDataFolderBe = " & XMLDataFolderBe
End Property

Public Property Let XMLDataFolderBe(ByVal strXMLDataFolderBe As String)
    On Error GoTo 0
    Debug.Print "Property Let XMLDataFolderBe"
    aegitXMLDataFolderBe = strXMLDataFolderBe
    'Debug.Print , "aegitXMLDataFolderBe = " & aegitXMLDataFolderBe
End Property

Public Property Get XMLFolder() As String
    On Error GoTo 0
    Debug.Print "Property Get XMLFolder"
    XMLFolder = aegitXMLFolder
    'Debug.Print , "XMLFolder = " & XMLFolder
End Property

Public Property Let XMLFolder(ByVal strXMLFolder As String)
    On Error GoTo 0
    Debug.Print "Property Let XMLFolder"
    aegitXMLFolder = strXMLFolder
    'Debug.Print , "aegitXMLFolder = " & aegitXMLFolder
End Property

Public Property Get XMLFolderBe() As String
    On Error GoTo 0
    Debug.Print "Property Get XMLFolderBe"
    XMLFolderBe = aegitXMLFolderBe
    'Debug.Print , "XMLFolderBe = " & XMLFolderBe
End Property

Public Property Let XMLFolderBe(ByVal strXMLFolderBe As String)
    On Error GoTo 0
    Debug.Print "Property Let XMLFolderBe"
    aegitXMLFolderBe = strXMLFolderBe
    'Debug.Print , "aegitXMLFolderBe = " & aegitXMLFolderBe
End Property

Private Sub aeBeginLogging(ByVal strProcName As String, Optional ByVal varOne As Variant = vbNullString, _
    Optional ByVal varTwo As Variant = vbNullString, Optional ByVal varThree As Variant = vbNullString)

    On Error GoTo 0
    mlngStartTime = timeGetTime()
    'Debug.Print ">aeBeginLogging"; Space$(1); "mlngStartTime=" & mlngStartTime
    If aeLog.blnNoTrace Then
        'Debug.Print "B1: aeBeginLogging", "blnNoTrace=" & aeLog.blnNoTrace
        Exit Sub
    End If
    If Not aeLog.blnNoTimer Then
        'Debug.Print "B2: aeBeginLogging", "blnNoTimer=" & aeLog.blnNoTimer
        'Debug.Print Format$(mlngStartTime, "0.00"); Space$(2);
    End If
    'Debug.Print Space$(lngIndent * 4); strProcName; Space$(1); "'" & varOne & "'"; Space$(1); "'" & varTwo & "'"; Space$(1); "'" & varThree & "'"
    lngIndent = lngIndent + 1
End Sub

Private Function aeDocumentRelations(Optional ByVal varDebug As Variant) As Boolean
    ' Ref: http://www.tek-tips.com/faqs.cfm?fid=6905
  
    Dim strDocument As String
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim prop As DAO.Property
    Dim strFile As String

    On Error GoTo PROC_ERR

    'Debug.Print "aeDocumentRelations"
    If IsMissing(varDebug) Then
        'Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentRelations"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentRelations"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    strFile = aestrSourceLocation & aeRelTxtFile
    If aegitFrontEndApp Then
        strFile = aestrSourceLocation & aeRelTxtFile
    Else
        strFile = aestrSourceLocationBe & aeRelTxtFile
    End If

    'Debug.Print "strFile=" & strFile
    If Not FileLocked(strFile) Then KillProperly (strFile)
    Open strFile For Append As #1

    For Each rel In CurrentDb.Relations
        'Debug.Print rel.Name
        If Not (Left$(rel.Name, 4) = "MSys" _
            Or Left$(rel.Name, 4) = "~TMP" _
            Or Left$(rel.Name, 3) = "zzz") Then
            strDocument = strDocument & "Name: " & rel.Name & vbCrLf
            strDocument = strDocument & "  " & "Table: " & rel.Table & vbCrLf
            strDocument = strDocument & "  " & "Foreign Table: " & rel.ForeignTable & vbCrLf
            For Each fld In rel.Fields
                strDocument = strDocument & "  PK: " & fld.Name & "   FK:" & fld.ForeignName
                strDocument = strDocument & vbCrLf
            Next fld
        End If

        If Not IsMissing(varDebug) Then
            Debug.Print strDocument
        Else
        End If
        Print #1, strDocument
        strDocument = vbNullString
    Next rel
    
    aeDocumentRelations = True

PROC_EXIT:
    Set prop = Nothing
    Set idx = Nothing
    Set fld = Nothing
    Set rel = Nothing
    Close 1
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentRelations of Class aegit_expClass", vbCritical, "ERROR"
    If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentRelations of Class aegit_expClass"
    aeDocumentRelations = False
    Resume PROC_EXIT

End Function

Private Function aeDocumentTables(Optional ByVal varDebug As Variant) As Boolean
    ' Ref: http://www.tek-tips.com/faqs.cfm?fid=6905
    ' Ref: http://allenbrowne.com/func-06.html
    ' Document the tables, fields, and relationships
    ' Tables, field type, primary keys, foreign keys, indexes
    ' Relationships in the database with table, foreign table, primary keys, foreign keys

    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim blnResult As Boolean
    Dim intFailCount As Integer
    Dim strFile As String

    On Error GoTo PROC_ERR
    If mblnIgnore Then Exit Function

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
        'Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTables"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTables"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    If aegitFrontEndApp Then
        strFile = aestrSourceLocation & aeTblTxtFile
    Else
        strFile = aestrSourceLocationBe & aeTblTxtFile
    End If

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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTables of Class aegit_expClass", vbCritical, "ERROR"
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

    Dim strTheXMLLocation As String
    If aegitXMLFolder = "default" Then
        strTheXMLLocation = aegitType.XMLFolder
    ElseIf aegitFrontEndApp Then
        strTheXMLLocation = aestrXMLLocation
    ElseIf Not aegitFrontEndApp Then
        strTheXMLLocation = aestrXMLLocationBe
    End If

    If Not FolderExists(strTheXMLLocation) Then
        MsgBox strTheXMLLocation & " does not exist!", vbCritical, "aeDocumentTablesXML"
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
        'Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTablesXML"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTablesXML"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    If Not IsMissing(varDebug) Then Debug.Print ">List of tables exported as XML to " & strTheXMLLocation
    If Not aegitExport.ExportNoODBCTablesInfo Then
        For Each tbl In dbs.TableDefs
            If Not IsLinkedTable(tbl.Name) And Not (tbl.Name Like "MSys*") And Not Left$(tbl.Name, 4) = "~TMP" Then
                strObjName = tbl.Name
                Application.ExportXML acExportTable, strObjName, , _
                    strTheXMLLocation & "tables_" & strObjName & ".xsd"
                If Not IsMissing(varDebug) Then
                    Debug.Print , "- " & strObjName & ".xsd"
                    PrettyXML strTheXMLLocation & "tables_" & strObjName & ".xsd", varDebug
                Else
                    PrettyXML strTheXMLLocation & "tables_" & strObjName & ".xsd"
                End If
            End If
        Next
    Else
        For Each tbl In dbs.TableDefs
            If IsLinkedODBC(tbl.Name) Then
                ' Do nothing
            Else
                If Not IsLinkedTable(tbl.Name) And Not (tbl.Name Like "MSys*") And Not Left$(tbl.Name, 4) = "~TMP" Then
                    strObjName = tbl.Name
                    Application.ExportXML acExportTable, strObjName, , _
                        strTheXMLLocation & "tables_" & strObjName & ".xsd"
                    If Not IsMissing(varDebug) Then
                        Debug.Print , "- " & strObjName & ".xsd"
                        PrettyXML strTheXMLLocation & "tables_" & strObjName & ".xsd", varDebug
                    Else
                        PrettyXML strTheXMLLocation & "tables_" & strObjName & ".xsd"
                    End If
                End If
            End If
        Next
    End If

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
    Select Case Err.Number
        Case 31532
            Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTablesXML of Class aegit_expClass"
            Debug.Print , "strObjName=" & strObjName, "strTheXMLLocation=" & strTheXMLLocation          ' from line 430
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTablesXML of Class aegit_expClass", vbCritical, "ERROR"
            If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTablesXML of Class aegit_expClass"
            aeDocumentTablesXML = False
    End Select
    Resume PROC_EXIT

End Function

Private Sub Setup_The_Source_Location()
    If aegitSourceFolder = "default" Then
        mstrTheSourceLocation = aegitType.SourceFolder
    ElseIf aegitFrontEndApp Then
        mstrTheSourceLocation = aestrSourceLocation
    ElseIf Not aegitFrontEndApp Then
        mstrTheSourceLocation = aestrSourceLocationBe
    End If
End Sub

Private Function DocumentTheQueries(Optional ByVal varDebug As Variant) As Boolean

    On Error GoTo PROC_ERR

    Dim qdf As DAO.QueryDef
    Dim i As Integer
    Dim strqdfName As String

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

    ' This will output each query specification to a file and convert UTF-16 to regular text
    For Each qdf In CurrentDb.QueryDefs
        strqdfName = qdf.Name
        If Not IsMissing(varDebug) Then Debug.Print , strqdfName
        If Not (Left$(strqdfName, 4) = "MSys" Or Left$(strqdfName, 4) = "~sq_" _
            Or Left$(strqdfName, 4) = "~TMP" _
            Or Left$(strqdfName, 3) = "zzz") Then
            i = i + 1
            Application.SaveAsText acQuery, strqdfName, mstrTheSourceLocation & strqdfName & ".qry"
            ' Convert UTF-16 to txt - fix for Access 2013+
            If aeReadWriteStream(mstrTheSourceLocation & strqdfName & ".qry") = True Then
                KillProperly (mstrTheSourceLocation & strqdfName & ".qry")
                Name mstrTheSourceLocation & strqdfName & ".qry" & ".clean.txt" As mstrTheSourceLocation & strqdfName & ".qry"
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
    
    DocumentTheQueries = True

PROC_EXIT:
    Set qdf = Nothing
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DocumentTheQueries of Class aegit_expClass", vbCritical, "ERROR"
    'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DocumentTheQueries of Class aegit_expClass"
    DocumentTheQueries = False
    Resume PROC_EXIT

End Function

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

    Dim cnt As DAO.Container
    Dim doc As DAO.Document

    On Error GoTo PROC_ERR

    If IsMissing(varDebug) Then
        'Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTheDatabase"
        'Debug.Print , "DEBUGGING IS OFF"
        VerifySetup
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
        VerifySetup ''' "varDebug"
    End If

    Setup_The_Source_Location

    ListOrCloseAllOpenQueries

    Debug.Print "aeDocumentTheDatabase", "aegitSetup = " & aegitSetup, "aegitFrontEndApp = " & aegitFrontEndApp
    'Stop

    aeBeginLogging "aeDocumentTheDatabase"
    If aegitSetup Then
        'MsgBox "aestrSourceLocation = " & aestrSourceLocation
        KillAllFiles aestrSourceLocation            '= aegitType.SourceFolder
        'MsgBox "aestrXMLLocation = " & aestrXMLLocation
        KillAllFiles aestrXMLLocation               '= aegitType.XMLFolder
        'MsgBox "aestrXMLDataLocation = " & aestrXMLDataLocation
        KillAllFiles aestrXMLDataLocation           '= aegitType.XMLDataFolder
    ElseIf aegitFrontEndApp Then
        'MsgBox "aestrSourceLocation = " & aestrSourceLocation
        KillAllFiles aestrSourceLocation
        'MsgBox "aestrXMLLocation = " & aestrXMLLocation
        KillAllFiles aestrXMLLocation
        'MsgBox "aestrXMLDataLocation = " & aestrXMLDataLocation
        KillAllFiles aestrXMLDataLocation
    ElseIf Not aegitFrontEndApp Then
        'MsgBox "aestrSourceLocationBe = " & aestrSourceLocationBe
        KillAllFiles aestrSourceLocationBe
        'MsgBox "aestrXMLLocationBe = " & aestrXMLLocationBe
        KillAllFiles aestrXMLLocationBe
        'MsgBox "aestrXMLDataLocationBe = " & aestrXMLDataLocationBe
        KillAllFiles aestrXMLDataLocationBe
    End If
    aePrintLog "Timing for KillAllFiles"
    aeEndLogging "aeDocumentTheDatabase"

    ' ===========================================
    '    FORMS REPORTS SCRIPTS MODULES QUERIES
    ' ===========================================
    ' NOTE: Erl(0) Error 2950 if the ouput location does not exist so test for it first: Resolved in VerifySetup

    If Not IsMissing(varDebug) Then
        DocumentTheContainer "Forms", "frm", varDebug
        DocumentTheContainer "Reports", "rpt", varDebug
        DocumentTheContainer "Scripts", "mac", varDebug
        DocumentTheContainer "Modules", "bas", varDebug
        DocumentTheQueries varDebug
    Else
        aeBeginLogging "aeDocumentTheDatabase"
        DocumentTheContainer "Forms", "frm"
        aePrintLog "Timing for DocumentTheContainer Forms"
        aeEndLogging "aeDocumentTheDatabase"

        aeBeginLogging "aeDocumentTheDatabase"
        DocumentTheContainer "Reports", "rpt"
        aePrintLog "Timing for DocumentTheContainer Reports"
        aeEndLogging "aeDocumentTheDatabase"

        aeBeginLogging "aeDocumentTheDatabase"
        DocumentTheContainer "Scripts", "mac"
        aePrintLog "Timing for DocumentTheContainer Macros"
        aeEndLogging "aeDocumentTheDatabase"

        aeBeginLogging "aeDocumentTheDatabase"
        DocumentTheContainer "Modules", "bas"
        aePrintLog "Timing for DocumentTheContainer Modules"
        aeEndLogging "aeDocumentTheDatabase"
    
        aeBeginLogging "aeDocumentTheDatabase"
        DocumentTheQueries
        aePrintLog "Timing for DocumentTheQueries"
        aeEndLogging "aeDocumentTheDatabase"
    End If

    ' =============
    '    OUTPUTS
    ' =============
    aeBeginLogging "aeDocumentTheDatabase Outputs"
    If Not IsMissing(varDebug) Then
        OutputListOfContainers aeAppListCnt, varDebug
        OutputListOfAccessApplicationOptions varDebug
        If aegitExport.EXPERIMENTAL_ExportCBID Then
            OutputListOfCommandBarIDs mstrTheSourceLocation & aeAppCmbrIds, varDebug
            SortTheFile mstrTheSourceLocation & aeAppCmbrIds, mstrTheSourceLocation & aeAppCmbrIds & ".sort"
            KillProperly (mstrTheSourceLocation & aeAppCmbrIds)
        End If
        If aegitExport.ExportNoODBCTablesInfo Then
            OutputListOfTables aegitExport.ExportNoODBCTablesInfo, varDebug
        Else
            OutputListOfTables Not (aegitExport.ExportNoODBCTablesInfo), varDebug
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
        OutputListOfForms varDebug
        OutputListOfMacros varDebug
        OutputListOfModules varDebug
        OutputListOfReports varDebug
        OutputBuiltInPropertiesText varDebug
        OutputAllContainerProperties varDebug
        OutputTableProperties varDebug
        aeGetReferences varDebug
        aeDocumentTables varDebug
        aeDocumentRelations varDebug
        aeDocumentTablesXML varDebug
    Else
        OutputListOfContainers aeAppListCnt
        OutputListOfAccessApplicationOptions
        If aegitExport.EXPERIMENTAL_ExportCBID Then
            OutputListOfCommandBarIDs mstrTheSourceLocation & aeAppCmbrIds
            'Debug.Print , "mstrTheSourceLocation = " & mstrTheSourceLocation
            'Debug.Print , "aeAppCmbrIds = " & aeAppCmbrIds
            'Stop
            SortTheFile mstrTheSourceLocation & aeAppCmbrIds, mstrTheSourceLocation & aeAppCmbrIds & ".sort"
            KillProperly (mstrTheSourceLocation & aeAppCmbrIds)
        End If
        If aegitExport.ExportNoODBCTablesInfo Then
            OutputListOfTables aegitExport.ExportNoODBCTablesInfo, varDebug
        Else
            OutputListOfTables Not (aegitExport.ExportNoODBCTablesInfo), varDebug
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
        OutputListOfForms
        OutputListOfMacros
        OutputListOfModules
        OutputListOfReports
        OutputBuiltInPropertiesText
        OutputAllContainerProperties
        OutputTableProperties
        aeGetReferences
        aeDocumentTables
        aeDocumentRelations
        aeDocumentTablesXML
    End If

    OutputListOfApplicationProperties
    OutputQueriesSqlText
    OutputFieldLookupControlTypeList
    If aegitExport.EXPERIMENTAL_ExportCreateDbScript Then
        OutputTheSchemaFile
        OutputTheSqlFile mstrTheSourceLocation & aeSchemaFile, mstrTheSourceLocation & aeSchemaFile & ".sql"
        OutputTheSqlOnlyFile mstrTheSourceLocation & aeSchemaFile & ".sql", mstrTheSourceLocation & aeSchemaFile & ".sql" & ".only"
        KillProperly (mstrTheSourceLocation & aeSchemaFile & ".sql")
        GenerateLovefieldSchema mstrTheSourceLocation & aeSchemaFile & ".sql" & ".only", mstrTheSourceLocation & aeLoveSchema
    End If
    OutputListOfIndexes mstrTheSourceLocation & aeIndexLists
    aePrintLog "Timing for aeDocumentTheDatabase Outputs"
    aeEndLogging "aeDocumentTheDatabase Outputs"
    'Stop

    If aegitExport.EXPERIMENTAL_ExportQAT Then
        If Not IsMissing(varDebug) Then
            OutputTheQAT aeAppListQAT, varDebug
        Else
            OutputTheQAT aeAppListQAT
        End If
    End If

    aeDocumentTheDatabase = True

PROC_EXIT:
    Set doc = Nothing
    Set cnt = Nothing
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTheDatabase of Class aegit_expClass", vbCritical, "ERROR"
    'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTheDatabase of Class aegit_expClass"
    aeDocumentTheDatabase = False
    Resume PROC_EXIT

End Function

Private Sub aeEndLogging(ByVal strProcName As String, Optional ByVal varOne As Variant = vbNullString, _
    Optional ByVal varTwo As Variant = vbNullString, Optional ByVal varThree As Variant = vbNullString)

    On Error GoTo 0
    If aeLog.blnNoTrace Then
        'Debug.Print "E1: aeEndLogging", "blnNoTrace=" & aeLog.blnNoTrace
        Exit Sub
    End If
    lngIndent = lngIndent - 1
    mlngEndTime = timeGetTime()
    If Not aeLog.blnNoEnd Then
        If Not aeLog.blnNoTimer Then
            'Debug.Print ">aeEndLogging"; Space$(1); "mlngEndTime=" & mlngEndTime
            mlngEndTime = timeGetTime()
            'Debug.Print "E2: aeEndLogging", "blnNoTimer=" & aeLog.blnNoTimer
            'Debug.Print Format$(mlngEndTime, "0.00"); Space$(2);
        End If
        'Debug.Print Space$(lngIndent * 4); "End " & lngIndent; Space$(1); varOne; Space$(1); varTwo; Space$(1); varThree
        Debug.Print , "It took " & (mlngEndTime - mlngStartTime) / 1000 & " seconds to process " & "'" & strProcName & "' procedure"
    End If
End Sub

Private Function aeExists(ByVal strAccObjType As String, _
    ByVal strAccObjName As String, Optional ByVal varDebug As Variant) As Boolean
    ' Ref: http://vbabuff.blogspot.com/2010/03/does-access-object-exists.html
    ' =======================================================================
    ' Author:     Peter F. Ennis
    ' Date:       February 18, 2011
    ' Comment:    Return True if the object exists
    ' Parameters: strAccObjType: "Tables", "Queries", "Forms",
    '                            "Reports", "Macros", "Modules"
    '             strAccObjName: The name of the object
    ' Updated:    All notes moved to change log
    ' History:    See comment details, basChangeLog, commit messages on github
    ' =======================================================================

    Dim objType As Object
    Dim obj As Variant
    
    'Debug.Print "aeExists", strAccObjType, strAccObjName
    On Error GoTo PROC_ERR

    If IsMissing(varDebug) Then
        'Debug.Print , "varDebug IS missing so no parameter is passed to aeExists"
        'Debug.Print , "DEBUGGING IS OFF"
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
            MsgBox "Wrong option! in procedure aeExists of Class aegit_expClass", vbCritical, "ERROR"
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
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class aegit_expClass", vbCritical, "ERROR"
        If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class aegit_expClass"
        aeExists = False
    End If
    Resume PROC_EXIT

End Function

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
    'Dim TheRef As String
    'Dim RefDesc As String
    Dim blnRefBroken As Boolean
    Dim strFile As String

    Dim vbaProj As Object
    Set vbaProj = Application.VBE.ActiveVBProject

    Debug.Print "aeGetReferences"
    On Error GoTo PROC_ERR

    If IsMissing(varDebug) Then
        'Debug.Print , "varDebug IS missing so no parameter is passed to aeGetReferences"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeGetReferences"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    If aegitFrontEndApp Then
        strFile = aestrSourceLocation & aeRefTxtFile
    Else
        strFile = aestrSourceLocationBe & aeRefTxtFile
    End If
    
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

        ' Output reference details
        If Not IsMissing(varDebug) Then Debug.Print , , vbaProj.References(i).Name, vbaProj.References(i).Desc
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeGetReferences of Class aegit_expClass", vbCritical, "ERROR"
    If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeGetReferences of Class aegit_expClass"
    aeGetReferences = False
    Resume PROC_EXIT

End Function

Private Sub aePrintLog(Optional ByVal varOne As Variant = vbNullString, _
    Optional ByVal varTwo As Variant = vbNullString, _
    Optional ByVal varThree As Variant = vbNullString)

    On Error GoTo 0
    If aeLog.blnNoTrace Or aeLog.blnNoPrint Then
        Exit Sub
    End If
    If Not aeLog.blnNoTimer Then
        'Debug.Print Format$(timeGetTime(), "0.00"); Space$(2);
    End If
    'Debug.Print Space$(lngIndent * 4); "'" & varOne & "'"; Space$(1); "'" & varTwo; "'"; Space$(1); "'" & varThree; "'"
    Debug.Print , "'" & varOne & "'"; Space$(1); "'" & varTwo; "'"; Space$(1); "'" & varThree; "'"
End Sub

Private Function aeReadWriteStream(ByVal strPathFileName As String) As Boolean

    'Debug.Print "aeReadWriteStream"
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
    If Asc(tstring) <> 254 And Asc(tstring) <> 255 And Asc(tstring) <> 0 Then
        Close #fnr
        aeReadWriteStream = False
        Exit Function
    End If

    fnr2 = FreeFile()
    Open fName2 For Binary Lock Read Write As #fnr2

    Do While Not EOF(fnr)
        Get #fnr, , tstring
        If Asc(tstring) = 254 Or Asc(tstring) = 255 Or Asc(tstring) = 0 Then
        Else
            Put #fnr2, , tstring
        End If
    Loop

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
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadWriteStream of Class aegit_expClass", vbCritical, "ERROR"
            Resume Next
    End Select

End Function

Private Sub CreateFormReportTextFile(ByVal strFileIn As String, ByVal strFileOut As String, Optional ByVal varDebug As Variant)
    ' Ref: http://social.msdn.microsoft.com/Forums/office/en-US/714d453c-d97a-4567-bd5f-64651e29c93a/how-to-read-text-a-file-into-a-string-1line-at-a-time-search-it-for-keyword-data?forum=accessdev
    ' Ref: http://bytes.com/topic/access/insights/953655-vba-standard-text-file-i-o-statements
    ' Ref: http://www.java2s.com/Code/VBA-Excel-Access-Word/File-Path/ExamplesoftheVBAOpenStatement.htm
    ' Ref: http://www.techonthenet.com/excel/formulas/instr.php
    ' Ref: http://stackoverflow.com/questions/8680640/vba-how-to-conditionally-skip-a-for-loop-iteration
    ' "Checksum =" , "NameMap = Begin",  "PrtMap = Begin",  "PrtDevMode = Begin"
    ' "PrtDevNames = Begin", "PrtDevModeW = Begin", "PrtDevNamesW = Begin"
    ' "OleData = Begin"

    'Debug.Print "CreateFormReportTextFile"
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
                    'GoTo SearchForEnd
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

Private Function DefineMyExclusions() As myExclusions
    Debug.Print "DefineMyExclusions"
    On Error GoTo 0
    myExclude.excludeOne = EXCLUDE_1
    myExclude.excludeTwo = EXCLUDE_2
    myExclude.excludeThree = EXCLUDE_3
End Function

Private Function Delay(ByVal mSecs As Long) As Boolean
    On Error GoTo 0
    Sleep mSecs ' delay milli seconds
End Function

Private Function DescribeIndexField(tdf As DAO.TableDef, strField As String) As Variant
    ' allenbrowne.com
    ' Purpose:   Indicate if the field is part of a primary key or unique index.
    ' Return:    String containing "P" if primary key, "U" if unique index, "I" if non-unique index.
    '            Lower case letters if secondary field in index. Can have multiple indexes.
    ' Arguments: tdf = the TableDef the field belongs to.
    '            strField = name of the field to search the Indexes for.

    Dim idx As DAO.Index        ' Each index of this table
    Dim fld As DAO.Field        ' Each field of the index
    Dim iCount As Integer
    Dim arrReturn() As Variant  ' Return array
    ReDim arrReturn(1, 0)       ' Ref: http://stackoverflow.com/questions/13183775/excel-vba-how-to-redim-a-2d-array

    For Each idx In tdf.Indexes
        iCount = 0
        For Each fld In idx.Fields
            If fld.Name = strField Then
                If idx.Primary Then
                    arrReturn(iCount, iCount) = arrReturn(iCount, iCount) & IIf(iCount = 0, "P", "p")
                    arrReturn(iCount, iCount + 1) = arrReturn(iCount, iCount + 1) & strField
                ElseIf idx.Unique Then
                    arrReturn(iCount, iCount) = arrReturn(iCount, iCount) & IIf(iCount = 0, "U", "u")
                    arrReturn(iCount, iCount + 1) = arrReturn(iCount, iCount + 1) & strField
                Else
                    arrReturn(iCount, iCount) = arrReturn(iCount, iCount) & IIf(iCount = 0, "I", "i")
                    arrReturn(iCount, iCount + 1) = arrReturn(iCount, iCount + 1) & strField
                End If
            End If
            iCount = iCount + 1
            ReDim Preserve arrReturn(1, iCount)
        Next
    Next
    DescribeIndexField = arrReturn()
End Function

Private Function DocumentTheContainer(ByVal strContainerType As String, _
    ByVal strExt As String, _
    Optional ByVal varDebug As Variant) As Boolean
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
        'Debug.Print , "varDebug IS missing so no parameter is passed to DocumentTheContainer"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to DocumentTheContainer"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    Setup_The_Source_Location

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
            MsgBox "Wrong Case Select in DocumentTheContainer", vbCritical, "DocumentTheContainer"
    End Select

    If Not IsMissing(varDebug) Then Debug.Print UCase$(strContainerType)

    For Each doc In cnt.Documents
        'Debug.Print "A", doc.Name, strContainerType
        If Not IsMissing(varDebug) Then Debug.Print , doc.Name
        If Not (Left$(doc.Name, 3) = "zzz" Or Left$(doc.Name, 4) = "~TMP") Then
            i = i + 1
            strTheCurrentPathAndFile = mstrTheSourceLocation & doc.Name & "." & strExt
            'Debug.Print "DocumentTheContainer", "strTheCurrentPathAndFile = " & strTheCurrentPathAndFile
            'Stop
            If IsFileLocked(strTheCurrentPathAndFile) Then
                MsgBox strTheCurrentPathAndFile & " is locked!", vbCritical, "STOP in DocumentTheContainer"
            End If
            KillProperly (strTheCurrentPathAndFile)
SaveAsText:
            If intAcObjType = 5 Then    ' Modules
                'Debug.Print "5:", doc.Name, Excluded(doc.Name)
                If Excluded(doc.Name) And pExclude Then
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
                If intAcObjType = 2 Or intAcObjType = 3 Then    ' Forms or Reports
                    ' Convert UTF-16 to txt - fix for Access 2013
                    If NoBOM(strTheCurrentPathAndFile) Then
                        ' Conversion done
                    Else
                        ' Fallback to old method
                        If aeReadWriteStream(strTheCurrentPathAndFile) = True Then
                            'If intAcObjType = 2 Or intAcObjType = 3 Then Pause (0.25)
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
            Or Not Excluded(doc.Name) Then
            If strContainerType = "Forms" Or strContainerType = "Reports" Then
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
        Debug.Print , "Err=2220 : Resume SaveAsText - " & doc.Name & " - " & strTheCurrentPathAndFile
        Err.Clear
        Pause (0.25)
        Resume SaveAsText
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DocumentTheContainer of Class aegit_expClass", vbCritical, "ERROR"
    DocumentTheContainer = False
    Resume PROC_EXIT

End Function

Private Function Excluded(ByVal strName As String) As Boolean
    'Debug.Print "Excluded"
    On Error GoTo 0
    Excluded = False
    'Debug.Print "1: Excluded", strName, "myExclude.excludeOne = " & myExclude.excludeOne
    If strName = myExclude.excludeOne Then
        Excluded = True
        Exit Function
    End If
    If strName = myExclude.excludeTwo Then
        Excluded = True
        Exit Function
    End If
    If strName = myExclude.excludeThree Then
        Excluded = True
        Exit Function
    End If
End Function

Private Function FieldLookupControlTypeList(Optional ByVal varDebug As Variant) As Boolean
    ' Ref: http://support.microsoft.com/kb/304274
    ' Ref: http://msdn.microsoft.com/en-us/library/office/bb225848(v=office.12).aspx
    ' 106 - acCheckBox, 109 - acTextBox, 110 - acListBox, 111 - acComboBox

    Debug.Print "FieldLookupControlTypeList"
    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDefs
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    Dim lng As Long
    Dim strCheckBoxTable As String
    Dim strCheckBoxField As String

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

    Setup_The_Source_Location

    fle = FreeFile()
    Open mstrTheSourceLocation & "\" & aeFLkCtrFile For Output As #fle

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
            Print #fle, "[" & tbl.Name & "]"
            For Each fld In tbl.Fields
                intAllFieldsCount = intAllFieldsCount + 1
                lng = fld.Properties("DisplayControl").Value
                Print #fle, , "[" & fld.Name & "]", lng, GetType(lng)
                Select Case lng
                    Case acCheckBox
                        intChk = intChk + 1
                        strCheckBoxTable = tbl.Name
                        strCheckBoxField = fld.Name
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
        'Debug.Print "Table with check box is " & strCheckBoxTable
        'Debug.Print "Field with check box is " & strCheckBoxField
    End If

    Print #fle, "Count of Check box = " & intChk
    Print #fle, "Count of Text box  = " & intTxt
    Print #fle, "Count of List box  = " & intLst
    Print #fle, "Count of Combo box = " & intCbo
    Print #fle, "Count of Else      = " & intElse
    Print #fle, "Count of Display Controls = " & intChk + intTxt + intLst + intCbo
    Print #fle, "Count of All Fields = " & intAllFieldsCount - intElse
    'Print #fle, "Table with check box is " & strCheckBoxTable
    'Print #fle, "Field with check box is " & strCheckBoxField

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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FieldLookupControlTypeList of Class aegit_expClass", vbCritical, "ERROR"
    Resume PROC_EXIT

End Function

Private Function FieldTypeName(ByVal fld As DAO.Field) As String
    ' Ref: http://allenbrowne.com/func-06.html
    ' Purpose: Converts the numeric results of DAO Field.Type to text
    
    'Debug.Print "FieldTypeName"
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
            ' Constants for complex types don't work prior to Access 2007 and later.
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

Private Function FileLocked(ByVal strFileName As String) As Boolean
    ' Ref: http://support.microsoft.com/kb/209189

    Debug.Print "FileLocked"
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
        MsgBox ">>> Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FileLocked of Class aegit_expClass", vbCritical, "ERROR"
        FileLocked = True
        Err.Clear
    End If

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FileLocked of Class aegit_expClass", vbCritical, "ERROR"
    Resume PROC_EXIT

End Function

Private Function FixHeaderXML(ByVal strPathFileName As String) As Boolean

    Debug.Print "FixHeaderXML"
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

    Do While Not EOF(fnr)
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
    Loop
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
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FixHeaderXML of Class aegit_expClass", vbCritical, "ERROR"
            Resume Next
    End Select

End Function

Private Function FolderExists(ByVal strPath As String) As Boolean
    ' Ref: http://allenbrowne.com/func-11.html
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Private Function FoundKeywordInLine(ByVal strLine As String, Optional ByVal varEnd As Variant) As Boolean

    'Debug.Print "FoundKeywordInLine"
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
    If InStr(1, strLine, "ImageData = Begin", vbTextCompare) > 0 Then
        FoundKeywordInLine = True
        Exit Function
    End If

End Function

Private Function FoundSqlKeywordInLine(ByVal strLine As String) ', Optional ByVal varEnd As Variant) As Boolean

    'Debug.Print "FoundSqlKeywordInLine"
    On Error GoTo 0

    FoundSqlKeywordInLine = False
    If InStr(1, strLine, "strSQL=", vbTextCompare) > 0 Then
        FoundSqlKeywordInLine = True
        Exit Function
    End If

End Function

Private Sub GenerateLovefieldSchema(ByVal strFileIn As String, ByVal strFileOut As String)
    On Error GoTo 0
    ReadInputWriteOutputLovefieldSchema strFileIn, strFileOut
End Sub

Private Function GetDescription(ByVal obj As Object) As String
    'Debug.Print "GetDescription"
    On Error Resume Next
    GetDescription = obj.Properties("Description")
End Function

Public Function GetFieldInfo(ByVal strSchemaLine As String) As String

    On Error GoTo 0

    Dim strResult As String
    Dim intPosOne As Integer
    Dim intPosTwo As Integer

    intPosOne = InStr(1, strSchemaLine, "[")
    If InStr(1, strSchemaLine, ",") <> 0 Then
        intPosTwo = InStr(1, strSchemaLine, ",")
    Else
        intPosTwo = InStr(1, strSchemaLine, " )") + 1
    End If
    strResult = Mid$(strSchemaLine, intPosOne, intPosTwo - intPosOne)
    'Debug.Print "GetFieldInfo", "strResult=" & strResult
    GetFieldInfo = strResult
    ' Shorten the parse string by removing the found field
    mstrToParse = Right$(strSchemaLine, Len(strSchemaLine) - intPosTwo)
    If strSchemaLine = " )" Then mstrToParse = vbNullString

End Function

Private Function GetIndex(ByVal strSQL As String) As String
    'Debug.Print "GetIndex"
    'Debug.Print , strSQL
    On Error GoTo PROC_ERR

    Dim intPosLB As Integer
    Dim intPosRB As Integer
    Dim strIndexName As String
    Dim strIndexNameIdx As String
    Dim strIndexField As String
    Dim strParseSQL As String

    intPosLB = InStr(strSQL, "[")
    intPosRB = InStr(strSQL, "]")
    strIndexName = Mid$(strSQL, intPosLB + 1, intPosRB - intPosLB - 1)
    'Debug.Print , "A>>>strIndexName", strIndexName, intPosLB + 1, intPosRB - 1
    strIndexNameIdx = "idx" & UCase$(Left$(strIndexName, 1)) & Right$(strIndexName, Len(strIndexName) - 1)
    'Debug.Print , "B>>>strIndexNameIdx", strIndexNameIdx

    Dim intPosPLB As Integer
    intPosPLB = InStr(strSQL, "([")
    strParseSQL = Right$(strSQL, Len(strSQL) - intPosPLB)
    'Debug.Print , "C>>>strParseSQL", strParseSQL
    intPosLB = InStr(strParseSQL, "[")
    intPosRB = InStr(strParseSQL, "])")
    strIndexField = Mid$(strParseSQL, intPosLB + 1, intPosRB - intPosLB - 1)
    'Debug.Print , "D>>>strIndexField", strIndexField, intPosLB + 1, intPosRB - 1
    GetIndex = Space$(4) & "addIndex('" & strIndexNameIdx & "', ['" & strIndexField & "'], false, lf.Order.ASC)"

PROC_EXIT:
    Exit Function

PROC_ERR:
    Select Case Err
        Case 5
            MsgBox "Erl=" & Erl & " Err=" & Err.Number & " (" & Err.Description & ") in procedure GetIndex of Class aegitClass"
            Stop
            Resume PROC_EXIT
        Case Else
            MsgBox "Erl=" & Erl & " Err=" & Err.Number & " (" & Err.Description & ") in procedure GetIndex of Class aegitClass"
            Resume PROC_EXIT
    End Select

End Function

Private Function GetLinkedTableCurrentPath(ByVal strTableName As String) As String
    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=198057
    ' =========================================================================
    ' Procedure: GetLinkedTableCurrentPath
    ' Purpose:   Return Current Path of a Linked Table in Access and do not show password
    ' Author:    Peter F. Ennis
    ' Updated:   All notes moved to change log
    ' History:   See comment details, basChangeLog, commit messages on github
    ' =========================================================================

    'Debug.Print "GetLinkedTableCurrentPath"
    On Error GoTo PROC_ERR
    If mblnIgnore Then Exit Function

    Dim strConnect As String
    Dim intStrConnectLen As Integer
    'Dim intEqualPos As Integer
    Dim intDatabasePos As Integer
    Dim strMidLink As String

    If Not aegitExport.ExportNoODBCTablesInfo Then
        If Len(CurrentDb.TableDefs(strTableName).Connect) > 0 Then
            ' Linked table exists, but is the link valid?
            ' The next line of code will generate Errors 3011 or 3024 if it isn't
            CurrentDb.TableDefs(strTableName).RefreshLink
            ' If you get to this point, you have a valid, Linked Table
            strConnect = CurrentDb.TableDefs(strTableName).Connect
            intStrConnectLen = Len(strConnect)
            intDatabasePos = InStr(1, strConnect, "Database=") + 8
            strMidLink = Mid$(strConnect, intDatabasePos + 1, Len(strConnect) - intDatabasePos)
            'MsgBox "strTableName = " & strTblName & vbCrLf & _
            '    "strConnect = " & strConnect & vbCrLf & _
            '    "intStrConnectLen = " & intStrConnectLen & vbCrLf & _
            '    "intDatabasePos = " & intDatabasePos & " : " & Left$(strConnect, intDatabasePos) & vbCrLf & _
            '    "strMidLink = " & Mid$(strConnect, intDatabasePos + 1, Len(strConnect) - intDatabasePos) & vbCrLf _
            '    , vbInformation, "GetLinkedTableCurrentPath"
            GetLinkedTableCurrentPath = strMidLink
        Else
            GetLinkedTableCurrentPath = "Local Table=>" & strTableName
        End If
    Else
        ' Check if it is an ODBC link
        If IsLinkedODBC(strTableName) Then
            ' Return an indicator that it is an ODBC table
            GetLinkedTableCurrentPath = "ODBC"
        Else
            If Len(CurrentDb.TableDefs(strTableName).Connect) > 0 Then
                CurrentDb.TableDefs(strTableName).RefreshLink
                ' If you get to this point, you have a valid, Linked Table
                strConnect = CurrentDb.TableDefs(strTableName).Connect
                intStrConnectLen = Len(strConnect)
                intDatabasePos = InStr(1, strConnect, "Database=") + 8
                strMidLink = Mid$(strConnect, intDatabasePos + 1, Len(strConnect) - intDatabasePos)
                GetLinkedTableCurrentPath = strMidLink
            Else
                GetLinkedTableCurrentPath = "Local Table=>" & strTableName
            End If
        End If
    End If

PROC_EXIT:
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 3151, 3059
            'MsgBox "mblnIgnore = " & mblnIgnore
            If mblnIgnore Then Resume PROC_EXIT
        Case 3265
            MsgBox "(" & strTableName & ") does not exist as either an Internal or Linked Table", _
                vbCritical, "Table Missing"
        Case 3011, 3024                 ' Linked Table does not exist or DB Path not valid
            MsgBox "(" & strTableName & ") is not a valid Linked Table", vbCritical, "Link Not Valid"
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetLinkedTableCurrentPath of Class aegit_expClass", vbCritical, "ERROR"
    End Select
    Resume PROC_EXIT
End Function

Public Function GetLovefieldType(ByVal strAccessFieldType As String) As String
    'Debug.Print "GetLovefieldType"
    On Error GoTo 0

    Dim accessFieldType As String

    'Debug.Print , "strAccessFieldType=" & strAccessFieldType
    If Left$(strAccessFieldType, 4) = "Text" Then
        accessFieldType = "Text"
    Else
        accessFieldType = strAccessFieldType
    End If

    Select Case accessFieldType
        Case "Attachment"
            GetLovefieldType = "', lf.Type.OBJECT)."        ' "OBJECT"
        Case "Counter"
            GetLovefieldType = "', lf.Type.INTEGER)."       ' "INTEGER"
        Case "Currency"
            GetLovefieldType = "', lf.Type.STRING)."        ' "STRING"
        Case "DateTime"
            GetLovefieldType = "', lf.Type.DATE_TIME)."     ' "DATE_TIME"
        Case "Double"
            GetLovefieldType = "', lf.Type.NUMBER)."        ' "NUMBER"
        Case "Hyperlink"
            GetLovefieldType = "', lf.Type.STRING)."        ' "STRING"
        Case "Integer"
            GetLovefieldType = "', lf.Type.INTEGER)."       ' "INTEGER"
        Case "Long"
            GetLovefieldType = "', lf.Type.INTEGER)."       ' "INTEGER"
        Case "Memo"
            GetLovefieldType = "', lf.Type.STRING)."        ' "STRING"
        Case "OleObject"
            GetLovefieldType = "', lf.Type.ARRAY_BUFFER)."  ' "OleObject/???"
        Case "Text"
            GetLovefieldType = "', lf.Type.STRING)."        ' "STRING"
        Case "YesNo"
            GetLovefieldType = "', lf.Type.BOOLEAN)."       ' "BOOLEAN"
            '
            ' NOTE: The following come from tables linked with SQL Database Azure
        Case "BYTE"
            GetLovefieldType = "', lf.Type.INTEGER)."
        Case "DECIMAL"
            GetLovefieldType = "', lf.Type.NUMBER)."
        Case "GUID"
            GetLovefieldType = "', lf.Type.STRING)."
        Case Else
            MsgBox "Unknown Access Field Type in procedure GetLovefieldType of Class aegitClass" & vbCrLf & _
                "accessFieldType=" & accessFieldType, vbCritical, "GetLovefieldType"
    End Select

End Function

Private Function GetPrimaryKey(ByVal strSQL As String) As String
    'Debug.Print "GetPrimaryKey"
    'Debug.Print , strSQL
    On Error GoTo 0

    Dim intPosRB As Integer
    Dim intPosPLB As Integer
    Dim intPosRBP As Integer
    Dim strPrimaryKeyName As String
    Dim strParse As String
    Dim strPrimaryField As String

    strParse = Right$(strSQL, Len(strSQL) - Len(mPRIMARYKEY))
    'Debug.Print , ">>strParse", strParse
    intPosRB = InStr(strParse, "]")
    strPrimaryKeyName = Left$(strParse, intPosRB - 1)
    'Debug.Print , ">>strPrimaryKeyName", strPrimaryKeyName
    intPosPLB = InStr(strParse, "([") + 1
    strParse = Right$(strParse, Len(strParse) - intPosPLB)
    'Debug.Print , ">>strParse", strParse
    intPosRBP = InStr(strParse, "])") - 1
    strPrimaryField = Left$(strParse, intPosRBP)
    'Debug.Print , ">>strPrimaryField", intPosRBP, strPrimaryField
    GetPrimaryKey = Space$(4) & "addPrimaryKey(['" & strPrimaryField & "'])"
    'Stop
End Function

Private Function GetPropEnum(ByVal typeNum As Long, Optional ByVal varDebug As Variant) As String
    ' Ref: http://msdn.microsoft.com/en-us/library/bb242635.aspx

    'Debug.Print "GetPropEnum"
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
            Debug.Print , "Unknown typeNum:" & typeNum & " in procedure GetPropEnum of aegit_expClass"
    End Select

PROC_EXIT:
    Exit Function

PROC_ERR:
    Select Case Err.Number
        'Case 3251
        '    strError = " " & Err.Number & ", '" & Err.Description & "'"
        '    varPropValue = Null
        '    Resume Next
        Case Else
            'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetPropEnum of Class aegit_expClass", vbCritical, "ERROR"
            If Not IsMissing(varDebug) Then Debug.Print "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetPropEnum of Class aegit_expClass"
            GetPropEnum = CStr(typeNum)
            Resume PROC_EXIT
    End Select

End Function

Private Function GetPropValue(ByVal obj As Object) As String
    'Debug.Print "GetPropValue"
    'On Error Resume Next
    On Error GoTo 0
    GetPropValue = obj.Properties("Value")
End Function

Public Function GetTableName(ByVal strSchemaLine As String) As String
    'Debug.print "GetTableName"
    On Error GoTo 0

    Dim intPosOne As Integer
    Dim intPosTwo As Integer
    Dim strTableName As String

    intPosOne = InStr(1, strSchemaLine, "[") + 1
    intPosTwo = InStr(1, strSchemaLine, "]")
    strTableName = Mid$(strSchemaLine, intPosOne, intPosTwo - intPosOne)
    'Debug.Print "strTableName=" & strTableName
    GetTableName = strTableName

End Function

Private Function GetType(ByVal Value As Long) As String
    ' Ref: http://bytes.com/topic/access/answers/557780-getting-string-name-enum

    'Debug.Print "GetType"
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
            MsgBox "A:Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegit_expClass", vbCritical, "ERROR"
            IsFileLocked = True
        Case 9
            MsgBox "B:Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegit_expClass" & _
                vbCrLf & "IsFileLocked Entry PathFileName=" & PathFileName, vbCritical, "ERROR=9"
            IsFileLocked = False
            'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegit_expClass"
            Resume PROC_EXIT
        Case Else
            MsgBox "C:Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegit_expClass", vbCritical, "ERROR"
            'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegit_expClass"
            Resume PROC_EXIT
    End Select
    Resume

End Function

Private Function IsFormHidden(ByVal strFormName As String) As Boolean
    'Debug.Print "IsFormHidden"
    On Error GoTo 0
    If IsNull(strFormName) Or strFormName = vbNullString Then
        IsFormHidden = False
    Else
        IsFormHidden = GetHiddenAttribute(acForm, strFormName)
    End If
End Function

Private Function IsLinkedODBC(strTableName As String) As Boolean
    ' Ref: http://www.pcreview.co.uk/threads/check-if-a-table-is-linked.3748757/
    ' Non-linked tables are type 1
    ' ODBC linked tables are type 4
    ' All other linked tables are type 6

    'Debug.Print strTableName, Nz(DLookup("Type", "MSysObjects", "Name = '" & strTableName & "'"))
    IsLinkedODBC = Nz(DLookup("Type", "MSysObjects", "Name = '" & strTableName & "'"), 0) = 4
End Function

Private Function IsLinkedTable(ByVal strTableName As String) As Boolean

    'Debug.Print "LinkedTable"
    On Error GoTo PROC_ERR
    If mblnIgnore Then Exit Function

    Dim intAnswer As Integer

    ' Check if it is an ODBC table and ignore if aegitExport.ExportNoODBCTablesInfo is true
    ' Use results from IsLinkedODBC(strTableName) before trying to connect so the ODBC login
    ' window does not open
    If Not aegitExport.ExportNoODBCTablesInfo Then
        ' Linked table connection string is > 0
        If Len(CurrentDb.TableDefs(strTableName).Connect) > 0 Then
            ' Linked table exists, but is the link valid?
            ' The next line of code will generate Errors 3011 or 3024 if it isn't
            CurrentDb.TableDefs(strTableName).RefreshLink
            'If you get to this point, you have a valid, Linked Table
            IsLinkedTable = True
            'Debug.Print "LinkedTable = True"
        Else
            ' Local table connect string length = 0
            ' MsgBox "[" & strTableName & "] is a Non-Linked Table", vbInformation, "Internal Table"
            IsLinkedTable = False
            'Debug.Print "LinkedTable = False"
        End If
    Else
        ' Check if it is an ODBC link
        If IsLinkedODBC(strTableName) Then
            ' Do nothing and treat the link as false
            IsLinkedTable = False
        Else
            If Len(CurrentDb.TableDefs(strTableName).Connect) > 0 Then
                CurrentDb.TableDefs(strTableName).RefreshLink
                IsLinkedTable = True
            Else
                IsLinkedTable = False
            End If
        End If
    End If

PROC_EXIT:
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 3151, 3059
            'MsgBox "mblnIgnore = " & mblnIgnore
            If mblnIgnore Then Resume PROC_EXIT
            MsgBox "Err=" & Err.Number & " " & Err.Description, vbExclamation, "IsLinkedTable Error"
            intAnswer = MsgBox("Ignore further errors of this type?", vbYesNo + vbQuestion, "IsLinkedTable Error")
            If intAnswer = vbYes Then
                mblnIgnore = True
            Else
                'do nothing
            End If
        Case 3265
            MsgBox "[" & strTableName & "] does not exist as either an Internal or Linked Table", _
                vbCritical, "Table Missing"
        Case 3011, 3024     'Linked Table does not exist or DB Path not valid
            MsgBox "[" & strTableName & "] is not a valid Linked Table", vbCritical, "Link Not Valid"
        Case Else
            MsgBox "Err=" & Err.Number & " " & Err.Description, vbExclamation, "IsLinkedTable Error"
    End Select
    Resume PROC_EXIT
End Function

Private Function IsLoaded(ByVal strFormName As String) As Boolean
    ' Returns True if the specified form is open in Form view or Datasheet view.
   
    On Error GoTo 0
    'Debug.Print "IsLoaded"
    Const conObjStateClosed As Integer = 0
    Const conDesignView As Integer = 0
    
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If
    
End Function

Private Function IsMacroHidden(ByVal strMacroName As String) As Boolean
    'Debug.Print "IsMacroHidden"
    On Error GoTo 0
    If IsNull(strMacroName) Or strMacroName = vbNullString Then
        IsMacroHidden = False
    Else
        IsMacroHidden = GetHiddenAttribute(acMacro, strMacroName)
    End If
End Function

Private Function IsModuleHidden(ByVal strModuleName As String) As Boolean
    'Debug.Print "IsModuleHidden"
    On Error GoTo 0
    If IsNull(strModuleName) Or strModuleName = vbNullString Then
        IsModuleHidden = False
    Else
        IsModuleHidden = GetHiddenAttribute(acModule, strModuleName)
    End If
End Function

Private Function IsPK(ByVal tdf As DAO.TableDef, ByVal strField As String) As Boolean
    'Debug.Print "isPK"
    On Error GoTo 0

    Dim idx As DAO.Index
    Dim fld As DAO.Field
    IsPK = False
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

Private Function IsReportHidden(ByVal strReportName As String) As Boolean
    'Debug.Print "IsReportHidden"
    On Error GoTo 0
    If IsNull(strReportName) Or strReportName = vbNullString Then
        IsReportHidden = False
    Else
        IsReportHidden = GetHiddenAttribute(acReport, strReportName)
    End If
End Function

Private Function IsSingleIndexField(ByVal tdf As DAO.TableDef, ByRef FieldCountResult As Integer) As Boolean

    Dim strIndexInfo As String
    strIndexInfo = SingleTableIndexSummary(tdf)
    Debug.Print strIndexInfo
    FieldCountResult = LCaseCountChar("I", strIndexInfo)
    If FieldCountResult = 1 Then
        IsSingleIndexField = True
        Debug.Print , "Single Field Index", "IsSingleIndexField is " & IsSingleIndexField
    ElseIf FieldCountResult > 1 Then
        IsSingleIndexField = False
        Debug.Print , "Multi Field Index", "IsSingleIndexField is " & IsSingleIndexField
    ElseIf FieldCountResult = 0 Then
        IsSingleIndexField = False
        Debug.Print , "No Index", "IsSingleIndexField is " & IsSingleIndexField
    End If

End Function

Private Function IsSinglePrimaryField(ByVal tdf As DAO.TableDef, ByRef PrimaryIndexFieldCount As Integer) As Boolean

    Dim strIndexInfo As String
    strIndexInfo = SingleTableIndexSummary(tdf)
    PrimaryIndexFieldCount = LCaseCountChar("P", strIndexInfo)
    If PrimaryIndexFieldCount = 1 Then
        IsSinglePrimaryField = True
        'Debug.Print , strIndexInfo, "Single Field Primary Key", IsSinglePrimaryField
    ElseIf PrimaryIndexFieldCount > 1 Then
        IsSinglePrimaryField = False
        Debug.Print , strIndexInfo, "Multi Field Primary Key"
    ElseIf PrimaryIndexFieldCount = 0 Then
        IsSinglePrimaryField = False
        'Debug.Print , strIndexInfo, "No Primary Key"
    End If

End Function

Private Function IsTableHidden(ByVal strTableName As String) As Boolean
    'Debug.Print "IsTableHidden"
    On Error GoTo 0
    If IsNull(strTableName) Or strTableName = vbNullString Then
        IsTableHidden = False
    Else
        IsTableHidden = GetHiddenAttribute(acTable, strTableName)
    End If
End Function

Private Function IsTableSchemaDone(ByVal strTableName As String, ByVal strSQL As String) As Boolean
    'Debug.print "IsTableSchemaDone"
    On Error GoTo 0
    
    IsTableSchemaDone = True
    If InStr(strSQL, strTableName) Then
        IsTableSchemaDone = False
    End If
End Function

Private Sub KillAllFiles(ByVal strLoc As String)

    Dim strFile As String

    Debug.Print "KillAllFiles"
    'Debug.Print , "strLoc = " & strLoc
    On Error GoTo PROC_ERR

    ' Test for relative path - it should already have been converted to an absolute location
    If Left$(strLoc, 1) = "." Then Stop
    'Stop

    ' Delete exported files
    strFile = Dir$(strLoc & "*.*")
    Do While strFile <> vbNullString
        KillProperly (strLoc & strFile)
        ' Need to specify full path again because a file was deleted
        strFile = Dir$(strLoc & "*.*")
    Loop

PROC_EXIT:
    Exit Sub

PROC_ERR:
    If Err = 70 Then    ' Permission denied
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure KillAllFiles of Class aegit_expClass" _
            & vbCrLf & vbCrLf & _
            "Manually delete the exported files, compact and repair the database, then try again!", vbCritical, "STOP"
        Stop
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure KillAllFiles of Class aegit_expClass", vbCritical, "ERROR"
    Resume PROC_EXIT

End Sub

Private Sub KillProperly(ByVal Killfile As String)
    ' Ref: http://word.mvps.org/faqs/macrosvba/DeleteFiles.htm

    'Debug.Print "KillProperly"
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
    ElseIf Err = 53 Then     ' File not found
        Resume PROC_EXIT
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " Killfile=" & Killfile & " (" & Err.Description & ") in procedure KillProperly of Class aegit_expClass", vbCritical, "ERROR"
    Resume PROC_EXIT

End Sub

Private Function LCaseCountChar(ByVal searchChar As String, ByVal searchString As String) As Long
    Dim i As Long
    For i = 1 To Len(searchString)
        If Mid$(LCase$(searchString), i, 1) = LCase(searchChar) Then
            LCaseCountChar = LCaseCountChar + 1
        End If
    Next
End Function

Private Sub ListAllContainerProperties(ByVal strContainer As String, Optional ByVal varDebug As Variant)
    ' Ref: http://www.dbforums.com/microsoft-access/1620765-read-ms-access-table-properties-using-vba.html
    ' Ref: http://msdn.microsoft.com/en-us/library/office/aa139941(v=office.10).aspx
    
    Debug.Print "ListAllContainerProperties"
    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim obj As Object
    Dim prp As DAO.Property
    Dim doc As DAO.Document
    Dim fle As Integer

    Set dbs = Application.CurrentDb
    Set obj = dbs.Containers(strContainer)

    Setup_The_Source_Location

    fle = FreeFile()
    Open mstrTheSourceLocation & "\OutputContainer" & strContainer & "Properties.txt" For Output As #fle

    ' Ref: http://stackoverflow.com/questions/16642362/how-to-get-the-following-code-to-continue-on-error
    For Each doc In obj.Documents
        If Left$(doc.Name, 4) <> "MSys" And Left$(doc.Name, 3) <> "zzz" _
            And Left$(doc.Name, 1) <> "~" Then
            If Not IsMissing(varDebug) Then Debug.Print ">>>" & doc.Name
            Print #fle, ">>>" & doc.Name
            For Each prp In doc.Properties
                On Error Resume Next
                If prp.Name = "GUID" And strContainer = "tables" Then
                    Print #fle, , prp.Name, "GUID"                  ' ListGUID(doc.Name) => just output "GUID" to file
                    If Not IsMissing(varDebug) Then Debug.Print , prp.Name, ListGUID(doc.Name)
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListAllContainerProperties of Class aegit_expClass", vbCritical, "ERROR"
    'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListAllContainerProperties of Class aegit_expClass"
    Resume PROC_EXIT

End Sub

Private Function ListGUID(ByVal strTableName As String) As String
    ' Ref: http://stackoverflow.com/questions/8237914/how-to-get-the-guid-of-a-table-in-microsoft-access
    ' e.g. ?ListGUID("tblThisTableHasSomeReallyLongNameButItCouldBeMuchLonger")

    'Debug.Print "ListGUID"
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
    ListGUID = Left$(strGuid, 23)

End Function

Private Sub ListOrCloseAllOpenQueries(Optional ByVal strCloseAll As Variant)
    ' Ref: http://msdn.microsoft.com/en-us/library/office/aa210652(v=office.11).aspx

    Debug.Print "ListOrCloseAllOpenQueries"
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

Private Function LongestFieldPropsName() As Boolean
    ' =======================================================================
    ' Author:   Peter F. Ennis
    ' Date:     December 5, 2012
    ' Comment:  Return length of field properties for text output alignment
    ' Updated:  All notes moved to change log
    ' History:  See comment details, basChangeLog, commit messages on github
    ' =======================================================================

    Dim tblDef As DAO.TableDef
    Dim fld As DAO.Field

    Debug.Print "LongestFieldPropsName"
    On Error GoTo PROC_ERR

    aeintFNLen = 0
    aeintFTLen = 0
    aeintFDLen = 0


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
                If Len(GetDescription(fld)) > aeintFDLen Then
                    aestrLFD = GetDescription(fld)
                    aeintFDLen = Len(GetDescription(fld))
                End If
            Next
        End If
    Next tblDef

    LongestFieldPropsName = True

PROC_EXIT:
    Set fld = Nothing
    Set tblDef = Nothing
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestFieldPropsName of Class aegit_expClass", vbCritical, "ERROR"
    LongestFieldPropsName = False
    Resume PROC_EXIT

End Function

Private Function LongestTableDescription(ByVal strTblName As String) As Integer
    ' ?LongestTableDescription("tblCaseManager")

    'Debug.Print "LongestTableDescription"
    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim strLFD As String

    On Error GoTo PROC_ERR

    Set dbs = CurrentDb()
    Set tdf = dbs.TableDefs(strTblName)

    For Each fld In tdf.Fields
        If Len(GetDescription(fld)) > aeintFDLen Then
            strLFD = GetDescription(fld)
            aeintFDLen = Len(GetDescription(fld))
        End If
    Next

    LongestTableDescription = aeintFDLen

PROC_EXIT:
    Set fld = Nothing
    Set tdf = Nothing
    Set dbs = Nothing
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestTableDescription of Class aegit_expClass", vbCritical, "ERROR"
    LongestTableDescription = -1
    Resume PROC_EXIT

End Function

Private Function MySortIt(ByVal strFPName As String, ByVal strExtension As String, _
    Optional ByVal varUnicode As Variant) As Long
    ' Ref: http://support.microsoft.com/kb/150700
    ' Ref: http://www.xtremevbtalk.com/showthread.php?t=291063
    ' Ref: http://www.ozgrid.com/forum/showthread.php?t=167349

    Debug.Print "MySortIt"
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
    'Debug.Print "DONE !!!"

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure MySortIt of Class aegit_expClass", vbCritical, "ERROR"
    Resume PROC_EXIT

End Function

Private Function NoBOM(ByVal strFileName As String) As Boolean
    ' Ref: http://www.experts-exchange.com/Programming/Languages/Q_27478996.html
    ' Use the same file name for input and output

    'Debug.Print "NoBOM"
    On Error GoTo PROC_ERR

    ' Define needed constants
    'Const ForReading As Integer = 1
    Const ForWriting As Integer = 2
    'Const TriStateUseDefault As Integer = -2
    Const adTypeText As Integer = 2
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
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure NoBOM of Class aegit_expClass", vbCritical, "ERROR"
            Resume PROC_EXIT
    End Select
    Resume PROC_EXIT

End Function

Private Sub OpenAllDatabases(ByVal blnInit As Boolean)
    ' Open a handle to all databases and keep it open during the entire time the application runs.
    ' Parameter: blnInit - TRUE to initialize (call when application starts), FALSE to close (call when application ends)
    ' Ref: http://stackoverflow.com/questions/29838317/issue-when-using-a-dao-handle-when-the-database-closes-unexpectedly

    Debug.Print "OpenAllDatabases"

    Dim intX As Integer
    Dim strName As String
    Dim strMsg As String

    If aestrBackEndDbOne = "NONE" Then
        Exit Sub
    Else
        Debug.Print , "aestrBackEndDbOne = " & aestrBackEndDbOne, "OpenAllDatabases"
    End If

    ' Maximum number of back end databases to link
    Const cintMaxDatabases As Integer = 1

    ' List of databases kept in a static array so we can close them later
    Static dbsOpen() As DAO.Database
 
    'MsgBox "aestrBackEndDbOne = " & aestrBackEndDbOne, vbInformation, "OpenAllDatabases"
    If blnInit Then
        ReDim dbsOpen(1 To cintMaxDatabases)
        For intX = 1 To cintMaxDatabases
            ' Specify your back end databases
            Select Case intX
                Case 1:
                    strName = aestrBackEndDbOne
                Case 2:
                    strName = "H:\folder\Backend2.mdb"
                Case Else
                    MsgBox "This should never occur!", vbCritical, "OpenAllDatabases"
                    Stop
            End Select
            strMsg = vbNullString
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
            If strMsg <> vbNullString Then
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

Public Sub OutputAllContainerProperties(Optional ByVal varDebug As Variant)

    Debug.Print "OutputAllContainerProperties"
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputAllContainerProperties of Class aegit_expClass", vbCritical, "ERROR"
    'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputAllContainerProperties of Class aegit_expClass"
    Resume PROC_EXIT

End Sub

Private Function OutputBuiltInPropertiesText(Optional ByVal varDebug As Variant) As Boolean
    ' Ref: http://www.jpsoftwaretech.com/listing-built-in-access-database-properties/

    Dim dbs As DAO.Database
    Dim prps As DAO.Properties
    Dim prp As DAO.Property
    Dim varPropValue As Variant
    Dim strFile As String
    Dim strError As String

    Debug.Print "OutputBuiltInPropertiesText"
    On Error GoTo PROC_ERR

    If aegitFrontEndApp Then
        strFile = aestrSourceLocation & aePrpTxtFile
    Else
        strFile = aestrSourceLocationBe & aePrpTxtFile
    End If

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
        varPropValue = GetPropValue(prp)
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
            'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputBuiltInPropertiesText of Class aegit_expClass", vbCritical, "ERROR"
            If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputBuiltInPropertiesText of Class aegit_expClass"
            OutputBuiltInPropertiesText = False
            Resume PROC_EXIT
    End Select

End Function

Public Sub OutputCatalogUserCreatedObjects(Optional ByVal varDebug As Variant)
    ' Ref: http://blogannath.blogspot.com/2010/03/microsoft-access-tips-tricks-working.html#ixzz3WCBJcxwc
    ' Ref: http://stackoverflow.com/questions/5286620/saving-a-query-via-access-vba-code

    Debug.Print "OutputCatalogUserCreatedObjects"
    On Error GoTo PROC_ERR

    Dim strSQL As String
    Dim fle As Integer
    fle = FreeFile()

'    Dim strPathFileName As String
    If aegitFrontEndApp Then
'        strPathFileName = aestrSourceLocation & aeCatalogObj
        Open aestrSourceLocation & aeCatalogObj For Output As #fle
    Else
'        strPathFileName = aestrSourceLocationBe & aeCatalogObj
        Open aestrSourceLocationBe & aeCatalogObj For Output As #fle
    End If

    If Not IsMissing(varDebug) Then
        Debug.Print "OutputCatalogUserCreatedObjects"
'        Debug.Print , strPathFileName
    Else
    End If

    ' Ref: https://support.office.com/en-za/article/FormatDateTime-Function-aef62949-f957-4ba4-94ff-ace14be4f1ca
    ' Format DateCreate as short date, vbShortDate = 2
    'SELECT IIf(type=1,"Table",IIf(type=6,"Linked Table",IIf(type=5,"Query",IIf(type=-32768,"Form",IIf(type=-32764,"Report",IIf(type=-32766,"Module",IIf(type=-32761,"Module","Unknown"))))))) AS [Object Type], MSysObjects.Name, FormatDateTime([DateCreate],2) AS DateCreated
    'FROM MSysObjects
    'WHERE (((MSysObjects.[Type]) In (1,5,6,-32768,-32764,-32766,-32761)) AND ((Left$([Name],4))<>"MSys") AND ((Left$([Name],1))<>"~"))
    'ORDER BY IIf(type=1,"Table",IIf(type=6,"Linked Table",IIf(type=5,"Query",IIf(type=-32768,"Form",IIf(type=-32764,"Report",IIf(type=-32766,"Module",IIf(type=-32761,"Module","Unknown"))))))), MSysObjects.Name;

    strSQL = "SELECT IIf(type = 1,""Table"", IIf(type = 6, ""Linked Table"", "
    strSQL = strSQL & vbCrLf & "IIf(type = 5,""Query"", IIf(type = -32768,""Form"", "
    strSQL = strSQL & vbCrLf & "IIf(type = -32764,""Report"", IIf(type=-32766,""Module"", "
    strSQL = strSQL & vbCrLf & "IIf(type = -32761,""Module"", ""Unknown""))))))) as [Object Type], "
    strSQL = strSQL & vbCrLf & "MSysObjects.Name, ""DateCreated"" AS DateCreated "
    'strSQL = strSQL & vbCrLf & "MSysObjects.Name, FormatDateTime([DateCreate],2) AS DateCreated "
    strSQL = strSQL & vbCrLf & "FROM MSysObjects "
    strSQL = strSQL & vbCrLf & "WHERE Type IN (1, 5, 6, -32768, -32764, -32766, -32761) "
    strSQL = strSQL & vbCrLf & "AND Left$(Name, 4) <> ""MSys"" AND Left$(Name, 1) <> ""~"" "
    strSQL = strSQL & vbCrLf & "ORDER BY IIf(type=1,""Table"",IIf(type=6,""Linked Table"",IIf(type=5,""Query"",IIf(type=-32768,""Form"",IIf(type=-32764,""Report"",IIf(type=-32766,""Module"",IIf(type=-32761,""Module"",""Unknown""))))))), MSysObjects.Name;"

    'Debug.Print strSQL

    CurrentProject.Connection.Execute "GRANT SELECT ON MSysObjects TO Admin;"

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim rst As DAO.Recordset
    Set rst = dbs.OpenRecordset(strSQL)

    Do While Not rst.EOF
        If Not IsMissing(varDebug) Then Debug.Print rst.Fields(0), rst.Fields(1)
        Print #fle, """" & rst.Fields(0) & """" & "," & """" & rst.Fields(1) & """" & "," & """" & "DateCreated" & """"
        rst.MoveNext
    Loop
    Close fle
    'Stop
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
'330       If Err = 3167 Then          ' Record is deleted
'              'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputCatalogUserCreatedObjects of Class aegit_expClass", vbCritical, "ERROR"
'340           Resume e3167
'350       Else
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputCatalogUserCreatedObjects of Class aegit_expClass", vbCritical, "ERROR"
'370       End If
    'Stop
    Resume PROC_EXIT

End Sub

Private Sub OutputFieldLookupControlTypeList()
    Debug.Print "OutputFieldLookupControlTypeList"
    On Error GoTo 0
    Dim bln As Boolean
    bln = FieldLookupControlTypeList()
    'Debug.Print , "FieldLookupControlTypeList()=" & bln
    'Stop
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
        'Debug.Print "OutputListOfAccessApplicationOptions"
        'Debug.Print , "varDebug IS missing so no parameter is passed to OutputListOfAccessApplicationOptions"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print "OutputListOfAccessApplicationOptions"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to OutputListOfAccessApplicationOptions"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    If Not IsMissing(varDebug) Then Debug.Print "aegitSourceFolder=" & aegitSourceFolder

    Setup_The_Source_Location

    fle = FreeFile()
    Open mstrTheSourceLocation & "\" & aeAppOptions For Output As #fle

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
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfAccessApplicationOptions of Class aegit_expClass", vbCritical, "ERROR"
    End If
    Resume Next

End Sub

Private Sub OutputListOfAllHiddenQueries(Optional ByVal varDebug As Variant)
    ' Ref: http://www.pcreview.co.uk/forums/runtime-error-7874-a-t2922352.html
    ' Ref: http://www.pcreview.co.uk/threads/re-help-dirk-goldgar-or-someone-familiar-with-dev-ashish-search.3482377/
    ' Query Flag Description
    '   0 Select Query (Visible)
    '   8 Select Query (Hidden)
    '  16 Crosstab Query (Visible)
    '  24 Crosstab Query (Hidden)
    '  32 Delete Query (Visible)
    '  40 Delete Query (Hidden)
    '  48 Update Query (Visible)
    '  56 Update Query (Hidden)
    '  64 Append Query (Visible)
    '  72 Append Query (Hidden)
    '  80 Make Table Query (Visible)
    '  88 Make Table Query (Hidden)
    '  96 Data Definition Query (Visible)
    ' 104 Data Definition Query (Hidden)
    ' 112 Pass Through Query (Visible)
    ' 120 Pass Through Query (Hidden)
    ' 128 Union Query (Visible)
    ' 136 Union Query (Hidden)

    Debug.Print "OutputListOfAllHiddenQueries"
    On Error GoTo PROC_ERR

    ' MSysObjects list of types - Ref: http://allenbrowne.com/func-DDL.html - Query = 5
    ' Object Type
    ' Table 1
    ' Query 5
    ' Linked Table 4, 6, or 8
    ' Form -32768
    ' Report -32764

    Const strSQL As String = "SELECT m.Name, m.Flags, """" AS Description " & _
        "FROM MSysObjects AS m " & _
        "WHERE (((m.Name) Not Like ""~%"" And (m.Name) Not Like ""zzz*"") AND " & _
        "((m.Type)=5) AND ((m.Flags)=8 Or (m.Flags)=24 Or (m.Flags)=40 Or (m.Flags)=56 " & _
        "Or (m.Flags)=72 Or (m.Flags)=88 Or (m.Flags)=104 Or (m.Flags)=120 Or (m.Flags)=136))" & _
        "ORDER BY m.Name;"

    Dim fle As Integer
    fle = FreeFile()

    If aegitFrontEndApp Then
        'Debug.Print aestrSourceLocation & aeAppHiddQry
        Open aestrSourceLocation & aeAppHiddQry For Output As #fle
    Else
        'Debug.Print aestrSourceLocationBe & aeAppHiddQry
        Open aestrSourceLocationBe & aeAppHiddQry For Output As #fle
    End If

    CurrentProject.Connection.Execute "GRANT SELECT ON MSysObjects TO Admin;"

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim rst As DAO.Recordset
    Set rst = dbs.OpenRecordset(strSQL)

    Do While Not rst.EOF
        If Not IsMissing(varDebug) Then Debug.Print rst.Fields(0), rst.Fields(1)
        Print #fle, "[" & rst.Fields(0) & "]", rst.Fields(1)
        rst.MoveNext
    Loop
    Close fle

    If Not IsMissing(varDebug) Then Debug.Print strSQL
    If Not IsMissing(varDebug) And _
        Application.VBE.ActiveVBProject.Name = "aegit" Then
        'Debug.Print "IsQryHidden('qpt_Dummy') = " & IsQryHidden("qpt_Dummy")
        'Debug.Print "IsQryHidden('qry_HiddenDummy') = " & IsQryHidden("qry_HiddenDummy")
    End If

PROC_EXIT:
    rst.Close
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing
    'Stop
    Exit Sub

PROC_ERR:
    If Err = 3192 Then
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfAllHiddenQueries of Class aegit_expClass" & vbCrLf & vbCrLf & _
            "Could not create temp table. You do not have exclusive access to the database. You are not in developer mode? Compact/Repair and try the export again.", vbCritical, "ERROR"
        Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfAllHiddenQueries of Class aegit_expClass", vbCritical, "ERROR"
    End If
    Resume PROC_EXIT

End Sub

Private Sub OutputListOfApplicationProperties()
    ' Ref: http://www.granite.ab.ca/access/settingstartupoptions.htm

    Debug.Print "OutputListOfApplicationProperties"
    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim fle As Integer
    Dim strError As String

    Setup_The_Source_Location

    fle = FreeFile()
    Open mstrTheSourceLocation & "\" & aeAppListPrp For Output As #fle

    Dim i As Integer
    Dim strPropName As String
    Dim varPropValue As Variant
    Dim varPropType As Variant
    Dim varPropInherited As Variant

    With dbs
        For i = 0 To (.Properties.Count - 1)
            strError = vbNullString
            strPropName = .Properties(i).Name
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
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfApplicationProperties of Class aegit_expClass", vbCritical, "ERROR"
    End If
    Resume Next

End Sub

Private Sub OutputListOfCommandBarIDs(ByVal strOutputFile As String, Optional ByVal varDebug As Variant)
    ' Programming Office Commandbars - get the ID of a CommandBarControl
    ' Ref: http://blogs.msdn.com/b/guowu/archive/2004/09/06/225963.aspx
    ' Ref: http://www.vbforums.com/showthread.php?392954-How-do-I-Find-control-IDs-in-Visual-Basic-for-Applications-for-office-2003

    Debug.Print "OutputListOfCommandBarIDs"
    'Debug.Print , "strOutputFile = " & strOutputFile
    On Error GoTo PROC_ERR

    Dim CBR As Object       ' CommandBar
    Set CBR = Application.CommandBars
    Dim CBTN As Object      ' CommandBarButton
    Set CBTN = Application.CommandBars.FindControls
    Dim fle As Integer

    fle = FreeFile()
    Open strOutputFile For Output As #fle

    On Error Resume Next

    For Each CBR In Application.CommandBars
        For Each CBTN In CBR.Controls
            If Not IsMissing(varDebug) Then Debug.Print CBR.Name & ": " & CBTN.ID & " - " & CBTN.Caption
            Print #fle, CBR.Name & ": " & CBTN.ID & " - " & CBTN.Caption
        Next
    Next

PROC_EXIT:
    Close fle
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfCommandBarIDs of Class aegit_expClass", vbCritical, "ERROR"
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
        'Debug.Print , "varDebug IS missing so no parameter is passed to OutputListOfContainers"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to OutputListOfContainers"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    Set dbs = CurrentDb
    lngFileNum = FreeFile()

    If aegitFrontEndApp Then
        strFile = aestrSourceLocation & strTheFileName
    Else
        strFile = aestrSourceLocationBe & strTheFileName
    End If

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
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfContainers of Class aegit_expClass", vbCritical, "ERROR"
            Resume Next
    End Select
    OutputListOfContainers = False
    Resume Next

End Function

Private Sub OutputListOfForms(Optional ByVal varDebug As Variant)
    ' Ref: http://www.pcreview.co.uk/forums/runtime-error-7874-a-t2922352.html
    ' Ref: http://www.pcreview.co.uk/threads/re-help-dirk-goldgar-or-someone-familiar-with-dev-ashish-search.3482377/

    'Debug.Print "OutputListOfForms"
    On Error GoTo PROC_ERR

    ' MSysObjects list of types - Ref: http://allenbrowne.com/func-DDL.html - Query = 5
    ' http://stackoverflow.com/questions/3994956/meaning-of-msysobjects-values-32758-32757-and-3-microsoft-access
    ' Type TypeDesc
    ' -32768  Form
    ' -32766  Macro
    ' -32764  Reports
    ' -32761  Module
    ' -32758  Users
    ' -32757  Database Document
    ' -32756  Data Access Pages
    ' 1   Table - Local Access Tables
    ' 2   Access Object - Database
    ' 3   Access Object - Containers
    ' 4   Table - Linked ODBC Tables
    ' 5   Queries
    ' 6   Table - Linked Access Tables
    ' 8   SubDataSheets

    Const strSQL As String = "SELECT m.Name, """" AS Attribute " & _
        "FROM MSysObjects AS m " & _
        "WHERE m.Name Not Like ""~%"" And m.Name Not Like ""zzz*"" AND " & _
        "m.Type=-32768 " & _
        "ORDER BY m.Name;"
    
    Dim fle As Integer
    fle = FreeFile()

    If aegitFrontEndApp Then
        'Debug.Print aestrSourceLocation & aeAppListFrm
        Open aestrSourceLocation & aeAppListFrm For Output As #fle
    Else
        'Debug.Print aestrSourceLocationBe & aeAppListFrm
        Open aestrSourceLocationBe & aeAppListFrm For Output As #fle
    End If

    CurrentProject.Connection.Execute "GRANT SELECT ON MSysObjects TO Admin;"

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim rst As DAO.Recordset
    Set rst = dbs.OpenRecordset(strSQL)

    Do While Not rst.EOF
        If Not IsMissing(varDebug) Then Debug.Print rst.Fields(0)
        Print #fle, "[" & rst.Fields(0) & "]", IsFormHidden(rst.Fields(0))
        rst.MoveNext
    Loop
    Close fle

    If Not IsMissing(varDebug) Then
        Debug.Print "OutputListOfForms"
        Debug.Print strSQL
    End If

PROC_EXIT:
    rst.Close
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing
    'Stop
    Exit Sub

PROC_ERR:
    If Err = 3192 Then
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfForms of Class aegit_expClass" & vbCrLf & vbCrLf & _
            "Could not create temp table. You do not have exclusive access to the database. You are not in developer mode? Compact/Repair and try the export again.", vbCritical, "ERROR"
        Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfForms of Class aegit_expClass", vbCritical, "ERROR"
        Resume PROC_EXIT
    End If

End Sub

Private Sub OutputListOfIndexes(ByVal strFileOut As String)
    Debug.Print "OutputListOfIndexes"
    'Debug.Print , strFileOut
    On Error GoTo 0

    Dim fle As Integer
    fle = FreeFile()
    Open strFileOut For Output As #fle
    Close fle

    Dim tdf As DAO.TableDef

    For Each tdf In CurrentDb.TableDefs
        If Not (Left$(tdf.Name, 4) = "MSys" _
            Or Left$(tdf.Name, 4) = "~TMP" _
            Or Left$(tdf.Name, 3) = "zzz") Then
            OutputTableListOfIndexesDAO strFileOut, tdf
        End If
    Next
    Set tdf = Nothing
    'Stop
End Sub

Private Sub OutputListOfMacros(Optional ByVal varDebug As Variant)

    'Debug.Print "OutputListOfMacros"
    On Error GoTo PROC_ERR

    Const strSQL As String = "SELECT m.Name, """" AS Attribute " & _
        "FROM MSysObjects AS m " & _
        "WHERE m.Name Not Like ""~%"" And m.Name Not Like ""zzz*"" And m.Name Not Like ""~TMP*"" AND " & _
        "m.Type=-32766 " & _
        "ORDER BY m.Name;"
    
    Dim fle As Integer
    fle = FreeFile()

    If aegitFrontEndApp Then
        'Debug.Print aestrSourceLocation & aeAppListMac
        Open aestrSourceLocation & aeAppListMac For Output As #fle
    Else
        'Debug.Print aestrSourceLocationBe & aeAppListMac
        Open aestrSourceLocationBe & aeAppListMac For Output As #fle
    End If

    CurrentProject.Connection.Execute "GRANT SELECT ON MSysObjects TO Admin;"

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim rst As DAO.Recordset
    Set rst = dbs.OpenRecordset(strSQL)

    Do While Not rst.EOF
        If Not IsMissing(varDebug) Then Debug.Print rst.Fields(0)
        Print #fle, "[" & rst.Fields(0) & "]", IsMacroHidden(rst.Fields(0))
        rst.MoveNext
    Loop
    Close fle

    If Not IsMissing(varDebug) Then
        Debug.Print "OutputListOfMacros"
        Debug.Print strSQL
    End If

PROC_EXIT:
    rst.Close
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing
    'Stop
    Exit Sub

PROC_ERR:
    If Err = 3192 Then
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfMacros of Class aegit_expClass" & vbCrLf & vbCrLf & _
            "Could not create temp table. You do not have exclusive access to the database. You are not in developer mode? Compact/Repair and try the export again.", vbCritical, "ERROR"
        Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfMacros of Class aegit_expClass", vbCritical, "ERROR"
        Resume PROC_EXIT
    End If

End Sub

Private Sub OutputListOfModules(Optional ByVal varDebug As Variant)

    'Debug.Print "OutputListOfModules"
    On Error GoTo PROC_ERR

    Const strSQL As String = "SELECT m.Name, """" AS Attribute " & _
        "FROM MSysObjects AS m " & _
        "WHERE m.Name Not Like ""~%"" And m.Name Not Like ""zzz*"" AND " & _
        "m.Type=-32761 " & _
        "ORDER BY m.Name;"
    
    Dim fle As Integer
    fle = FreeFile()

    If aegitFrontEndApp Then
        'Debug.Print aestrSourceLocation & aeAppListMod
        Open aestrSourceLocation & aeAppListMod For Output As #fle
    Else
        'Debug.Print aestrSourceLocationBe & aeAppListMod
        Open aestrSourceLocationBe & aeAppListMod For Output As #fle
    End If

    CurrentProject.Connection.Execute "GRANT SELECT ON MSysObjects TO Admin;"

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim rst As DAO.Recordset
    Set rst = dbs.OpenRecordset(strSQL)

    Do While Not rst.EOF
        If Not IsMissing(varDebug) Then Debug.Print rst.Fields(0)
        Print #fle, "[" & rst.Fields(0) & "]", IsModuleHidden(rst.Fields(0))
        rst.MoveNext
    Loop
    Close fle

    If Not IsMissing(varDebug) Then
        Debug.Print "OutputListOfModules"
        Debug.Print strSQL
    End If

PROC_EXIT:
    rst.Close
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing
    'Stop
    Exit Sub

PROC_ERR:
    If Err = 3192 Then
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfModules of Class aegit_expClass" & vbCrLf & vbCrLf & _
            "Could not create temp table. You do not have exclusive access to the database. You are not in developer mode? Compact/Repair and try the export again.", vbCritical, "ERROR"
        Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfModules of Class aegit_expClass", vbCritical, "ERROR"
        Resume PROC_EXIT
    End If

End Sub

Private Sub OutputListOfReports(Optional ByVal varDebug As Variant)

    'Debug.Print "OutputListOfReports"
    On Error GoTo PROC_ERR

    Const strSQL As String = "SELECT m.Name, """" AS Attribute " & _
        "FROM MSysObjects AS m " & _
        "WHERE m.Name Not Like ""~%"" And m.Name Not Like ""zzz*"" AND " & _
        "m.Type=-32764 " & _
        "ORDER BY m.Name;"
    
    Dim fle As Integer
    fle = FreeFile()

    If aegitFrontEndApp Then
        'Debug.Print aestrSourceLocation & aeAppListRpt
        Open aestrSourceLocation & aeAppListRpt For Output As #fle
    Else
        'Debug.Print aestrSourceLocationBe & aeAppListRpt
        Open aestrSourceLocationBe & aeAppListRpt For Output As #fle
    End If

    CurrentProject.Connection.Execute "GRANT SELECT ON MSysObjects TO Admin;"

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim rst As DAO.Recordset
    Set rst = dbs.OpenRecordset(strSQL)

    Do While Not rst.EOF
        If Not IsMissing(varDebug) Then Debug.Print rst.Fields(0)
        Print #fle, "[" & rst.Fields(0) & "]", IsReportHidden(rst.Fields(0))
        rst.MoveNext
    Loop
    Close fle

    If Not IsMissing(varDebug) Then
        Debug.Print "OutputListOfReports"
        Debug.Print strSQL
    End If

PROC_EXIT:
    rst.Close
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing
    'Stop
    Exit Sub

PROC_ERR:
    If Err = 3192 Then
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfReports of Class aegit_expClass" & vbCrLf & vbCrLf & _
            "Could not create temp table. You do not have exclusive access to the database. You are not in developer mode? Compact/Repair and try the export again.", vbCritical, "ERROR"
        Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfReports of Class aegit_expClass", vbCritical, "ERROR"
        Resume PROC_EXIT
    End If

End Sub

Private Sub OutputListOfTables(blnNoODBC As Boolean, Optional ByVal varDebug As Variant)
    ' 1   Table - Local Access Tables
    ' 4   Table - Linked ODBC Tables
    ' 6   Table - Linked Access Tables

    Debug.Print "OutputListOfTables"
    On Error GoTo PROC_ERR

    Dim strSQL As String
    Const strAllTables As String = "(m.Type=1 OR m.Type=4 OR m.Type=6) "
    Const strAccessTables As String = "(m.Type=1 OR m.Type=6) "

    If blnNoODBC Then
        strSQL = "SELECT m.Name, """" AS Attribute " & _
            "FROM MSysObjects AS m " & _
            "WHERE m.Name Not Like ""~%"" AND m.Name Not Like ""zzz*"" AND " & _
            "m.Name Not Like ""~*"" AND " & _
            "m.Name Not Like ""MSys*"" AND " & _
            strAccessTables & _
            "ORDER BY m.Name;"
    Else
        strSQL = "SELECT m.Name, """" AS Attribute " & _
            "FROM MSysObjects AS m " & _
            "WHERE m.Name Not Like ""~%"" And m.Name Not Like ""zzz*"" AND " & _
            "m.Name Not Like ""~*"" AND " & _
            "m.Name Not Like ""MSys*"" AND " & _
            strAllTables & _
            "ORDER BY m.Name;"
    End If
    Debug.Print "StrSQL = " & strSQL
    'Stop

    Dim fle As Integer
    fle = FreeFile()

    If aegitFrontEndApp Then
        'Debug.Print aestrSourceLocation & aeAppListTbl
        Open aestrSourceLocation & aeAppListTbl For Output As #fle
    Else
        'Debug.Print aestrSourceLocationBe & aeAppListTbl
        Open aestrSourceLocationBe & aeAppListTbl For Output As #fle
    End If

    CurrentProject.Connection.Execute "GRANT SELECT ON MSysObjects TO Admin;"

    Dim i As Integer
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim rst As DAO.Recordset
    Set rst = dbs.OpenRecordset(strSQL)

    i = 0
    Do While Not rst.EOF
        If Not IsMissing(varDebug) Then Debug.Print rst.Fields(0)
        'Debug.Print "IsTableHidden(rst.Fields(0)) = " & IsTableHidden(rst.Fields(0))
        Print #fle, "[" & rst.Fields(0) & "]", IsTableHidden(rst.Fields(0))
        ReDim Preserve aeListOfTables(i)
        aeListOfTables(i) = rst.Fields(0)
        i = i + 1
        rst.MoveNext
    Loop
    Close fle

    'For i = 0 To UBound(aeListOfTables)
    '    Debug.Print i, aeListOfTables(i)
    'Next i
    'Stop

    If Not IsMissing(varDebug) Then
        Debug.Print "OutputListOfTables"
        Debug.Print strSQL
    End If

PROC_EXIT:
    rst.Close
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing
    'Stop
    Exit Sub

PROC_ERR:
    If Err = 3192 Then
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfTables of Class aegit_expClass" & vbCrLf & vbCrLf & _
            "Could not create temp table. You do not have exclusive access to the database. You are not in developer mode? Compact/Repair and try the export again.", vbCritical, "ERROR"
        Resume PROC_EXIT
    ElseIf Err = 3011 Then
        'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfTables of Class aegit_expClass", vbCritical, "ERROR"
        Resume Next
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfTables of Class aegit_expClass", vbCritical, "ERROR"
        Resume PROC_EXIT
    End If

End Sub

Public Sub OutputMyUnicode(ByRef strPathFileName As String, _
    ByVal arrUnicode As Variant)
    ' Ref: http://www.experts-exchange.com/Database/MS_Access/Q_26282187.html
    ' Ref: http://accessblog.net/2007/06/how-to-write-out-unicode-text-files-in.html

    Debug.Print "OutputMyUnicode"
    On Error GoTo PROC_ERR

    Dim i As Integer
    Dim MyStream As Object
    Set MyStream = CreateObject("ADODB.Stream")
    ' `It is summer in Geneva`, said Yu Zhou.
    ' strUnicode = "«" & Chr(160) & "C'est l'été à Genève" & Chr(160) & "»," _
    '    & " said " & ChrW(20446) & ChrW(-32225) & "."
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputMyUnicode of Class aegit_expClass", vbCritical, "ERROR"
    Resume PROC_EXIT

End Sub

Public Sub OutputPrinterInfo(Optional ByVal varDebug As Variant)
    ' Ref: http://msdn.microsoft.com/en-us/library/office/aa139946(v=office.10).aspx
    ' Ref: http://answers.microsoft.com/en-us/office/forum/office_2010-access/how-do-i-change-default-printers-in-vba/d046a937-6548-4d2b-9517-7f622e2cfed2

    Debug.Print "OutputPrinterInfo"
    On Error GoTo PROC_ERR

    Dim prt As Printer
    Dim prtCount As Integer
    Dim i As Integer
    Dim fle As Integer

    If Not mblnOutputPrinterInfo Then Exit Sub

    Setup_The_Source_Location

    fle = FreeFile()
    Open mstrTheSourceLocation & "\" & aePrnterInfo For Output As #fle

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
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputPrinterInfo of Class aegit_expClass", vbCritical, "ERROR"
            Resume Next
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputPrinterInfo of Class aegit_expClass", vbCritical, "ERROR"
            Resume Next
    End Select

End Sub

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

    Debug.Print "OutputQueriesSqlText"
    On Error GoTo PROC_ERR

    'Dim strTheSchemaFile As String
    If aegitFrontEndApp Then
        strFile = aestrSourceLocation & aeSqlTxtFile
    Else
        strFile = aestrSourceLocationBe & aeSqlTxtFile
    End If

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
            Print #1, "<<<[" & qdf.Name & "]>>>" & vbCrLf & qdf.sql
        End If
    Next

    OutputQueriesSqlText = True

PROC_EXIT:
    Set qdf = Nothing
    Set dbs = Nothing
    Close 1
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputQueriesSqlText of Class aegit_expClass", vbCritical, "ERROR"
    OutputQueriesSqlText = False
    Resume PROC_EXIT

End Function

Private Sub OutputTableDataAsFormattedText(ByVal strTblName As String, Optional ByVal varDebug As Variant)
    ' Ref: http://bytes.com/topic/access/answers/856136-access-2007-vba-select-external-data-ribbon

    On Error GoTo 0

    Dim strPathFileNameFD As String
    Dim strPathFileNameTT As String
    If aegitFrontEndApp Then
        strPathFileNameFD = aestrSourceLocation & strTblName & "_FormattedData.txt"
        strPathFileNameTT = aestrSourceLocation & strTblName & "_TransferText.txt"
    Else
        strPathFileNameFD = aestrSourceLocationBe & strTblName & "_FormattedData.txt"
        strPathFileNameTT = aestrSourceLocationBe & strTblName & "_TransferText.txt"
    End If

    If Not IsMissing(varDebug) Then
        Debug.Print "OutputTableDataAsFormattedText"
        Debug.Print , strPathFileNameFD
        Debug.Print , strPathFileNameTT
    Else
    End If
    ' AcFormat can be one of these AcFormat constants.
    ' acFormatASP
    ' acFormatDAP
    ' acFormatHTML
    ' acFormatIIS
    ' acFormatRTF
    ' acFormatSNP
    ' acFormatTXT
    ' acFormatXLS
    DoCmd.OutputTo acOutputTable, strTblName, acFormatTXT, strPathFileNameFD
    DoCmd.TransferText acExportDelim, , strTblName, strPathFileNameTT

End Sub

Private Sub OutputTableDataMacros(Optional ByVal varDebug As Variant)
    ' Ref: http://stackoverflow.com/questions/9206153/how-to-export-access-2010-data-macros

    Debug.Print "OutputTableDataMacros"
    On Error GoTo PROC_ERR

    Dim tdf As DAO.TableDef
    Dim strFile As String

    Dim strTheXMLDataLocation As String
    If aegitFrontEndApp Then
        strTheXMLDataLocation = aestrXMLDataLocation
    Else
        strTheXMLDataLocation = aestrXMLDataLocationBe
    End If

    For Each tdf In CurrentDb.TableDefs
        If Not IsLinkedTable(tdf.Name) Or _
            Not (Left$(tdf.Name, 4) = "MSys" _
            Or Left$(tdf.Name, 4) = "~TMP" _
            Or Left$(tdf.Name, 3) = "zzz") Then
            strFile = strTheXMLDataLocation & "tables_" & tdf.Name & "_DataMacro.xml"
            'Debug.Print "OutputTableDataMacros: strTheXMLDataLocation = " & strTheXMLDataLocation
            'Debug.Print "OutputTableDataMacros: strFile = " & strFile
            SaveAsText acTableDataMacro, tdf.Name, strFile
TwentyTwoTwenty:
            If Not IsMissing(varDebug) Then
                Debug.Print "OutputTableDataMacros:", tdf.Name, strTheXMLDataLocation, strFile
                PrettyXML strFile, varDebug
            Else
                PrettyXML strFile
            End If
        End If
NextTdf:
    Next tdf

PROC_EXIT:
    Set tdf = Nothing
    Exit Sub

PROC_ERR:
    If Err = 2950 Then ' Reserved Error
        Resume NextTdf
    ElseIf Err = 2220 Then
        Resume TwentyTwoTwenty
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputTableDataMacros of Class aegit_expClass", vbCritical, "ERROR"
    Resume PROC_EXIT

End Sub

Private Sub OutputTableListOfIndexesDAO(ByVal strFileOut As String, ByVal tdfIn As DAO.TableDef)
    'Debug.Print "OutputTableListOfIndexesDAO"
    On Error GoTo 0

    Dim fle As Integer
    fle = FreeFile()
    Open strFileOut For Append As #fle

    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim strIndexName As String
    Dim strFieldName As String

    'Debug.Print tdfIn.Name
    Print #fle, "<<<[" & tdfIn.Name & "]>>>"
    ' List values for each index
    For Each idx In tdfIn.Indexes
        ' List collection of fields the index contains
        strIndexName = "[" & idx.Name & "]"
        'Debug.Print , "Index:" & strIndexName
        Print #fle, , "Index:" & strIndexName
 
        For Each fld In idx.Fields
            'Debug.Print , , "Field Name:" & fld.Name
            Print #fle, , , "Field Name:" & fld.Name
            strFieldName = "[" & fld.Name & "], "
        Next fld
        'Debug.Print ">" & strIndexName, strFieldName
        Print #fle, ">" & strIndexName, strFieldName
    Next idx
    'Debug.Print "========================================"
    Print #fle, "========================================"
    Close fle

End Sub

Private Sub OutputTableProperties(Optional ByVal varDebug As Variant)
    ' Ref: http://bytes.com/topic/access/answers/709190-how-export-table-structure-including-description

    Debug.Print "OutputTableProperties"
    On Error GoTo PROC_ERR

    Dim tdf As DAO.TableDef
    Dim prp As DAO.Property
    Dim fldprp As DAO.Property
    Dim fld As DAO.Field

    If IsMissing(varDebug) Then
        'Debug.Print "OutputTableProperties"
        'Debug.Print , "varDebug IS missing so no parameter is passed to OutputTableProperties"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print "OutputTableProperties"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to OutputTableProperties"
        Debug.Print , "DEBUGGING TURNED ON"
    End If

    If Not IsMissing(varDebug) Then Debug.Print "aegitSourceFolder=" & aegitSourceFolder

    Setup_The_Source_Location

    For Each tdf In CurrentDb.TableDefs

        If Not (Left$(tdf.Name, 4) = "MSys" _
            Or Left$(tdf.Name, 4) = "~TMP" _
            Or Left$(tdf.Name, 3) = "zzz") Then

            Open mstrTheSourceLocation & "Properties_" & tdf.Name & ".txt" For Output As #1

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
            Print #1, "--------------------------------------------------"
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
    
PROC_EXIT:
    Set tdf = Nothing
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputTableProperties of Class aegit_expClass", vbCritical, "ERROR"
    Resume PROC_EXIT

End Sub

Private Sub OutputTheQAT(ByVal strTheFile As String, Optional ByVal varDebug As Variant)
    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=52635
    ' Set focus to the Access window

    Debug.Print "OutputTheQAT"
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

    Setup_The_Source_Location

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
    wsc.SendKeys mstrTheSourceLocation & strTheFile
    Delay 250
    wsc.SendKeys "%S"
    wsc.SendKeys "{ESC}"
    Pause 1

    ' NOTE: the QAT Ribbon XML as <mso:cmd app="Access" dt="0" /> at the start
    ' and it messes parsing as standard XML. Remove it, rename the file as .xml,
    ' save to the xml folder and prettify for reading.

    FixHeaderXML (mstrTheSourceLocation & strTheFile & ".exportedUI")
    'Stop

    If Not IsMissing(varDebug) Then
        PrettyXML mstrTheSourceLocation & strTheFile & ".exportedUI.fixed.xml", varDebug
    Else
        PrettyXML mstrTheSourceLocation & strTheFile & ".exportedUI.fixed.xml"
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    If Err = 3270 Then ' Property not found
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputTheQAT of Class aegit_expClass" & vbCrLf & _
            "No AppTitle found so QAT is not being exported!", vbInformation, "OutputTheQAT"
        Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputTheQAT of Class aegit_expClass", vbCritical, "ERROR"
        Resume Next
    End If

End Sub

Private Sub OutputTheSchemaFile(Optional ByVal varDebug As Variant) ' CreateDbScript()
    ' Ref: http://stackoverflow.com/questions/698839/how-to-extract-the-schema-of-an-access-mdb-database/9910716#9910716

    Debug.Print "OutputTheSchemaFile"
    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim ndx As DAO.Index
    Dim strSQL As String
    Dim strFlds As String
    Dim strLongFlds As String
    Dim blnLongFlds As Boolean
    Dim strCn As String
    Dim strLinkedTablePath As String

    Dim strTheSchemaFile As String
    If aegitFrontEndApp Then
        strTheSchemaFile = aestrSourceLocation & aeSchemaFile
    Else
        strTheSchemaFile = aestrSourceLocationBe & aeSchemaFile
    End If

    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim f As Object
    Set f = fs.CreateTextFile(strTheSchemaFile)

    strSQL = "Public Sub CreateTheDb()" & vbCrLf
    f.WriteLine strSQL
    strSQL = "Dim strSQL As String"
    f.WriteLine strSQL
    strSQL = "On Error GoTo PROC_ERR"
    f.WriteLine strSQL

    For Each tdf In dbs.TableDefs
        blnLongFlds = False
        If Not (Left$(tdf.Name, 4) = "MSys" _
            Or Left$(tdf.Name, 4) = "~TMP" _
            Or Left$(tdf.Name, 3) = "zzz") Then

            strLinkedTablePath = GetLinkedTableCurrentPath(tdf.Name)
            'Debug.Print "strLinkedTablePath = " & strLinkedTablePath
            If Left$(strLinkedTablePath, 13) <> "Local Table=>" Then
                f.WriteLine vbCrLf & "'OriginalLink=>" & strLinkedTablePath
            Else
                f.WriteLine vbCrLf & "'Local Table"
            End If

            strSQL = "strSQL=""CREATE TABLE [" & tdf.Name & "] ("
            strFlds = vbNullString

            For Each fld In tdf.Fields

                If Len(strFlds) <= 900 Then
                    strFlds = strFlds & ", [" & fld.Name & "] "
                Else    ' Hack to deal with 1024 limit for immediate window output
                    blnLongFlds = True
                    strFlds = strFlds & ", [" & fld.Name & "] " & """"
                    strLongFlds = strFlds
                    strFlds = vbNullString
                    'Stop
                End If

                ' Constants for complex types don't work prior to Access 2007
                Select Case fld.Type
                    Case dbText
                        ' No look-up fields
                        strFlds = strFlds & "Text (" & fld.size & ")"
                    Case 109&                                   ' dbComplexText
                        strFlds = strFlds & "Text (" & fld.size & ")"
                    Case dbMemo, dbByte, 102&, dbInteger, 103&, _
                        104&, dbSingle, 105&, dbDouble, 106&, dbGUID, _
                        107&, dbDecimal, 108&, dbCurrency, _
                        101&, dbBinary
                        strFlds = strFlds & FieldTypeName(fld)
                    Case dbLong
                        If (fld.Attributes And dbAutoIncrField) = 0& Then
                            strFlds = strFlds & "Long"
                        Else
                            strFlds = strFlds & "Counter"
                        End If
                        ' Case dbGUID
                        '    strFlds = strFlds & "GUID"
                        '    'strFlds = strFlds & "Replica"
                    Case dbDate
                        strFlds = strFlds & "DateTime"
                    Case dbBoolean
                        strFlds = strFlds & "YesNo"
                    Case dbLongBinary
                        strFlds = strFlds & "OLEObject"
                    Case Else
                        MsgBox "Unknown fld.Type=" & fld.Type & " in procedure OutputTheSchemaFile of aegit_expClass", vbCritical, "ERROR"
                        Debug.Print "Unknown fld.Type=" & fld.Type & " in procedure OutputTheSchemaFile of aegit_expClass" & vbCrLf & _
                            "tdf.Name=" & tdf.Name & " strFlds=" & strFlds
                End Select

            Next

            'Debug.Print Len(strLongFlds), strLongFlds
            'Debug.Print Len(strFlds), strFlds
            If Not blnLongFlds Then
                strSQL = strSQL & Mid$(strFlds, 2) & " )""" & vbCrLf & "Currentdb.Execute strSQL"
                f.WriteLine vbCrLf & strSQL
            Else
                strSQL = strSQL & Mid$(strLongFlds, 2)
                f.WriteLine vbCrLf & strSQL
                strSQL = "strSQL=strSQL & " & """" & strFlds & " )""" & vbCrLf & "Currentdb.Execute strSQL"
                f.WriteLine strSQL
            End If
            If Not IsMissing(varDebug) Then Debug.Print strSQL
            'Stop

            ' Create Indexes here
            
            ' Create relationships where?

        End If
    Next
    'Stop

    'strSQL = vbCrLf & "Debug.Print " & """" & "DONE !!!" & """"
    'f.WriteLine strSQL
    f.WriteLine
    f.WriteLine "'Access 2010 - Compact And Repair"
    strSQL = "SendKeys " & """" & "%F{END}{ENTER}%F{TAB}{TAB}{ENTER}" & """" & ", False"
    f.WriteLine strSQL
    f.WriteLine "Exit Sub"
    f.WriteLine "PROC_ERR:"
    f.WriteLine "If Err = 3010 Then Resume Next"
    f.WriteLine "If Err = 3283 Then Resume Next"
    f.WriteLine "If Err = 3375 Then Resume Next"
    'MsgBox "Erl=" & Erl & vbCrLf & "Err.Number=" & Err.Number & vbCrLf & "Err.Description=" & Err.Description
    strSQL = "MsgBox " & """" & "Erl=" & """" & " & Erl & vbCrLf & " & _
        """" & "Err.Number=" & """" & " & Err.Number & vbCrLf & " & _
        """" & "Err.Description=" & """" & " & Err.Description"
    f.WriteLine strSQL & vbCrLf
    f.WriteLine "End Sub"

    f.Close
    'Debug.Print "DONE !!!"

PROC_EXIT:
    Set dbs = Nothing
    Set fs = Nothing
    Exit Sub

PROC_ERR:
    Select Case Err
        Case Else
            MsgBox "Erl=" & Erl & " Err=" & Err.Number & " (" & Err.Description & ") in procedure OutputTheSchemaFile of Class aegitClass"
            Resume PROC_EXIT
    End Select

End Sub

Private Sub OutputTheSqlFile(ByVal strFileIn As String, ByVal strFileOut As String)
    On Error GoTo 0
    ReadInputWriteOutputFileSql strFileIn, strFileOut
End Sub

Private Sub OutputTheSqlOnlyFile(ByVal strFileIn As String, ByVal strFileOut As String)
    'Debug.Print "OutputTheSqlOnlyFile"
    'Debug.Print , "strFileIn=" & strFileIn
    'Debug.Print , "strFileOut=" & strFileOut
    On Error GoTo 0
    'Stop
    ReadInputWriteOutputSqlSchemaOnlyFile strFileIn, strFileOut
End Sub

Private Sub OutputTheTableDataAsXML(ByRef avarTableNames() As Variant, Optional ByVal varDebug As Variant)
    ' Ref: http://wiki.lessthandot.com/index.php/Output_Access_/_Jet_to_XML
    ' Ref: http://msdn.microsoft.com/en-us/library/office/aa164887(v=office.10).aspx

    Debug.Print "OutputTheTableDataAsXML"
    On Error GoTo PROC_ERR

    Const adOpenStatic As Integer = 3
    Const adLockOptimistic As Integer = 3
    Const adPersistXML As Integer = 1

    Dim i As Integer
    Dim strFileName As String
    Dim strSQL As String
    Dim strTheXMLDataLocation As String

    If aegitXMLDataFolder = "default" Then
        strTheXMLDataLocation = aegitType.XMLDataFolder
    ElseIf aegitFrontEndApp Then
        strTheXMLDataLocation = aestrXMLDataLocation
    ElseIf Not aegitFrontEndApp Then
        strTheXMLDataLocation = aestrXMLDataLocationBe
    End If

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

            strFileName = strTheXMLDataLocation & avarTableNames(i) & ".xml"

            If aegitSetup Then
                If Not IsMissing(varDebug) Then Debug.Print "aegitSetup=True XML Data Location=" & strTheXMLDataLocation
            Else
                If Not IsMissing(varDebug) Then Debug.Print "aegitSetup=False XML Data Location=" & strTheXMLDataLocation
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
                "strFileName = " & strFileName & vbCrLf & "in procedure OutputTheTableDataAsXML of Class aegit_expClass", vbCritical, "ERROR"
            'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputTheTableDataAsXML of Class aegit_expClass"
    End Select
    Resume PROC_EXIT

End Sub

Private Function PassFail(ByVal blnPassFail As Boolean, Optional ByVal varOther As Variant) As String
    On Error GoTo 0
    If Not IsMissing(varOther) Then
        PassFail = "NotUsed"
        Exit Function
    End If
    If blnPassFail Then
        PassFail = "Pass"
    Else
        PassFail = "Fail"
    End If
End Function

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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Pause of Class aegit_expClass", vbCritical, "ERROR"
    Resume PROC_EXIT

End Function

Public Sub PrettyXML(ByVal strPathFileName As String, Optional ByVal varDebug As Variant)

    'Debug.Print "PrettyXML"
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
    If Left$(strPathFileName, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid$(strPathFileName, 2, Len(strPathFileName) - 1)
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
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrettyXML of Class aegit_expClass", vbCritical, "ERROR"
            'If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrettyXML of Class aegit_expClass"
    End Select
    Resume PROC_EXIT

End Sub

Private Sub ReadInputWriteOutputFileSql(ByVal strFileIn As String, ByVal strFileOut As String)

    'Debug.Print "ReadInputWriteOutputFile"
    On Error GoTo 0

    Dim fleIn As Integer
    Dim fleOut As Integer
    Dim strIn As String
    Dim i As Integer

    fleIn = FreeFile()
    Open strFileIn For Input As #fleIn

    fleOut = FreeFile()
    Open strFileOut For Output As #fleOut

    i = 0
    Do While Not EOF(fleIn)
        i = i + 1
        Line Input #fleIn, strIn
        If FoundSqlKeywordInLine(strIn) Then
            'Debug.Print i & ">", strIn
            Print #fleOut, strIn
        Else
            'Debug.Print i, strIn
        End If
    Loop
    'Debug.Print "DONE !!!"

    Close fleIn
    Close fleOut

End Sub

Public Sub ReadInputWriteOutputLovefieldSchema(ByVal strFileIn As String, ByVal strFileOut As String)
    ' Ref: https://github.com/google/lovefield/blob/master/docs/spec/01_schema.md
    ' Type, Default Value, Nullable by default, Description
    ' lf.Type.ARRAY_BUFFER, null, Yes, JavaScript ArrayBuffer object
    ' lf.Type.BOOLEAN, false, No, JavaScript boolean object
    ' lf.Type.DATE_TIME, Date(0), No, JavaScript Date - will be converted to timestamp integer internally
    ' lf.Type.INTEGER, 0, No, 32-bit integer
    ' lf.Type.NUMBER, 0, No, JavaScript number type
    ' lf.Type.String, '', No, JavaScript string type
    ' lf.Type.OBJECT, null, Yes, JavaScript Object - stored as-is

    'Debug.Print "ReadInputWriteOutputLovefieldSchema"
    On Error GoTo PROC_ERR

    Const SEMICOLON As String = ";"
    Const PERIOD As String = "."
    Dim fleIn As Integer
    Dim fleOut As Integer
    Dim i As Integer
    Dim strLfCreateTable As String
    Dim strFieldInfoToParse As String
    Dim strFieldName As String
    Dim strAccFieldType As String
    Dim strLfFieldType As String
    Dim strLfFieldName As String
    Dim strThePrimaryKeyField As String
    Dim strTheIndex As String
    Dim strTableName As String

    Dim strAppName As String
    strAppName = Application.VBE.ActiveVBProject.Name
    Dim strLfBegin As String
    strLfBegin = "// Begin schema creation" & vbCrLf & "var schemaBuilder = lf.schema.create('" & strAppName & "', 1);"

    fleOut = FreeFile()
    Open strFileOut For Output As #fleOut

    Dim arrSQL() As String
    i = 0
    fleIn = FreeFile()
    Open strFileIn For Input As #fleIn
    Do While Not EOF(fleIn)
        ReDim Preserve arrSQL(i)
        Line Input #fleIn, arrSQL(i)
        i = i + 1
    Loop
    Close fleIn

    'For i = 0 To UBound(arrSQL)
    '    Debug.Print i & ">", arrSQL(i)
    'Next

    'Debug.Print strLfBegin

    For i = 0 To UBound(arrSQL)
        If Left$(arrSQL(i), Len(mTABLE)) = mTABLE Then
            ' Get the table name
            strTableName = GetTableName(arrSQL(i))
            ' Test if the table schema is finished
            If i < UBound(arrSQL) Then
                ' FIX THIS !!!
            End If
            strLfCreateTable = "schemaBuilder.createTable('" & strTableName & "')."
            'Debug.Print i & ">", strLfCreateTable
            Print #fleOut, strLfCreateTable
            mstrToParse = Right$(arrSQL(i), Len(arrSQL(i)) - InStr(arrSQL(i), "("))
            Do While mstrToParse <> vbNullString
                ' Create the table
                strFieldInfoToParse = GetFieldInfo(mstrToParse)
                'Debug.Print strLfCreateTable
                'Debug.Print , mstrToParse
                'Debug.Print , "strFieldInfoToParse=" & strFieldInfoToParse, Len(strFieldInfoToParse)
                strFieldName = Trim$(Left$(strFieldInfoToParse, InStrRev(strFieldInfoToParse, "] ")))
                'strAccFieldType = Trim$(Right$(strFieldInfoToParse, Len(strFieldInfoToParse) - InStr(strFieldInfoToParse, " ")))
                strAccFieldType = Trim$(Right$(strFieldInfoToParse, Len(strFieldInfoToParse) - InStrRev(strFieldInfoToParse, "] ")))
                'Debug.Print , "strFieldName=" & strFieldName, Len(strFieldName), "strAccFieldType=" & strAccFieldType, Len(strAccFieldType)
                strLfFieldType = GetLovefieldType(strAccFieldType)
                strLfFieldName = Space$(4) & "addColumn('" & Mid$(strFieldName, 2, Len(strFieldName) - 2)
                'Debug.Print , strLfFieldName & strLfFieldType, "{" & strAccFieldType & "}"
                Print #fleOut, strLfFieldName & strLfFieldType
                'Stop
            Loop
        ElseIf Left$(arrSQL(i), Len(mPRIMARYKEY)) = mPRIMARYKEY Then
            ' Create the PrimaryKey
            strThePrimaryKeyField = GetPrimaryKey(arrSQL(i))
            If i <> UBound(arrSQL) Then
                If IsTableSchemaDone(strTableName, arrSQL(i + 1)) Then
                    strThePrimaryKeyField = strThePrimaryKeyField & SEMICOLON
                Else
                    strThePrimaryKeyField = strThePrimaryKeyField & PERIOD
                End If
            End If
            If i = UBound(arrSQL) Then
                'Debug.Print i & ">", strThePrimaryKeyField & SEMICOLON
                Print #fleOut, strThePrimaryKeyField & SEMICOLON
            Else
                'Debug.Print i & ">", strThePrimaryKeyField
                Print #fleOut, strThePrimaryKeyField
            End If
            'Stop
        ElseIf Left$(arrSQL(i), Len(mINDEX)) = mINDEX Then
            ' Create the Index
            strTheIndex = GetIndex(arrSQL(i))
            If i <> UBound(arrSQL) Then
                If IsTableSchemaDone(strTableName, arrSQL(i + 1)) Then
                    strTheIndex = strTheIndex & SEMICOLON
                Else
                    strTheIndex = strTheIndex & PERIOD
                End If
            End If
            If i = UBound(arrSQL) Then
                'Debug.Print i & ">", strTheIndex & SEMICOLON
                Print #fleOut, strTheIndex & SEMICOLON
            Else
                'Debug.Print i & ">", strTheIndex
                Print #fleOut, strTheIndex
            End If
        End If
    Next
    'Debug.Print "DONE !!!"

PROC_EXIT:
    Close fleIn
    Close fleOut
    Exit Sub

PROC_ERR:
    Select Case Err
        Case 5
            Debug.Print "mstrToParse=" & mstrToParse
            MsgBox "Erl=" & Erl & " Err=" & Err.Number & " (" & Err.Description & ") in procedure ReadInputWriteOutputLovefieldSchema of Class aegitClass"
            Stop
            Resume PROC_EXIT
        Case Else
            MsgBox "Erl=" & Erl & " Err=" & Err.Number & " (" & Err.Description & ") in procedure ReadInputWriteOutputLovefieldSchema of Class aegitClass"
            Resume PROC_EXIT
    End Select

End Sub

Private Sub ReadInputWriteOutputSqlSchemaOnlyFile(ByVal strFileIn As String, ByVal strFileOut As String)

    'Debug.Print "ReadInputWriteOutputSqlSchemaOnlyFile"
    On Error GoTo PROC_ERR

    Dim fleIn As Integer
    Dim fleOut As Integer
    Dim i As Integer
    Dim strSqlA As String
    Dim strSqlB As String

    fleOut = FreeFile()
    Open strFileOut For Output As #fleOut

    Dim arrSQL() As String
    i = 0
    fleIn = FreeFile()
    Open strFileIn For Input As #fleIn
    Do While Not EOF(fleIn)
        ReDim Preserve arrSQL(i)
        Line Input #fleIn, arrSQL(i)
        i = i + 1
    Loop
    Close fleIn

    'For i = 0 To UBound(arrSQL)
    '    Debug.Print i & ">", arrSQL(i)
    'Next

    For i = 0 To UBound(arrSQL)
        If (i <> UBound(arrSQL)) Then
            If Left$(arrSQL(i + 1), 16) = "strSQL=strSQL & " Then
                If Left$(arrSQL(i), 7) = "strSQL=" Then
                    strSqlA = Right$(arrSQL(i), Len(arrSQL(i)) - 8)
                    strSqlA = Left$(strSqlA, Len(strSqlA) - 1)
                    'Debug.Print i & ">", "strSqlA=" & strSqlA
                    strSqlB = Right$(arrSQL(i + 1), Len(arrSQL(i + 1)) - 17)
                    strSqlB = Left$(strSqlB, Len(strSqlB) - 1)
                    'Debug.Print i & ">", "strSqlB=" & strSqlB
                    Print #fleOut, strSqlA & strSqlB
                    i = i + 1
                End If
            ElseIf Left$(arrSQL(i), 7) = "strSQL=" Then
                strSqlA = Right$(arrSQL(i), Len(arrSQL(i)) - 8)
                strSqlA = Left$(strSqlA, Len(strSqlA) - 1)
                'Debug.Print i, strSqlA
                Print #fleOut, strSqlA
            End If
        Else
            If Left$(arrSQL(i), 7) = "strSQL=" Then
                strSqlA = Right$(arrSQL(i), Len(arrSQL(i)) - 8)
                strSqlA = Left$(strSqlA, Len(strSqlA) - 1)
                'Debug.Print i, strSqlA
                Print #fleOut, strSqlA
            End If
            'Debug.Print "UBound"
        End If
    Next
    'Debug.Print "DONE !!!"

PROC_EXIT:
    Close fleIn
    Close fleOut
    Exit Sub

PROC_ERR:
    Select Case Err
        Case Else
            MsgBox "Erl=" & Erl & " Err=" & Err.Number & " (" & Err.Description & ") in procedure ReadInputWriteOutputSqlSchemaOnlyFile of Class aegitClass"
            Resume PROC_EXIT
    End Select

End Sub

Private Function RecordsetUpdatable(ByVal strSQL As String) As Boolean
    ' Ref: http://msdn.microsoft.com/en-us/library/office/ff193796(v=office.15).aspx

    Debug.Print "RecordsetUpdatable"

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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RecordsetUpdatable of Class aegit_expClass", vbCritical, "ERROR"
    Resume Next

End Function

Private Sub ResetWorkspace()
    Debug.Print "ResetWorkspace"
    On Error Resume Next

    Application.MenuBar = vbNullString
    DoCmd.SetWarnings False
    DoCmd.Hourglass False
    DoCmd.Echo True

    Dim intCounter As Integer
    ' Clean up workspace by closing open forms and reports
    For intCounter = 0 To Forms.Count - 1
        DoCmd.Close acForm, Forms(intCounter).Name
    Next intCounter

    For intCounter = 0 To Reports.Count - 1
        DoCmd.Close acReport, Reports(intCounter).Name
    Next intCounter
End Sub

Private Function SingleTableIndexSummary(ByVal tdf As DAO.TableDef) As String

    Dim strIndexFieldInfo As String
    strIndexFieldInfo = vbNullString
    Dim fld As DAO.Field

    For Each fld In tdf.Fields
        strIndexFieldInfo = strIndexFieldInfo & DescribeIndexField(tdf, fld.Name)
        'Debug.Print fld.Name, "strIndexFieldInfo=" & strIndexFieldInfo
    Next
    'Debug.Print "TableIndexSummary", tdf.Name, "strIndexFieldInfo=" & strIndexFieldInfo
    SingleTableIndexSummary = strIndexFieldInfo

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

    'Debug.Print "SizeString"
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

Private Sub SortTheFile(ByVal strInFile As String, ByVal strOutFile As String)
    ' Ref: http://www.vbaexpress.com/forum/showthread.php?46362-Sort-contents-in-a-text-file
    ' Ref: http://www.vbaexpress.com/forum/showthread.php?48491Function ArrayListSort(sn As Variant, Optional bAscending As Boolean = True)

    On Error GoTo PROC_ERR
    Debug.Print "SortTheFile"
    'Debug.Print , "strInFile = " & strInFile
    'Debug.Print , "strOutFile = " & strOutFile
    If Dir$(strInFile) = vbNullString Then Stop

    Dim str As String
    Dim ary() As Variant
    Dim L As Long

    Open strInFile For Input As #1
    Open strOutFile For Output As #2

    With CreateObject("System.Collections.ArrayList")
        Do Until EOF(1)
            Line Input #1, str
            .Add Trim$(CStr(str))
        Loop
        .Sort
        ary = .ToArray
        For L = LBound(ary) To UBound(ary)
            Print #2, ary(L)
        Next
    End With


PROC_EXIT:
    Close #1
    Close #2
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case 53
            MsgBox "Erl=" & Erl & " Err=" & Err.Number & " (" & Err.Description & ") in procedure SortTheFile of Class aegit_expClass", vbCritical, "ERROR"
            Resume PROC_EXIT
        Case Else
            MsgBox "Erl=" & Erl & " Err=" & Err.Number & " (" & Err.Description & ") in procedure SortTheFile of Class aegit_expClass", vbCritical, "ERROR"
            Resume Next
    End Select

End Sub

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

    'Debug.Print "TableInfo"
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
        If Left$(strLinkedTablePath, 13) <> "Local Table=>" Then
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
    If Left$(strLinkedTablePath, 13) <> "Local Table=>" Then
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

    If Not aegitExport.ExportNoODBCTablesInfo Then
        For Each fld In tdf.Fields
            If Not IsMissing(varDebug) Then
                'If Not IsMissing(varDebug) And aeintFDLen <> 11 Then
                Debug.Print SizeString(fld.Name, aeintFNLen, TextLeft, " ") _
                    & aestr4 & SizeString(FieldTypeName(fld), aeintFTLen, TextLeft, " ") _
                    & aestr4 & SizeString(fld.size, aeintFSize, TextLeft, " ") _
                    & aestr4 & SizeString(GetDescription(fld), aeintFDLen, TextLeft, " ")
            End If
            Print #1, SizeString(fld.Name, aeintFNLen, TextLeft, " ") _
                & aestr4 & SizeString(FieldTypeName(fld), aeintFTLen, TextLeft, " ") _
                & aestr4 & SizeString(fld.size, aeintFSize, TextLeft, " ") _
                & aestr4 & SizeString(GetDescription(fld), aeintFDLen, TextLeft, " ")
        Next
    Else
        If IsLinkedODBC(strTableName) Then
            ' Do nothing
        Else
            For Each fld In tdf.Fields
                If Not IsMissing(varDebug) Then
                    'If Not IsMissing(varDebug) And aeintFDLen <> 11 Then
                    Debug.Print SizeString(fld.Name, aeintFNLen, TextLeft, " ") _
                        & aestr4 & SizeString(FieldTypeName(fld), aeintFTLen, TextLeft, " ") _
                        & aestr4 & SizeString(fld.size, aeintFSize, TextLeft, " ") _
                        & aestr4 & SizeString(GetDescription(fld), aeintFDLen, TextLeft, " ")
                End If
                Print #1, SizeString(fld.Name, aeintFNLen, TextLeft, " ") _
                    & aestr4 & SizeString(FieldTypeName(fld), aeintFTLen, TextLeft, " ") _
                    & aestr4 & SizeString(fld.size, aeintFSize, TextLeft, " ") _
                    & aestr4 & SizeString(GetDescription(fld), aeintFDLen, TextLeft, " ")
            Next
        End If
    End If

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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TableInfo of Class aegit_expClass", vbCritical, "ERROR"
    If Not IsMissing(varDebug) Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TableInfo of Class aegit_expClass"
    TableInfo = False
    Resume PROC_EXIT

End Function

Private Sub TestForRelativePath()
    On Error GoTo 0
    Debug.Print "TestForRelativePath"
    ' Test for relative path
    Dim strTestPath As String
    strTestPath = aestrSourceLocation
    If Left$(aestrSourceLocation, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid$(aestrSourceLocation, 2, Len(aestrSourceLocation) - 1)
        aestrSourceLocation = strTestPath
    End If
    'Debug.Print , "aestrSourceLocation = " & aestrSourceLocation
    '
    strTestPath = aestrSourceLocationBe
    If Left$(aestrSourceLocationBe, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid$(aestrSourceLocationBe, 2, Len(aestrSourceLocationBe) - 1)
        aestrSourceLocationBe = strTestPath
    End If
    'Debug.Print , "aestrSourceLocationBe = " & aestrSourceLocationBe
    '
    strTestPath = aestrXMLLocation
    If Left$(aestrXMLLocation, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid$(aestrXMLLocation, 2, Len(aestrXMLLocation) - 1)
        aestrXMLLocation = strTestPath
    End If
    'Debug.Print , "aestrXMLLocation = " & aestrXMLLocation
    '
    strTestPath = aestrXMLLocationBe
    If Left$(aestrXMLLocationBe, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid$(aestrXMLLocationBe, 2, Len(aestrXMLLocationBe) - 1)
        aestrXMLLocationBe = strTestPath
    End If
    'Debug.Print , "aestrXMLLocationBe = " & aestrXMLLocationBe
    '
    strTestPath = aestrXMLDataLocation
    If Left$(aestrXMLDataLocation, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid$(aestrXMLDataLocation, 2, Len(aestrXMLDataLocation) - 1)
        aestrXMLDataLocation = strTestPath
    End If
    'Debug.Print , "aestrXMLDataLocation = " & aestrXMLDataLocation
    '
    strTestPath = aestrXMLDataLocationBe
    If Left$(aestrXMLDataLocationBe, 1) = "." Then
        strTestPath = CurrentProject.Path & Mid$(aestrXMLDataLocationBe, 2, Len(aestrXMLDataLocationBe) - 1)
        aestrXMLDataLocationBe = strTestPath
    End If
    'Debug.Print , "aestrXMLDataLocationBe = " & aestrXMLDataLocationBe
    'Debug.Print , "--------------------------------------------------"

End Sub

Private Sub VerifySetup()   '(Optional ByVal varDebug As Variant)
    On Error GoTo 0
    Debug.Print "VerifySetup"
    'Debug.Print , "aegitFrontEndApp = " & aegitFrontEndApp
    'Debug.Print , "aegitSourceFolder = " & aegitSourceFolder

    ' Test for aegit setup
    If aegitSourceFolder = "default" Then
        aegitSetup = True
        aestrSourceLocation = aegitType.SourceFolder
        aestrXMLLocation = aegitType.XMLFolder
        aestrXMLDataLocation = aegitType.XMLDataFolder
        'Debug.Print , "aegitSetup = True"
        'Debug.Print , "aegitSourceFolder = ""default"""
        'Debug.Print , "aestrSourceLocation = " & aestrSourceLocation
        'Debug.Print , "--------------------------------------------------"

        TestForRelativePath

        ' Check folders exist
        If Not FolderExists(aestrSourceLocation) Then
            MsgBox "aestrSourceLocation does not exist!", vbCritical, "VerifySetup"
            Stop
        End If
        'Debug.Print , "aestrXMLLocation = " & aestrXMLLocation
        If Not FolderExists(aestrXMLLocation) Then
            MsgBox "aestrXMLLocation does not exist!", vbCritical, "VerifySetup"
            Stop
        End If
        'Debug.Print , "aestrXMLDataLocation = " & aestrXMLDataLocation
        If Not FolderExists(aestrXMLDataLocation) Then
            MsgBox "aestrXMLDataLocation does not exist!", vbCritical, "VerifySetup"
            Stop
        End If
    ElseIf aegitFrontEndApp Then
        aestrSourceLocation = aegitSourceFolder
        aestrSourceLocationBe = aegitSourceFolderBe
        aestrXMLLocation = aegitXMLFolder
        aestrXMLLocationBe = aegitXMLFolderBe
        aestrXMLDataLocation = aegitXMLDataFolder
        aestrXMLDataLocationBe = aegitXMLDataFolderBe
        'Debug.Print , "aestrSourceLocation = " & aestrSourceLocation
        'Debug.Print , "aestrSourceLocationBe = " & aestrSourceLocationBe
        'Debug.Print , "aestrXMLLocation = " & aestrXMLLocation
        'Debug.Print , "aestrXMLLocationBe = " & aestrXMLLocationBe
        'Debug.Print , "aestrXMLDataLocation = " & aestrXMLDataLocation
        'Debug.Print , "aestrXMLDataLocationBe = " & aestrXMLDataLocationBe
        'Debug.Print , "--------------------------------------------------"

        TestForRelativePath

        ' Check folders exist
        If Not FolderExists(aestrSourceLocation) Then
            MsgBox "aestrSourceLocation does not exist!", vbCritical, "VerifySetup"
            Stop
        End If
        'Debug.Print , "aestrXMLLocation = " & aestrXMLLocation
        If Not FolderExists(aestrXMLLocation) Then
            MsgBox "aestrXMLLocation does not exist!", vbCritical, "VerifySetup"
            Stop
        End If
        'Debug.Print , "aestrXMLDataLocation = " & aestrXMLDataLocation
        If Not FolderExists(aestrXMLDataLocation) Then
            MsgBox "aestrXMLDataLocation does not exist!", vbCritical, "VerifySetup"
            Stop
        End If
    ElseIf Not aegitFrontEndApp Then
        aestrSourceLocationBe = aegitSourceFolderBe
        aestrXMLLocationBe = aegitXMLFolderBe
        aestrXMLDataLocationBe = aegitXMLDataFolderBe
        'Debug.Print , "aestrSourceLocationBe = " & aestrSourceLocationBe
        'Debug.Print , "aestrXMLLocationBe = " & aestrXMLLocationBe
        'Debug.Print , "aestrXMLDataLocationBe = " & aestrXMLDataLocationBe
        'Debug.Print , "--------------------------------------------------"

        TestForRelativePath

        ' Check folders exist
        If Not FolderExists(aestrSourceLocationBe) Then
            MsgBox "aestrSourceLocationBe does not exist!" & vbCrLf & _
                "aestrSourceLocationBe = " & aestrSourceLocationBe, vbCritical, "VerifySetup"
            Stop
        End If
        If Not FolderExists(aestrXMLLocationBe) Then
            MsgBox "aestrXMLLocationBe does not exist!" & vbCrLf & _
                "aestrXMLLocationBe = " & aestrXMLLocationBe, vbCritical, "VerifySetup"
            Stop
        End If
        If Not FolderExists(aestrXMLDataLocationBe) Then
            MsgBox "aestrXMLDataLocationBe does not exist!" & vbCrLf & _
                "aestrXMLDataLocationBe = " & aestrXMLDataLocationBe, vbCritical, "VerifySetup"
            Stop
        End If
    End If

    ' Final paths are absolute
    'Debug.Print "VerifySetup"
    'Debug.Print , ">==> Final Paths >==>"
    'Debug.Print , "Property Get SourceFolder:       aestrSourceLocation = " & aestrSourceLocation
    'Debug.Print , "Property Get SourceFolderBe:     aestrSourceLocationBe = " & aestrSourceLocationBe
    'Debug.Print , "Property Get XMLFolder:          aestrXMLLocation = " & aestrXMLLocation
    'Debug.Print , "Property Get XMLFolderBe:        aestrXMLLocationBe = " & aestrXMLLocationBe
    'Debug.Print , "Property Get XMLDataFolder:      aestrXMLDataLocation = " & aestrXMLDataLocation
    'Debug.Print , "Property Get XMLDataFolderBe:    aestrXMLDataLocationBe = " & aestrXMLDataLocationBe
    'Debug.Print , "--------------------------------------------------"

    '???
    If aestrBackEndDbOne = vbNullString Then
        MsgBox "aestrBackEndDbOne is not set!", vbCritical, "VerifySetup"
        Stop
    End If

    'Debug.Print , "aegitDataXML(0) = " & aegitDataXML(0)

    If aestrBackEndDbOne <> "default" Then
        OpenAllDatabases True
    End If
    'Debug.Print , "Property Get BackEndDbOne = " & aestrBackEndDbOne
    'Stop

End Sub

Private Sub WaitSeconds(ByVal intSeconds As Integer)
    ' Ref: http://www.fmsinc.com/MicrosoftAccess/modules/examples/AvoidDoEvents.asp
    ' ====================================================================
    ' Comments:  Waits for a specified number of seconds
    ' Parameter: intSeconds, Number of seconds to wait
    ' Source:    Total Visual SourceBook
    ' ====================================================================

    Debug.Print "WaitSeconds"
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure WaitSeconds of Class aegit_expClass", vbCritical, "ERROR"
    Resume PROC_EXIT
End Sub

Private Sub WriteStringToFile(ByVal lngFileNum As Long, ByVal strTheString As String, _
    ByVal strTheAbsoluteFileName As String, Optional ByVal varDebug As Variant)

    'Debug.Print "WriteStringToFile"
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
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure WriteStringToFile of Class aegit_expClass", vbCritical, "ERROR"
        Resume Next
    End If

End Sub