Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' Ref: http://www.di-mgt.com.au/cl_Simple.html
'====================================================================
' Author:   Peter F. Ennis
' Date:     February 24, 2011
' Comment:  Create class for revision control
'====================================================================

Private Const VERSION As String = "0.1.5"
Private Const VERSION_DATE As String = "Nov 27, 2012"
Private Const THE_DRIVE As String = "C"
'
'20121127 v015  Update version, export using OASIS and commit to github
'               Reverse order of version comments so newest is at the top
'               Skip ~TMP* names for scripts (macros)
'20110303 v014  Make class PublicNotCreatable, project name aegitClassProvider
'               http://support.microsoft.com/kb/555159
'20110303 v013  Initialize class using Private Type
'20110303 v012  Fix bug in skip export of all zzz objects, must use doc.Name
'20110303 v011  Skip export of all zzz objects, create module basTESTaegitClass
'20110303 v010  Add Option blnDebug to ReadDocDatabase property
'20110302 v010  Delete basRevisionControl
'20110302 v009  Skip export of ~TMP queries, debug message output singular and plural
'20110302 v008  Move other finctions from basRevisionControl to asgitClass
'20110302 v007  Add private function aeDocumentTheDatabase from DocumentTheDatabase
'               Test with updated aegitClassTest
'20110226 v006  TEST_FOLDER=>THE_FOLDER, TEST_DRIVE=>THE_DRIVE, BuildTestDirectory=>BuildTheDirectory
'               Objects have obj prefix, use For Each qdf, output "Macros EXPORTED" (not Scripts)
'20110222 v004  Create aegitClass shell and basTestRevisionControl
'               Use ?aegitClassTest of basTestRevisionControl in the immediate window to check basic operation
'

Private Type mySetupType
    SourceFolder As String
    TestFolder As String
End Type

Private aegitType As mySetupType
'

Private Sub Class_Initialize()
' Ref: http://www.cadalyst.com/cad/autocad/programming-with-class-part-1-5050
' Ref: http://www.bigresource.com/Tracker/Track-vb-cyJ1aJEyKj/
' Ref: http://stackoverflow.com/questions/1731052/is-there-a-way-to-overload-the-constructor-initialize-procedure-for-a-class-in
    
'    If IsMissing(varDebug) Then
'        blnDebug = False
'    Else
'        blnDebug = True
'    End If
    aegitType.SourceFolder = "C:\ae\aegit\aerc\src\"
    aegitType.TestFolder = "C:\ae\aegit\aerc\tst\"
'    If blnDebug Then
        Debug.Print "Class_Initialize"
        Debug.Print , "aegitType.SourceFolder=" & aegitType.SourceFolder
        Debug.Print , "aegitType.TestFolder=" & aegitType.TestFolder
'    End If
End Sub

Property Get SourceFolder() As String
    SourceFolder = aegitType.SourceFolder
End Property

Property Get TestFolder() As String
    TestFolder = aegitType.TestFolder
End Property

Property Get DocumentTheDatabase(Optional blnDebug As Variant) As Boolean
    If IsMissing(blnDebug) Then
        DocumentTheDatabase = aeDocumentTheDatabase
    Else
        DocumentTheDatabase = aeDocumentTheDatabase(blnDebug)
    End If
End Property

Property Get Exists(strAccObjType As String, _
                        strAccObjName As String) As Boolean
    Exists = aeExists(strAccObjType, strAccObjName)
End Property

Property Get ReadDocDatabase(Optional blnDebug As Variant) As Boolean
    If IsMissing(blnDebug) Then
        ReadDocDatabase = aeReadDocDatabase
    Else
        ReadDocDatabase = aeReadDocDatabase(blnDebug)
    End If
End Property

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
' Updated:
'
' 20121127: Reverse comment order, newest at top
'           Skip export of ~TMP macros
' 20110303: Add Optional blnDebug parameter
'           Skip export of all zzz objects (using doc.Name)
' 20110302: Skip export of ~TMP queries
'           debug message output singular and plural
' 20110302: Change to aeDocumentTheDatabase for use in aegitClass
' 20110226: Skip export of MSys (hiddem system queries) and
'           ~sq_ (hidden ODBC queries) objects
'           Add count of objects in debug output
' 20110224: Make this a function. Add optional debug flag
' 20110218: Forms->frm, Reports->rpt, Scripts->mac
'           Modules->bas, Queries->qry
'           Error handler
'====================================================================
'

    Dim dbs As DAO.Database
    Dim cnt As DAO.Container
    Dim doc As DAO.Document
    Dim qdf As DAO.QueryDef
    Dim i As Integer

    Dim blnDebug As Boolean

    On Error GoTo aeDocumentTheDatabase_Error

    If IsMissing(varDebug) Then
        blnDebug = False
    Else
        blnDebug = True
    End If
    
    Set dbs = CurrentDb() ' use CurrentDb() to refresh Collections

    '=============
    ' FORMS
    '=============
    i = 0
    Set cnt = dbs.Containers("Forms")
    Debug.Print "FORMS"
    For Each doc In cnt.Documents
        If Not (Left(doc.Name, 3) = "zzz") Then
            i = i + 1
            Application.SaveAsText acForm, doc.Name, aegitType.SourceFolder & doc.Name & ".frm"
        End If
    Next doc
    If blnDebug Then
        If i = 1 Then
            Debug.Print , "1 Form EXPORTED!"
        Else
            Debug.Print , i & " Forms EXPORTED!"
        End If
    End If
    If blnDebug Then
        If cnt.Documents.Count = 1 Then
            Debug.Print , "1 Form EXISTING!"
        Else
            Debug.Print , cnt.Documents.Count & " Forms EXISTING!"
        End If
    End If
    
    '=============
    ' REPORTS
    '=============
    i = 0
    Set cnt = dbs.Containers("Reports")
    Debug.Print "REPORTS"
    For Each doc In cnt.Documents
        If Not (Left(doc.Name, 3) = "zzz") Then
            i = i + 1
            Application.SaveAsText acReport, doc.Name, aegitType.SourceFolder & doc.Name & ".rpt"
        End If
    Next doc
    If blnDebug Then
        If i = 1 Then
            Debug.Print , "1 Report EXPORTED!"
        Else
            Debug.Print , i & " Reports EXPORTED!"
        End If
    End If
    If blnDebug Then
        If cnt.Documents.Count = 1 Then
            Debug.Print , "1 Report EXISTING!"
        Else
            Debug.Print , cnt.Documents.Count & " Reports EXISTING!"
        End If
    End If

    '=============
    ' MACROS
    '=============
    i = 0
    Set cnt = dbs.Containers("Scripts")
    Debug.Print "MACROS"
    For Each doc In cnt.Documents
        Debug.Print , doc.Name
        If Not (Left(doc.Name, 3) = "zzz" Or Left(doc.Name, 4) = "~TMP") Then
            i = i + 1
            Application.SaveAsText acMacro, doc.Name, aegitType.SourceFolder & doc.Name & ".mac"
        End If
    Next doc
    If blnDebug Then
        If i = 1 Then
            Debug.Print , "1 Macro EXPORTED!"
        Else
            Debug.Print , i & " Macros EXPORTED!"
        End If
    End If
    If blnDebug Then
        If cnt.Documents.Count = 1 Then
            Debug.Print , "1 Macro EXISTING!"
        Else
            Debug.Print , cnt.Documents.Count & " Macros EXISTING!"
        End If
    End If

    '=============
    ' MODULES
    '=============
    i = 0
    Debug.Print "MODULES"
    Set cnt = dbs.Containers("Modules")
    For Each doc In cnt.Documents
        Debug.Print , doc.Name
        If Not (Left(doc.Name, 3) = "zzz") Then
            i = i + 1
            Application.SaveAsText acModule, doc.Name, aegitType.SourceFolder & doc.Name & ".bas"
        End If
    Next doc
    If blnDebug Then
        If i = 1 Then
            Debug.Print , "1 Module EXPORTED!"
        Else
            Debug.Print , i & " Modules EXPORTED!"
        End If
    End If
    If blnDebug Then
        If cnt.Documents.Count = 1 Then
            Debug.Print , "1 Module EXISTING!"
        Else
            Debug.Print , cnt.Documents.Count & " Modules EXISTING!"
        End If
    End If

    '=============
    ' QUERIES
    '=============
    i = 0
    Debug.Print "QUERIES"
    For Each qdf In CurrentDb.QueryDefs
        Debug.Print , qdf.Name
        If Not (Left(qdf.Name, 4) = "MSys" Or Left(qdf.Name, 4) = "~sq_" _
                        Or Left(qdf.Name, 4) = "~TMP" _
                        Or Left(qdf.Name, 3) = "zzz") Then
            i = i + 1
            Application.SaveAsText acQuery, qdf.Name, aegitType.SourceFolder & qdf.Name & ".qry"
        End If
    Next qdf
    If blnDebug Then
        If i = 1 Then
            Debug.Print , "1 Query EXPORTED!"
        Else
            Debug.Print , i & " Queries EXPORTED!"
        End If
    End If
    If blnDebug Then
        If CurrentDb.QueryDefs.Count = 1 Then
            Debug.Print , "1 Query EXISTING!"
        Else
            Debug.Print , CurrentDb.QueryDefs.Count & " Queries EXISTING!"
        End If
    End If

    Set doc = Nothing
    Set cnt = Nothing
    Set dbs = Nothing
    Set qdf = Nothing

    On Error GoTo 0
    aeDocumentTheDatabase = True
    Exit Function

aeDocumentTheDatabase_Error:

    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTheDatabase of Class aegitClass"

End Function

Private Function BuildTheDirectory(FSO As Scripting.FileSystemObject, _
                                        Optional blnDebug As Variant) As Boolean
' Ref: http://msdn.microsoft.com/en-us/library/ebkhfaaz(v=vs.85).aspx
'====================================================================
' Author:   Peter F. Ennis
' Date:     February 8, 2011
' Comment:  Add optional debug parameter
' Requires: Reference to Microsoft Scripting Runtime
' 20110302: Add error handler and include in aegitClass
'====================================================================

    Dim objTestFolder As Object
    
    On Error GoTo BuildTheDirectory_Error

    If IsMissing(blnDebug) Then blnDebug = False

    If blnDebug Then Debug.Print ">==> BuildTestDirectory >==>"

    ' Bail out if (a) the drive does not exist, or if (b) the directory already exists.

    If blnDebug Then Debug.Print "THE_DRIVE = " & THE_DRIVE
    If blnDebug Then Debug.Print "FSO.DriveExists(THE_DRIVE) = " & FSO.DriveExists(THE_DRIVE)
    If Not FSO.DriveExists(THE_DRIVE) Then
        Debug.Print "FSO.DriveExists(THE_DRIVE) = FALSE - The drive DOES NOT EXIST !!!"
        BuildTheDirectory = False
        Exit Function
    End If
    If blnDebug Then Debug.Print "The drive EXISTS !!!"

    If blnDebug Then Debug.Print "The test folder is: " & aegitType.TestFolder
    If FSO.FolderExists(aegitType.TestFolder) Then
        If blnDebug Then Debug.Print "FSO.FolderExists(aegitType.TestFolder) = TRUE - The directory EXISTS !!!"
        BuildTheDirectory = False
        Exit Function
    End If
    If blnDebug Then Debug.Print "The test directory does NOT EXIST !!!"

    Set objTestFolder = FSO.CreateFolder(aegitType.TestFolder)
    If blnDebug Then Debug.Print aegitType.TestFolder & " has been CREATED !!!"

    Set objTestFolder = Nothing

    On Error GoTo 0
    BuildTheDirectory = True
    Exit Function

BuildTheDirectory_Error:

    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure BuildTheDirectory of Class Module aegitClass"

End Function

Public Function aeReadDocDatabase(Optional varDebug As Variant) As Boolean
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
' 20110224: Make this a function
' 20110302: Change to aeReadDocDatabase for use in aegitClass
'           Add Skipping: to MsgBox for existing objects
' 20110303: Add Debug.Print output for Skipping: message
'           Output VERSION and VERSION_DATE for debug
'====================================================================
'

    Dim blnDebug As Boolean
    
    On Error GoTo aeReadDocDatabase_Error
    
    If IsMissing(varDebug) Then
        blnDebug = False
    Else
        blnDebug = True
    End If

    Const acQuery = 1

    Dim myFile As Object
    Dim strFileType As String
    Dim strFileBaseName As String
    
    Dim bln As Boolean

    If blnDebug Then Debug.Print ">==> ReadDocDatabase >==>"
    If blnDebug Then Debug.Print "aegit VERSION: " & VERSION
    If blnDebug Then Debug.Print "aegit VERSION_DATE: " & VERSION_DATE
    If blnDebug Then Debug.Print "aegitType.SourceFolder=" & aegitType.SourceFolder
    If blnDebug Then Debug.Print "aegitType.TestFolder=" & aegitType.TestFolder

    '''''''''' Create needed objects
    Dim wsh As Object  ' As Object if late-bound
    Set wsh = CreateObject("WScript.Shell")
        If blnDebug Then Debug.Print "wsh.CurrentDirectory=" & wsh.CurrentDirectory
        ' CurDir Function
        If blnDebug Then Debug.Print "CurDir=" & CurDir
    Dim FSO As Scripting.FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

    If blnDebug Then
        bln = BuildTheDirectory(FSO, blnDebug)
        Debug.Print "BuildTestDirectory(FSO," & blnDebug & ") = " & bln
    Else
        bln = BuildTheDirectory(FSO)
    End If

    Dim objFolder As Object
    Set objFolder = FSO.GetFolder(aegitType.TestFolder)

    For Each myFile In objFolder.Files
        If blnDebug Then Debug.Print "myFile = " & myFile
        If blnDebug Then Debug.Print "myFile.Name = " & myFile.Name
        strFileBaseName = FSO.GetBaseName(myFile.Name)
        strFileType = FSO.GetExtensionName(myFile.Name)
        If blnDebug Then Debug.Print strFileBaseName & " (" & strFileType & ")"

        If (strFileType = "frm") Then
            If Exists("FORMS", strFileBaseName) Then
                MsgBox "Skipping: FORM " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                If blnDebug Then Debug.Print "Skipping: FORM " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
            Else
                Application.LoadFromText acForm, strFileBaseName, myFile.Path
            End If
        ElseIf (strFileType = "rpt") Then
            If Exists("REPORTS", strFileBaseName) Then
                MsgBox "Skipping: REPORT " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                If blnDebug Then Debug.Print "Skipping: REPORT " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
            Else
                Application.LoadFromText acReport, strFileBaseName, myFile.Path
            End If
        ElseIf (strFileType = "bas") Then
            If Exists("MODULES", strFileBaseName) Then
                MsgBox "Skipping: MODULE " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                If blnDebug Then Debug.Print "Skipping: MODULE " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
            Else
                Application.LoadFromText acModule, strFileBaseName, myFile.Path
            End If
        ElseIf (strFileType = "mac") Then
            If Exists("MACROS", strFileBaseName) Then
                MsgBox "Skipping: MACRO " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                If blnDebug Then Debug.Print "Skipping: MACRO " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
            Else
                Application.LoadFromText acMacro, strFileBaseName, myFile.Path
            End If
        ElseIf (strFileType = "qry") Then
            If Exists("QUERIES", strFileBaseName) Then
                MsgBox "Skipping: QUERY " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                If blnDebug Then Debug.Print "Skipping: QUERY " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
            Else
                Application.LoadFromText acQuery, strFileBaseName, myFile.Path
            End If
        End If
    Next

    Debug.Print "DONE !!!"

    On Error GoTo 0
    aeReadDocDatabase = True
    Exit Function

aeReadDocDatabase_Error:

    MsgBox "Erl=" & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadDocDatabase of Class aegitClass"

End Function

Private Function aeExists(strAccObjType As String, _
                        strAccObjName As String) As Boolean
'
'====================================================================
' Author:     Peter F. Ennis
' Date:       February 18, 2011
' Comment:    Return True if the object exists
' Parameters:
'             strAccObjType: "Tables", "Queries", "Forms",
'                            "Reports", "Macros", "Modules"
'             strAccObjName: The name of the object
' 20110302:   Make aeExists private in aegitClass
'====================================================================

    Dim objType As Object
    Dim obj As Variant
    
    On Error GoTo aeExists_Error

    aeExists = False

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
    End Select

    For Each obj In objType
        If obj.Name = strAccObjName Then
            aeExists = True
            Exit For ' Found it!
        End If
    Next

    On Error GoTo 0
    Set obj = Nothing
    Exit Function

aeExists_Error:

    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class aegitClass"

End Function