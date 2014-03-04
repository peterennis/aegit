Option Compare Database
Option Explicit

Public Const Desktop = &H10&
Public Const MyDocuments = &H5&
'

Public Function SpFolder(SpName)

    Dim objShell As Object
    Set objShell = CreateObject("Shell.Application")
    Dim objFolder As Object
    Set objFolder = objShell.Namespace(SpName)
    Dim objFolderItem As Object
    Set objFolderItem = objFolder.Self

    SpFolder = objFolderItem.Path

End Function
   
Public Sub ExportAllModulesToFile()
' Ref: http://wiki.lessthandot.com/index.php/Code_and_Code_Windows
' Ref: http://stackoverflow.com/questions/2794480/exporting-code-from-microsoft-access
' The reference for the FileSystemObject Object is Windows Script Host Object Model
' but it not necessary to add the reference for this procedure.

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fil As Object
    Dim strMod As String
    Dim mdl As Object
    Dim i As Integer
    Dim strTxtFile As String


    ' Set up the file
    Debug.Print "CurrentProject.Name = " & CurrentProject.Name
    strTxtFile = SpFolder(Desktop) & "\" & Replace(CurrentProject.Name, ".", "_") & ".txt"
    Debug.Print "strTxtFile = " & strTxtFile
    Set fil = fso.CreateTextFile(SpFolder(Desktop) & "\" _
            & Replace(CurrentProject.Name, ".", " ") & ".txt")

    ' For each component in the project ...
    For Each mdl In VBE.ActiveVBProject.VBComponents
        ' using the count of lines ...
        If Left(mdl.Name, 3) <> "zzz" Then
            Debug.Print mdl.Name
            i = VBE.ActiveVBProject.VBComponents(mdl.Name).CodeModule.CountOfLines
            ' put the code in a string ...
            If i > 0 Then
                strMod = VBE.ActiveVBProject.VBComponents(mdl.Name).CodeModule.Lines(1, i)
            End If
            ' and then write it to a file, first marking the start with
            ' some equal signs and the component name.
            fil.WriteLine String(15, "=") & vbCrLf & mdl.Name _
                & vbCrLf & String(15, "=") & vbCrLf & strMod
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

    Dim VBP As Object               'VBIDE.VBProject
    Dim VBComp As Object            'VBIDE.VBComponent
    Dim LineCount As Long

    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=245480
    Const vbext_pp_locked = 1

    Set VBP = Access.Application.VBE.ActiveVBProject

    If VBP.Protection = vbext_pp_locked Then
        CodeLinesInProjectCount = -1
        Exit Function
    End If

    For Each VBComp In VBP.VBComponents
        If Left(VBComp.Name, 3) <> "zzz" Then
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

Public Sub IsAppOpen(strAppName As String)
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
 
    Dim qry As DAO.QueryDef
    Dim cnt As DAO.Container
 
    ' Delete all TEMP queries ...
    For Each qry In CurrentDb.QueryDefs
        If Left(qry.Name, 1) = "~" Then
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