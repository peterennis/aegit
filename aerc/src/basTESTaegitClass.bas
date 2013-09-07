Option Compare Database
Option Explicit

' Default Usage:
' The following folders are used if no custom configuration is provided:
' aegitType.SourceFolder = "C:\ae\aegit\aerc\src\"
' aegitType.ImportFolder = "C:\ae\aegit\aerc\imp\"
' Run in immediate window:                  aegitClassTest
' Show debug output in immediate window:    aegitClassTest("debug")
'
' Custom Usage:
' Public Const THE_SOURCE_FOLDER = "Z:\The\Source\Folder\src.MYPROJECT\"
' For custom configuration of the output source folder in aegitClassTest use:
' oDbObjects.SourceFolder = THE_SOURCE_FOLDER
' Run in immediate window: MYPROJECT_TEST
'
Public Function MYPROJECT_TEST()
    'aegitClassTest
    'aegitClassTest ("debug")
End Function

Public Sub ListOrCloseAllOpenQueries(Optional strCloseAll As Variant)
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

Private Function PassFail(bln As Boolean) As String
    If bln Then
        PassFail = "Pass"
    Else
        PassFail = "Fail"
    End If
End Function

Public Function aegitClassTest(Optional Debugit As Variant) As Boolean

    Dim oDbObjects As aegitClass
    Set oDbObjects = New aegitClass

    Dim bln1 As Boolean
    Dim bln2 As Boolean
    Dim bln3 As Boolean
    Dim bln4 As Boolean
    Dim bln5 As Boolean
    Dim bln6 As Boolean
    Dim bln7 As Boolean

    oDbObjects.SourceFolder = THE_SOURCE_FOLDER

Test1:
    '=============
    ' TEST 1
    '=============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "1. aegitClassTest => DocumentTheDatabase"
    Debug.Print "aegitClassTest"
    If IsMissing(Debugit) Then
        Debug.Print , "Debugit IS missing so no parameter is passed to DocumentTheDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        bln1 = oDbObjects.DocumentTheDatabase()
    Else
        Debug.Print , "Debugit IS NOT missing so blnDebug is set to True"
        bln1 = oDbObjects.DocumentTheDatabase("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

Test2:
    '=============
    ' TEST 2
    '=============
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "2. aegitClassTest => Exists"
    Debug.Print "aegitClassTest"
    If IsMissing(Debugit) Then
        Debug.Print , "Debugit IS missing so no parameter is passed to Exists"
        Debug.Print , "DEBUGGING IS OFF"
        bln2 = oDbObjects.Exists("Modules", "zzzaegitClass")
    Else
        Debug.Print , "Debugit IS NOT missing so blnDebug is set to True"
        bln2 = oDbObjects.Exists("Modules", "zzzaegitClass", "WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

Test3:
    '=============
    ' TEST 3
    '=============
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "3. aegitClassTest => ReadDocDatabase"
    Debug.Print "aegitClassTest"
    If IsMissing(Debugit) Then
        Debug.Print , "Debugit IS missing so no parameter is passed to ReadDocDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        bln3 = oDbObjects.ReadDocDatabase(False)
    Else
        Debug.Print , "Debugit IS NOT missing so blnDebug is set to True"
        bln3 = oDbObjects.ReadDocDatabase(False, "WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

Test4:
    '=============
    ' TEST 4
    '=============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "4. aegitClassTest => GetReferences"
    Debug.Print "aegitClassTest"
    If IsMissing(Debugit) Then
        Debug.Print , "Debugit IS missing so no parameter is passed to GetReferences"
        Debug.Print , "DEBUGGING IS OFF"
        bln4 = oDbObjects.GetReferences()
    Else
        Debug.Print , "Debugit IS NOT missing so blnDebug is set to True"
        bln4 = oDbObjects.GetReferences("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print
    
Test5:
    '=============
    ' TEST 5
    '=============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "5. aegitClassTest => DocumentTables"
    Debug.Print "aegitClassTest"
    If IsMissing(Debugit) Then
        Debug.Print , "Debugit IS missing so no parameter is passed to DocumentTables"
        Debug.Print , "DEBUGGING IS OFF"
        bln5 = oDbObjects.DocumentTables()
    Else
        Debug.Print , "Debugit IS NOT missing so blnDebug is set to True"
        bln5 = oDbObjects.DocumentTables("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print
    
Test6:
    '=============
    ' TEST 6
    '=============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "6. aegitClassTest => DocumentRelations"
    Debug.Print "aegitClassTest"
    If IsMissing(Debugit) Then
        Debug.Print , "Debugit IS missing so no parameter is passed to DocumentRelations"
        Debug.Print , "DEBUGGING IS OFF"
        bln6 = oDbObjects.DocumentRelations()
    Else
        Debug.Print , "Debugit IS NOT missing so blnDebug is set to True"
        bln6 = oDbObjects.DocumentRelations("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print
    
    ' CompactAndRepair
    ' When run this will close the database so obviously no results will be seen.
    ' Use True/False parameter to run it or not.

Test7:
    '=============
    ' TEST 7
    '=============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "7. aegitClassTest => CompactAndRepair"
    Debug.Print "aegitClassTest"

        bln7 = oDbObjects.CompactAndRepair("Do Not Run")

    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

RESULTS:
    Debug.Print "Test 1: DocumentTheDatabase"
    Debug.Print "Test 2: Exists"
    Debug.Print "Test 3: ReadDocDatabase"
    Debug.Print "Test 4: GetReferences"
    Debug.Print "Test 5: DocumentTables"
    Debug.Print "Test 6: DocumentRelations"
    Debug.Print "Test 7: CompactAndRepair"
    Debug.Print
    Debug.Print "Test 1", "Test 2", "Test 3", "Test 4", "Test 5", "Test 6", "Test 7"
    Debug.Print PassFail(bln1), PassFail(bln2), PassFail(bln3), PassFail(bln4), PassFail(bln5), PassFail(bln6), PassFail(bln7)

End Function

Public Function aegitClassImportTest(Optional Debugit As Variant) As Boolean
' Usage:
' Run in immediate window:                  aegitClassImportTest

    Dim oDbObjects As aegitClass
    Set oDbObjects = New aegitClass

    Dim bln1 As Boolean

ImportTest1:
    '==============
    ' IMPORT TEST 1
    '==============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "1. aegitClassImportTest => ReadDocDatabase"
    Debug.Print "aegitClassImportTest"

    oDbObjects.UseImportFolder = True

    If IsMissing(Debugit) Then
        Debug.Print , "Debugit IS missing so no parameter is passed to ReadDocDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        bln1 = oDbObjects.ReadDocDatabase(True)
    Else
        Debug.Print , "Debugit IS NOT missing so blnDebug is set to True"
        bln1 = oDbObjects.ReadDocDatabase(True, "WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

End Function

Public Sub TestHideQueryDef()
' Ref: http://social.msdn.microsoft.com/Forums/office/en-US/25d9dafd-b446-40ba-8dbd-a0efa983f2ff/how-to-programatically-hide-a-querydef

    ' Query1 returns the list of all queries
    ' Ref: http://stackoverflow.com/questions/10882317/get-list-of-queries-in-project-ms-access
    Const strQueryName As String = "Query1"

    Dim fIsHidden As Boolean

    ' To determine if the query is hidden:
    fIsHidden = GetHiddenAttribute(acQuery, strQueryName)
    Debug.Print strQueryName, fIsHidden

    ' To show the query:
    SetHiddenAttribute acQuery, strQueryName, False

    ' To hide the query:
    SetHiddenAttribute acQuery, strQueryName, True

End Sub

Public Function IsQryHidden(strQueryName As String) As Boolean
    IsQryHidden = GetHiddenAttribute(acQuery, strQueryName)
End Function

Public Sub ListOfAllQueries()
' Ref: http://www.pcreview.co.uk/forums/runtime-error-7874-a-t2922352.html

    Const strSQL As String = "SELECT m.Name " & vbCrLf & _
                                "FROM MSysObjects AS m " & vbCrLf & _
                                "WHERE (((m.Name) Not ALike ""~%"") AND ((m.Type)=5)) " & vbCrLf & _
                                "ORDER BY m.Name;"
    ' NOTE: Use zzz* for the query name so that it will be ignored by aegit code export
    Const strTempQuery As String = "zzz___MyTempQuery___"

    Debug.Print strSQL

    Dim qdfCurr As DAO.QueryDef

    ' Create the temp query that will have the query names
    On Error Resume Next
    Set qdfCurr = CurrentDb.QueryDefs(strTempQuery)
    If Err.Number = 3265 Then ' 3265 is "Item not found in this collection."
        Set qdfCurr = CurrentDb.CreateQueryDef(strTempQuery)
    End If
    qdfCurr.SQL = strSQL
    'Debug.Print """" & strTempQuery & """"
    DoCmd.OpenQuery strTempQuery
    'DoCmd.Close acQuery, strTempQuery
    Debug.Print "The number of queries in the database is: " & DCount("Name", strTempQuery)

End Sub

Public Sub MakeTableWithListOfAllQueries()

    Const strTempTable As String = "zzzTmpTblQueries"
    ' NOTE: Use zzz* for the table name so that it will be ignored by aegit code export
    Const strSQL As String = "SELECT m.Name, 0 AS Hidden INTO " & strTempTable & " " & vbCrLf & _
                                "FROM MSysObjects AS m " & vbCrLf & _
                                "WHERE (((m.Name) Not ALike ""~%"") AND ((m.Type)=5)) " & vbCrLf & _
                                "ORDER BY m.Name;"

    ' RunSQL works for Action queries
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True
    Debug.Print "The number of queries in the database is: " & DCount("Name", strTempTable)

End Sub

Public Sub MakeTableWithListOfAllHiddenQueries()

    Const strTempTable As String = "zzzTmpTblQueries"
    ' NOTE: Use zzz* for the table name so that it will be ignored by aegit code export
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
    DoCmd.SetWarnings True
    Debug.Print "The number of hidden queries in the database is: " & DCount("Name", strTempTable)

End Sub