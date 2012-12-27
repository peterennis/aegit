Option Compare Database
Option Explicit

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
' Usage:
' Run in immediate window:                  aegitClassTest
' Show debug output in immediate window:    aegitClassTest("debug")

    Dim oDbObjects As aegitClass
    Set oDbObjects = New aegitClass

    Dim bln1 As Boolean
    Dim bln2 As Boolean
    Dim bln3 As Boolean
    Dim bln4 As Boolean
    Dim bln5 As Boolean
    Dim bln6 As Boolean

    'oDbObjects.SourceFolder = "C:\Users\Peter\Documents\GitHub\aegit\aerc\src\"
    'oDbObjects.SourceFolder = "C:\TEMP\aegit\"

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
        bln3 = oDbObjects.ReadDocDatabase()
    Else
        Debug.Print , "Debugit IS NOT missing so blnDebug is set to True"
        bln3 = oDbObjects.ReadDocDatabase("WithDebugging")
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
    
    Debug.Print "Test 1: DocumentTheDatabase"
    Debug.Print "Test 2: Exists"
    Debug.Print "Test 3: ReadDocDatabase"
    Debug.Print "Test 4: GetReferences"
    Debug.Print "Test 5: DocumentTables"
    Debug.Print "Test 6: DocumentRelations"
    Debug.Print "Test 1", "Test 2", "Test 3", "Test 4", "Test 5", "Test 6"
    Debug.Print PassFail(bln1), PassFail(bln2), PassFail(bln3), PassFail(bln4), PassFail(bln5), PassFail(bln6)

End Function

Public Sub ContainerObjectX()
' Ref: http://msdn.microsoft.com/en-us/library/office/bb177484(v=office.12).aspx

   Dim dbs As Database
   Dim ctrLoop As Container
   Dim prpLoop As Property

   Set dbs = CurrentDb

   With dbs

      ' Enumerate Containers collection.
      For Each ctrLoop In .Containers
         Debug.Print "Properties of " & ctrLoop.Name _
            & " container"

         ' Enumerate Properties collection of each Container object.
         For Each prpLoop In ctrLoop.Properties
            Debug.Print "  " & prpLoop.Name _
               & " = "; prpLoop
         Next prpLoop

      Next ctrLoop

      .Close
   End With

End Sub

Public Function ListContainers()
' Ref: http://www.susandoreydesigns.com/software/AccessVBATechniques.pdf
    Dim conItem As Container
    Dim strName As String
    Dim strOwner As String
    Dim strText As String
    For Each conItem In DBEngine.Workspaces(0).Databases(0).Containers
        strName = conItem.Name
        strOwner = conItem.Owner
        strText = "Container name: " & strName & ", Owner: " & strOwner
        Debug.Print strText
    Next conItem
    ListContainers = True
End Function