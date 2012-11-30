Option Compare Database
Option Explicit

' RESEARCH:
' Ref: http://stackoverflow.com/questions/47400/best-way-to-test-a-ms-access-application#70572
' Ref: http://sourceforge.net/projects/vb-lite-unit/
' Using VB 2008 to access a Microsoft Access .accdb database
' Ref: http://boards.straightdope.com/sdmb/showthread.php?t=514884

Public Function New_aegitClass() As aegitClass
' Ref: http://support.microsoft.com/kb/555159#top
'===========================================================================================
' Author:   Peter F. Ennis
' Date:     March 3, 2011
' Comment:  Instantiation of PublicNotCreatable aegitClass
' Updated:  November 27, 2012
'           Added project to github and fixed aegitClassTest configuration for the new setup
'===========================================================================================
    Set New_aegitClass = New aegitClass
End Function

Public Sub aegitClass_EarlyBinding()
'    Dim my_aegitSetup As aegitClassProvider.aegitClass
'    Set my_aegitSetup = aegitClassProvider.
'    anEmployee.Name = "Tushar Mehta"
'    MsgBox anEmployee.Name
End Sub
    
Public Sub aegitClass_LateBinding()
'    Dim anEmployee As Object
'    Set anEmployee = Application.Run("'g:\temp\class provider.xls'!new_clsEmployee")
'    anEmployee.Name = "Tushar Mehta"
'    MsgBox anEmployee.Name
End Sub

Private Function PassFail(bln As Boolean) As String
    If bln Then
        PassFail = "Pass"
    Else
        PassFail = "Fail"
    End If
End Function

Public Function aegitClassTest(Optional Debugit As Variant)
' Usage:
' Run in immediate window:                  aegitClassTest
' Show debug output in immediate window:    aegitClassTest("debug")

    Dim oDbObjects As aegitClass
    Set oDbObjects = New aegitClass

    Dim bln1 As Boolean
    Dim bln2 As Boolean
    Dim bln3 As Boolean
    Dim bln4 As Boolean

    'oDbObjects.SourceFolder = "C:\Users\Peter\Documents\GitHub\aegit\aerc\src\"
    'oDbObjects.SourceFolder = "C:\TEMP\aegit\"

    'MsgBox IsMissing(Debugit)

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
        bln2 = oDbObjects.Exists("Modules", "basRevisionControl")
    Else
        Debug.Print , "Debugit IS NOT missing so blnDebug is set to True"
        bln2 = oDbObjects.Exists("Modules", "aegitClass", "WithDebugging")
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
    
    Debug.Print "Test 1: DocumentTheDatabase"
    Debug.Print "Test 2: Exists"
    Debug.Print "Test 3: ReadDocDatabase"
    Debug.Print "Test 4: GetReferences"
    Debug.Print "Test 1", "Test 2", "Test 3", "Test 4"
    Debug.Print PassFail(bln1), PassFail(bln2), PassFail(bln3), PassFail(bln4)

    'Stop
End Function