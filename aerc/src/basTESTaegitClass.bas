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

Public Function aegitClassTest(Optional Debugit As Variant)

    Dim oDbObjects As aegitClass
    Set oDbObjects = New aegitClass
    
    Dim bln1 As Boolean
    Dim bln2 As Boolean
    Dim bln3 As Boolean
    Dim bln4 As Boolean

    'oDbObjects.SourceFolder = "C:\Users\Peter\Documents\GitHub\aegit\aerc\src\"
    'oDbObjects.SourceFolder = "C:\TEMP\aegit\"

    'MsgBox IsMissing(Debugit)
    
    '=============
    ' TEST 1
    '=============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "1. aegitClassTest => DocumentTheDatabase"
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
   
    '=============
    ' TEST 2
    '=============
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "2. aegitClassTest => Exists"
    bln2 = oDbObjects.Exists("Modules", "basRevisionControl")
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print
    
    '=============
    ' TEST 3
    '=============
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "3. aegitClassTest => ReadDocDatabase"
    bln3 = oDbObjects.ReadDocDatabase(True)
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

    '=============
    ' TEST 4
    '=============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "4. aegitClassTest => GetReferences"
    If IsMissing(Debugit) Then
        Debug.Print , "Debugit IS missing so no parameter is passed to DocumentTheDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        bln4 = oDbObjects.GetReferences()
    Else
        Debug.Print , "Debugit IS NOT missing so blnDebug is set to True"
        bln4 = oDbObjects.GetReferences("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print
    
    Debug.Print bln1, bln2, bln3, bln4

    'Stop
End Function