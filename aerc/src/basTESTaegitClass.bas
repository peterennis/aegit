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

Public Function aegitClassTest()

    Dim objDbObjects As aegitClass
    Set objDbObjects = New aegitClass
    
    Dim bln1 As Boolean
    Dim bln2 As Boolean
    Dim bln3 As Boolean

    Debug.Print "objDbObjects.SourceFolder=" & objDbObjects.SourceFolder
    Debug.Print "objDbObjects.TestFolder=" & objDbObjects.TestFolder
    bln1 = objDbObjects.DocumentTheDatabase(True)
    bln2 = objDbObjects.Exists("Modules", "basRevisionControl")
    bln3 = objDbObjects.ReadDocDatabase(True)

    Debug.Print bln1, bln2, bln3

    'Stop
End Function