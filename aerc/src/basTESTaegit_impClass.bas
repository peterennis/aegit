Option Compare Database
Option Explicit

' Default Usage:
' The following folders are used if no custom configuration is provided:
' aegitType.SourceFolder = "C:\ae\aegit\aerc\src\"
' aegitType.ImportFolder = "C:\ae\aegit\aerc\imp\"
' Run in immediate window:                  aegitClassTest
' Show debug output in immediate window:    aegitClassTest varDebug:="Debugit"
'
' Custom Usage:
' Public Const THE_SOURCE_FOLDER = "Z:\The\Source\Folder\src.MYPROJECT\"
' For custom configuration of the output source folder in aegitClassTest use:
' oDbObjects.SourceFolder = THE_SOURCE_FOLDER
' oDbObjects.XMLFolder = THE_XML_FOLDER
' Run in immediate window: MYPROJECT_TEST
'

Public Function IMPORT_TEST() As Boolean
    On Error GoTo 0
    'aegitClassImportTest
    aegitClassImportTest varDebug:="Debugit", varImpFldr:="C:\TEMP\imp"
End Function

Public Sub ALTERNATIVE_TEST()

    On Error GoTo 0
    Dim THE_SOURCE_FOLDER As String
    Dim THE_XML_FOLDER As String

    THE_SOURCE_FOLDER = "C:\TEMP\aealt\src\"
    THE_XML_FOLDER = "C:\TEMP\aealt\src\xml\"

    On Error GoTo PROC_ERR
    'aegitClassTest varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER
    aegitClassTest varDebug:="Debugit", varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ALTERNATIVE_TEST"
    Resume Next

End Sub

Private Function ImpPassFail(ByVal bln As Boolean) As String
    On Error GoTo 0
    If bln Then
        ImpPassFail = "Pass"
    Else
        ImpPassFail = "Fail"
    End If
End Function

Public Function aegitClassImportTest(Optional ByVal varDebug As Variant, _
                                Optional ByVal varImpFldr As Variant) As Boolean
' Usage:
' Run in immediate window: aegitClassImportTest

    On Error GoTo 0
    Dim oDbObjects As aegit_impClass
    Set oDbObjects = New aegit_impClass

    Dim bln1 As Boolean

    If Not IsMissing(varImpFldr) Then oDbObjects.ImportFolder = varImpFldr      ' THE_IMPORT_FOLDER

ImportTest1:
    '==============
    ' IMPORT TEST 1
    '==============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "1. aegitClassImportTest => ReadDocDatabase"
    Debug.Print "aegitClassImportTest"

    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to ReadDocDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        bln1 = oDbObjects.ReadDocDatabase(True)
    Else
        Debug.Print , "varDebug IS NOT missing so blnDebug is set to True"
        bln1 = oDbObjects.ReadDocDatabase(True, "WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

RESULTS:
    Debug.Print "Test 1: ReadDocDatabase"
    Debug.Print
    Debug.Print "Test 1", "Test 2", "Test 3", "Test 4", "Test 5", "Test 6", "Test 7", "Test 8"
    Debug.Print ImpPassFail(bln1)

End Function