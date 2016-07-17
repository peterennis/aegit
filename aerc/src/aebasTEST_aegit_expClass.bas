Option Compare Database
Option Explicit

' Default Usage:
' The following folders are used if no custom configuration is provided:
' aegitType.SourceFolder = "C:\ae\aegit\aerc\src\"
' aegitType.aegitXMLfolder = "C:\ae\aegit\aerc\src\xml\"
' Run in immediate window:                  aegit_EXPORT
' Show debug output in immediate window:    aegit_EXPORT varDebug:="varDebug"
'                                           aegit_EXPORT 1
'
' Custom Usage:
' FRONT END SETUP
'Public Const THE_FRONT_END_APP = True
'Public Const THE_SOURCE_FOLDER = ".\src\"                  ' "C:\THE\DATABASE\PATH\src\"
'Public Const THE_XML_FOLDER = ".\src\xml\"                 ' "C:\THE\DATABASE\PATH\src\xml\"
'Public Const THE_XML_DATA_FOLDER = ".\src\xmldata\"        ' "C:\THE\DATABASE\PATH\src\xmldata\"
'Public Const THE_BACK_END_DB1 = "C:\MY\BACKEND\DATA.accdb"
'Public Const THE_BACK_END_SOURCE_FOLDER = "NONE"           ' ".\srcbe\"
'Public Const THE_BACK_END_XML_FOLDER = "NONE"              ' ".\srcbe\xml\"
'Public Const THE_BACK_END_XML_DATA_FOLDER = "NONE"         ' ".\srcbe\xmldata\"

' BACK END SETUP
'Public Const THE_FRONT_END_APP = False
'Public Const THE_SOURCE_FOLDER = "NONE"                     ' ".\src\"
'Public Const THE_XML_FOLDER = "NONE"                        ' ".\src\xml\"
'Public Const THE_XML_DATA_FOLDER = "NONE"                   ' ".\src\xmldata\"
'Public Const THE_BACK_END_DB1 = "NONE"
'Public Const THE_BACK_END_SOURCE_FOLDER = "C:\THE\DATABASE\PATH\srcbe\"             ' ".\srcbe\"
'Public Const THE_BACK_END_XML_FOLDER = "C:\THE\DATABASE\PATH\srcbe\xml\"            ' ".\srcbe\xml\"
'Public Const THE_BACK_END_XML_DATA_FOLDER = "C:\THE\DATABASE\PATH\srcbe\xmldata\"   ' ".\srcbe\xmldata\"
'
' Run in immediate window:                  ALTERNATIVE_EXPORT
' Show debug output in immediate window:    ALTERNATIVE_EXPORT varDebug:="varDebug"
'                                           ALTERNATIVE_EXPORT 1
'
' Sample constants for settings of "TheProjectName"
'Public Const gstrDATE_TheProjectName As String = "January 1, 2000"
'Public Const gstrVERSION_TheProjectName As String = "0.0.0"
'Public Const gstrPROJECT_TheProjectName As String = "TheProjectName"
'Public Const gblnTEST_TheProjectName As Boolean = False

Public Const gstrPROJECT_aegit As String = "aegit export project"
Public Const gstrVERSION_aegit As String = "0.0.0"
Public gvarMyTablesForExportToXML() As Variant
'

Public Sub aegit_EXPORT(Optional ByVal varDebug As Variant)

    On Error GoTo 0

    If Application.VBE.ActiveVBProject.Name <> "aegit" Then
        MsgBox "The is not the aegit project!", vbCritical, "aegit_EXPORT"
        Exit Sub
    End If

    If Not IsMissing(varDebug) Then
        aegitClassTest varDebug:="varDebug", varFrontEndApp:=True
    Else
        aegitClassTest varFrontEndApp:=True
    End If
End Sub

Public Sub ALTERNATIVE_EXPORT(Optional ByVal varDebug As Variant)

    Dim THE_SOURCE_FOLDER As String
    THE_SOURCE_FOLDER = "C:\TEMP\aealt\src\"
    Dim THE_XML_FOLDER As String
    THE_XML_FOLDER = "C:\TEMP\aealt\src\xml\"
    Dim THE_XML_DATA_FOLDER As String
    THE_XML_DATA_FOLDER = "C:\TEMP\aealt\src\xmldata\"

    On Error GoTo PROC_ERR

    If Not IsMissing(varDebug) Then
        aegitClassTest varDebug:="varDebug", varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER, varXmlDataFldr:=THE_XML_DATA_FOLDER, varFrontEndApp:=True
    Else
        aegitClassTest varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER, varXmlDataFldr:=THE_XML_DATA_FOLDER, varFrontEndApp:=True
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ALTERNATIVE_EXPORT"
    Resume Next

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

Private Function IsArrayInitialized(ByVal arr As Variant) As Boolean
    If Not IsArray(arr) Then Err.Raise 13
    On Error Resume Next
    IsArrayInitialized = (LBound(arr) <= UBound(arr))
End Function

Public Sub aegitClassTest(Optional ByVal varDebug As Variant, _
                                Optional ByVal varSrcFldr As Variant, _
                                Optional ByVal varXmlFldr As Variant, _
                                Optional ByVal varXmlDataFldr As Variant, _
                                Optional ByVal varSrcFldrBe As Variant, _
                                Optional ByVal varXmlFldrBe As Variant, _
                                Optional ByVal varXmlDataFldrBe As Variant, _
                                Optional ByVal varBackEndDbOne As Variant, _
                                Optional ByVal varFrontEndApp As Variant)

    On Error GoTo PROC_ERR

    Dim oDbObjects As aegit_expClass
    Set oDbObjects = New aegit_expClass

    Dim blnTestOne As Boolean
    Dim blnTestTwo As Boolean
    Dim blnTestThree As Boolean
    Dim blnTestFour As Boolean
    Dim blnTestFive As Boolean
    Dim blnTestSix As Boolean
    Dim blnTestSeven As Boolean
    Dim blnTestEight As Boolean

    If Not IsMissing(varSrcFldr) Then oDbObjects.SourceFolder = varSrcFldr                  ' THE_SOURCE_FOLDER
    If Not IsMissing(varXmlFldr) Then oDbObjects.XMLFolder = varXmlFldr                     ' THE_XML_FOLDER
    If Not IsMissing(varXmlDataFldr) Then oDbObjects.XMLDataFolder = varXmlDataFldr         ' THE_XML_DATA_FOLDER
    If Not IsMissing(varSrcFldrBe) Then oDbObjects.SourceFolderBe = varSrcFldrBe            ' THE_BACK_END_SOURCE_FOLDER
    If Not IsMissing(varXmlFldrBe) Then oDbObjects.XMLFolderBe = varXmlFldrBe               ' THE_BACK_END_XML_FOLDER
    If Not IsMissing(varXmlDataFldrBe) Then oDbObjects.XMLDataFolderBe = varXmlDataFldrBe   ' THE_XML_DATA_FOLDER
    If Not IsMissing(varBackEndDbOne) Then oDbObjects.BackEndDbOne = varBackEndDbOne              ' THE_BACK_END_DB1
    If Not IsMissing(varFrontEndApp) Then oDbObjects.FrontEndApp = varFrontEndApp           ' THE_FRONT_END_APP
    'MsgBox "varBackEndDbOne = " & varBackEndDbOne, vbInformation, "Procedure aegitClassTest"

    ' Define tables for xml data export
    gvarMyTablesForExportToXML = Array("USysRibbons")
    oDbObjects.TablesExportToXML = gvarMyTablesForExportToXML()

    Debug.Print "aegitClassTest"
    If IsArrayInitialized(gvarMyTablesForExportToXML) Then
        Debug.Print , "UBound(gvarMyTablesForExportToXML) = " & UBound(gvarMyTablesForExportToXML)
        oDbObjects.TablesExportToXML = gvarMyTablesForExportToXML
    Else
        Debug.Print "Array gvarMyTablesForExportToXML is not initialized! There are no tables selected for data export."
    End If

    If Application.VBE.ActiveVBProject.Name = "aegit" Then
        Dim myArray() As Variant
        myArray = Array("aeItems", "aetlkpStates", "USysRibbons")
        oDbObjects.TablesExportToXML = myArray()
        oDbObjects.ExcludeFiles = False
        Debug.Print , "oDbObjects.ExcludeFiles = " & oDbObjects.ExcludeFiles
    End If
    'Stop

TestOne:
    '=============
    ' TEST 1
    '=============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "1. aegitClassTest => DocumentTheDatabase"
    Debug.Print "aegitClassTest"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to DocumentTheDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        blnTestOne = oDbObjects.DocumentTheDatabase()
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to DocumentTheDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
        blnTestOne = oDbObjects.DocumentTheDatabase("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

TestThree:
    '=============
    ' TEST 3
    '=============
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "3. NOT USED"
    Debug.Print "aegitClassTest"

    blnTestThree = False

    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

TestFour:
    '=============
    ' TEST 4
    '=============
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "4. aegitClassTest => GetReferences"
    Debug.Print "aegitClassTest"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to GetReferences"
        Debug.Print , "DEBUGGING IS OFF"
        blnTestFour = oDbObjects.GetReferences()
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to GetReferences"
        Debug.Print , "DEBUGGING TURNED ON"
        blnTestFour = oDbObjects.GetReferences("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print
    
TestFive:
    '=============
    ' TEST 5
    '=============
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "5. aegitClassTest => DocumentTables"
    Debug.Print "aegitClassTest"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to DocumentTables"
        Debug.Print , "DEBUGGING IS OFF"
        blnTestFive = oDbObjects.DocumentTables()
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to DocumentTables"
        Debug.Print , "DEBUGGING TURNED ON"
        blnTestFive = oDbObjects.DocumentTables("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print
    
TestSix:
    '=============
    ' TEST 6
    '=============
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "6. aegitClassTest => DocumentRelations"
    Debug.Print "aegitClassTest"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to DocumentRelations"
        Debug.Print , "DEBUGGING IS OFF"
        blnTestSix = oDbObjects.DocumentRelations()
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to DocumentRelations"
        Debug.Print , "DEBUGGING TURNED ON"
        blnTestSix = oDbObjects.DocumentRelations("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

TestSeven:
    '=============
    ' TEST 7
    '=============
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "7. aegitClassTestXML => DocumentTablesXML"
    Debug.Print "aegitClassTestXML"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to DocumentTheDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        blnTestSeven = oDbObjects.DocumentTablesXML()
    Else
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to DocumentTheDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
        blnTestSeven = oDbObjects.DocumentTablesXML("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

TestEight:
    '=============
    ' TEST 8
    '=============
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "8. NOT USED"
    Debug.Print "aegitClassTest"

    blnTestEight = False

    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

RESULTS:
    Debug.Print "Test 1: DocumentTheDatabase"
    Debug.Print "Test 2: Exists"
    Debug.Print "Test 3: NOT USED"
    Debug.Print "Test 4: GetReferences"
    Debug.Print "Test 5: DocumentTables"
    Debug.Print "Test 6: DocumentRelations"
    Debug.Print "Test 7: DocumentTablesXML"
    Debug.Print "Test 8: NOT USED"
    Debug.Print
    Debug.Print "Test 1", "Test 2", "Test 3", "Test 4", "Test 5", "Test 6", "Test 7", "Test 8"
    Debug.Print PassFail(blnTestOne), PassFail(blnTestTwo), PassFail(blnTestThree, "X"), PassFail(blnTestFour), PassFail(blnTestFive), PassFail(blnTestSix), PassFail(blnTestSeven), PassFail(blnTestEight, "X")

PROC_EXIT:
    Exit Sub

PROC_ERR:
    Select Case Err.Number
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aegitClassTest"
            Stop
    End Select

End Sub