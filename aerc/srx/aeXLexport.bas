Attribute VB_Name = "aeXLexport"
Option Explicit

' Ref: http://www.rondebruin.nl/win/s9/win002.htm
' Requires reference to the VBA extensibility library. Click on Tools-References in the VBE, and
' scroll down and tick the entry for Microsoft Visual Basic for Applications Extensibility 5.3
' Scripting.FileSystemObject requires reference to Microsoft Scripting Runtime

Private Const FOLDER_WITH_VBA_PROJECT_FILES As String = "C:\ae\aegit\aerc\srx"
'

Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim strSourceWorkbook As String
    Dim strExportPath As String
    Dim strFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If

    'Debug.Print FolderWithVBAProjectFiles & "\*.*"
    'Stop

    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    strSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(strSourceWorkbook)

    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If

    strExportPath = FolderWithVBAProjectFiles & "\"

    For Each cmpComponent In wkbSource.VBProject.VBComponents

        bExport = True
        strFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                strFileName = strFileName & ".cls"
            Case vbext_ct_MSForm
                strFileName = strFileName & ".frm"
            Case vbext_ct_StdModule
                strFileName = strFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select

        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export strExportPath & strFileName

        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent

        End If

    Next cmpComponent

    MsgBox "Export is ready"
End Sub

Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim strTargetWorkbook As String
    Dim strImportPath As String
    Dim strFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    strTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(strTargetWorkbook)

    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    strImportPath = FolderWithVBAProjectFiles & "\"

    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(strImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents

    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(strImportPath).Files

        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
    
    Next objFile

    MsgBox "Import is ready"
End Sub

Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If IsNull(FOLDER_WITH_VBA_PROJECT_FILES) Then

        If Right(SpecialPath, 1) <> "\" Then
            SpecialPath = SpecialPath & "\"
        End If

        If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
            On Error Resume Next
            MkDir SpecialPath & "VBAProjectFiles"
            On Error GoTo 0
        End If

        If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
            FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
        Else
            FolderWithVBAProjectFiles = "Error"
        End If

    Else

        FolderWithVBAProjectFiles = FOLDER_WITH_VBA_PROJECT_FILES

    End If

End Function

Function DeleteVBAModulesAndUserForms()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = ActiveWorkbook.VBProject

    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            'Thisworkbook or worksheet module
            'We do nothing
        Else
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
End Function

