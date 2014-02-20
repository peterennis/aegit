Option Compare Database
Option Explicit

Private mblnSubFolder As Boolean
Private mintSubFolderLevel As Integer

Public Sub TestListFilesRecursively()
    Const TEST_FILE_PATH As String = "C:\TEMP\"
    Dim strPath As String
    strPath = TEST_FILE_PATH
    mblnSubFolder = False
    mintSubFolderLevel = 1
    ListFilesRecursively strPath, "FoldersOnly"
End Sub

Private Sub ListFilesRecursively(strRootPathName As String, Optional varFoldersOnly As Variant)
' Ref: http://blogs.msdn.com/b/gstemp/archive/2004/08/10/212113.aspx
'====================================================================
' Purpose:  List Files Recursively
' Author:   Peter Ennis
' Date:     February 10, 2011
' Comment:  Fix to work in VBA. Based on MSDN sample for WScript
' Requires: Reference to Microsoft Scripting Runtime
'====================================================================

    Dim strFolder As String
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim colFiles As Object

    strFolder = strRootPathName

    ' Create needed objects
    Dim wsh As Object  ' As Object if late-bound
    Set wsh = CreateObject("WScript.Shell")
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(strFolder)

    'Debug.Print "objFolder.Path = " & objFolder.Path
    Debug.Print ">" & objFolder.Path

    Set colFiles = objFolder.Files

    If IsMissing(varFoldersOnly) Then
        For Each objFile In colFiles
            Debug.Print "objFile.Path = " & objFile.Path
        Next
    End If

    If IsMissing(varFoldersOnly) Then
        ShowSubFolders objFolder
    Else
        Debug.Print "Root Level=" & mintSubFolderLevel
        ShowSubFolders objFolder, varFoldersOnly
    End If
    Debug.Print "DONE !!!"

    Set wsh = Nothing
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set colFiles = Nothing

End Sub
 
Private Sub ShowSubFolders(objFolder As Object, Optional varFoldersOnly As Variant)
' Ref: http://blogs.msdn.com/b/gstemp/archive/2004/08/10/212113.aspx

'    Static i As Integer
    Dim objFile As Object
    Dim objSubFolder As Object
    Dim colFiles As Object
    
    Dim colFolders As Object
    Set colFolders = objFolder.SubFolders

    Dim wsh As Object  ' As Object if late-bound
    Set wsh = CreateObject("WScript.Shell")

'    i = i + 1
    Debug.Print mintSubFolderLevel, mblnSubFolder
    For Each objSubFolder In colFolders

        'Debug.Print "objSubFolder.Path = " & objSubFolder.Path
        Debug.Print ">>" & objSubFolder.Path
        Set colFiles = objSubFolder.Files

        If IsMissing(varFoldersOnly) Then
            For Each objFile In colFiles
                Debug.Print "objFile.Path = " & objFile.Path
            Next
        End If
        
        If IsMissing(varFoldersOnly) Then
            ShowSubFolders objSubFolder
        Else
            mintSubFolderLevel = mintSubFolderLevel + 1
            Debug.Print "Sub Level=" & mintSubFolderLevel
            ShowSubFolders objSubFolder, varFoldersOnly
        End If
        mintSubFolderLevel = mintSubFolderLevel - 1
    Next

    Set wsh = Nothing
    Set colFolders = Nothing

End Sub