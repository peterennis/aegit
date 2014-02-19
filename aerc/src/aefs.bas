Option Compare Database
Option Explicit

Private Const TEST_FILE_PATH As String = "C:\TEMP\"
'

Public Sub ListFilesRecursively()
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

   strFolder = TEST_FILE_PATH

   ' Create needed objects
   Dim wsh As Object  ' As Object if late-bound
   Set wsh = CreateObject("WScript.Shell")
    
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Set objFolder = objFSO.GetFolder(strFolder)

   Debug.Print "objFolder.Path = " & objFolder.Path

   Set colFiles = objFolder.Files

   For Each objFile In colFiles
       Debug.Print "objFile.Path = " & objFile.Path
   Next

   ShowSubFolders objFolder
   Debug.Print "DONE !!!"

   Set wsh = Nothing
   Set objFSO = Nothing
   Set objFolder = Nothing
   Set colFiles = Nothing

End Sub
 
Private Sub ShowSubFolders(objFolder)
' Ref: http://blogs.msdn.com/b/gstemp/archive/2004/08/10/212113.aspx

   Dim objFile As Object
   Dim objSubFolder As Object
   Dim colFolders As Object
   Dim colFiles As Object
   Dim wsh As Object  ' As Object if late-bound
   Set wsh = CreateObject("WScript.Shell")

   Set colFolders = objFolder.SubFolders
    
   For Each objSubFolder In colFolders
  
       Debug.Print "objSubFolder.Path = " & objSubFolder.Path
       Set colFiles = objSubFolder.Files
  
       For Each objFile In colFiles
           Debug.Print "objFile.Path = " & objFile.Path
       Next

       ShowSubFolders objSubFolder
   Next

   Set wsh = Nothing
   Set colFolders = Nothing

End Sub