Option Compare Database
Option Explicit

Public Sub Test_aeProjectExport()
    ' Ref: http://www.cpearson.com/excel/vbe.aspx
    ' Ref: http://www.utteraccess.com/forum/lofiversion/index.php/t2019941.html
    ' Ref: http://www.vbaexpress.com/forum/archive/index.php/t-41530.html
    ' Adapted for Access to use late binding and avoid manually setting Applications Extensibility library reference

    '''Call AddLibrary("VBIDE", "{0002E157-0000-0000-C000-000000000046}", 5, 3)

    Dim MyVBAProject As Object
    Set MyVBAProject = VBE.ActiveVBProject
    On Error Resume Next
    aeAddReferenceVBIDE
    'Stop

    Dim VBComp As Object
    Set VBComp = MyVBAProject.VBIDE.VBComponent
    Dim strName As String

    Set MyVBAProject = VBE.ActiveVBProject
    strName = CurrentProject.Path & "\Export"
    Debug.Print "strName = " & strName

    For Each VBComp In MyVBAProject.VBComponents
        ExportVBComponent VBComp, strName, VBComp.Name, True
    Next VBComp

    strName = ""
    Set VBComp = Nothing
    Set MyVBAProject = Nothing

End Sub

Private Sub aeAddReferenceVBIDE()

    Dim chkRef As Variant
    For Each chkRef In Application.VBE.VBProjects.VBE.ActiveVBProject.References              ' ActiveDocument.VBProject.References
        If chkRef.Name = "VBIDE" Then
            Exit Sub
        End If
    Next
    Application.VBE.VBProjects.VBE.ActiveVBProject.References.AddFromFile _
        "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.olb"
End Sub

Private Function AddLibrary(libName As String, GUID As String, major As Long, minor As Long)
    ' AddLibrary: Add a library reference programmatically

    Dim extObj As Object
    Set extObj = GetObject(, "Access.Application")
    Dim vbaProject As Object
    Set vbaProject = extObj.VBE.ActiveVBProject
    Dim chkRef As Object

    ' Check if the library has already been added
    For Each chkRef In vbaProject.References
        If chkRef.Name = libName Then
            GoTo PROC_EXIT
        End If
    Next

    vbaProject.References.AddFromGuid GUID, major, minor

PROC_EXIT:
    Set vbaProject = Nothing

End Function

Public Function IsEditorInSync() As Boolean
    '=======================================================================
    ' IsEditorInSync
    ' This tests if the VBProject selected in the Project window, and
    ' therefore the ActiveVBProject is the same as the VBProject associated
    ' with the ActiveCodePane. If these two VBProjects are the same,
    ' the editor is in sync and the result is True. If these are not the
    ' same project, the editor is out of sync and the result is False.
    '=======================================================================
    On Error GoTo 0
    Dim blnTest As Boolean
    With Application.VBE
        Debug.Print ".ActiveVBProject = " & .ActiveVBProject.Name
        Debug.Print ".ActiveCodePane.CodeModule.Parent.Collection.Parent = " & .ActiveCodePane.CodeModule.Parent.Collection.Parent.Name
        blnTest = .ActiveVBProject Is .ActiveCodePane.CodeModule.Parent.Collection.Parent
    End With
    IsEditorInSync = blnTest
End Function

Public Sub ProjectImport()

    Dim MyVBAObj As Object
    Set MyVBAObj = CreateObject("VBIDE.VBComponent")

    Dim MyVBAProj As Object
    Set MyVBAProj = MyVBAObj.VBProject

    Dim MyVBAComp As Object
    Set MyVBAComp = MyVBAObj

    Dim fName As String
    Dim CompName As String
    Dim s As String
    Dim TempVBComp As Object
    Set TempVBComp = MyVBAObj

    Set MyVBAProj = VBE.ActiveVBProject

    Const vbext_ct_Document = 100

    For Each MyVBAComp In MyVBAProj.VBComponents
        CompName = MyVBAComp.Name
        fName = CurrentProject.Path & "\Export\" & CompName & GetFileExtension(MyVBAComp)
        If CompName <> "basVersioning" Then
            If MyVBAComp.Type = vbext_ct_Document Then
                ' MyVBAComp is destination module
                Set TempVBComp = MyVBAProj.VBComponents.Import(fName)
                ' TempVBComp is source module
                With MyVBAComp.CodeModule
                    .DeleteLines 1, .CountOfLines
                    s = TempVBComp.CodeModule.Lines(1, TempVBComp.CodeModule.CountOfLines)
                    .InsertLines 1, s
                End With
                MyVBAProj.VBComponents.Remove TempVBComp
            Else
                With MyVBAProj.VBComponents
                    .Remove MyVBAComp
                    .Import FileName:=fName
                End With
            End If
        End If
    Next MyVBAComp

End Sub

Public Function ExportVBComponent(VBComp As Object, _
    FolderName As String, _
    Optional FileName As String, _
    Optional OverwriteExisting As Boolean = True) As Boolean
    ' Export the code module of a VBComponent to a text file.
    ' If FileName is missing, the code will be exported to
    ' a file with the same name as the VBComponent followed by the
    ' appropriate extension.

    Dim Extension As String
    Dim fName As String
    Extension = GetFileExtension(VBComp:=VBComp)
    If Trim(FileName) = vbNullString Then
        fName = VBComp.Name & Extension
    Else
        fName = FileName
        If InStr(1, fName, ".", vbBinaryCompare) = 0 Then
            fName = fName & Extension
        End If
    End If

    If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
        fName = FolderName & fName
    Else
        fName = FolderName & "\" & fName
    End If

    If Dir(fName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
        If OverwriteExisting = True Then
            Kill fName
        Else
            ExportVBComponent = False
            Exit Function
        End If
    End If

    VBComp.Export FileName:=fName
    ExportVBComponent = True

End Function
    
Public Function GetFileExtension(VBComp As Object) As String
    ' Return the file extension based on the Type of VBComponent
    ' Ref: https://msdn.microsoft.com/en-us/library/office/gg264162.aspx

    Const vbext_ct_StdModule = 1
    Const vbext_ct_ClassModule = 2
    Const vbext_ct_MSForm = 3
    Const vbext_ct_Document = 100

    Select Case VBComp.Type
        Case vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case vbext_ct_ClassModule
            GetFileExtension = ".cls"
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case vbext_ct_Document
            GetFileExtension = ".cls"
        Case Else
            GetFileExtension = ".bas"
    End Select

End Function