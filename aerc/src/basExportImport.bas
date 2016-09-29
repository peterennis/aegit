Option Compare Database
Option Explicit

Public Sub ProjectExport()
    ' Ref: http://www.cpearson.com/excel/vbe.aspx
    ' Ref: http://www.utteraccess.com/forum/lofiversion/index.php/t2019941.html
    ' Adapted to use late binding and avoid needing Applications Extensibility library reference

    Dim MyVBAObj As Object
    Set MyVBAObj = CreateObject("VBIDE.VBComponent")

    Dim strName As String

    Dim MyVBAProj As Object
    Set MyVBAProj = MyVBAObj.ActiveVBProject
    strName = CurrentProject.Path & "\Export"
    Debug.Print "strName = " & strName

    Dim MyVBAComp As Object
    For Each MyVBAComp In MyVBAProj.VBComponents
        aeExportVBComponent MyVBAComp, strName, MyVBAComp.Name, True
    Next MyVBAComp

    strName = ""
    Set MyVBAComp = Nothing
    Set MyVBAProj = Nothing

End Sub

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

    ' Ref: https://msdn.microsoft.com/en-us/library/office/gg264162.aspx
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

Public Function aeExportVBComponent(MyVBAComp As Object, _
    FolderName As String, _
    Optional FileName As String, _
    Optional OverwriteExisting As Boolean = True) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This function exports the code module of a VBComponent to a text
    ' file. If FileName is missing, the code will be exported to
    ' a file with the same name as the VBComponent followed by the
    ' appropriate extension.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Extension As String
    Dim fName As String
    Extension = GetFileExtension(MyVBAComp)
    If Trim(FileName) = vbNullString Then
        fName = MyVBAComp.Name & Extension
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
            aeExportVBComponent = False
            Exit Function
        End If
    End If

    MyVBAComp.Export FileName:=fName
    aeExportVBComponent = True

End Function
    
Public Function GetFileExtension(MyVBComp As Object) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This returns the appropriate file extension based on the Type of
    ' the VBComponent.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Const vbext_ct_ClassModule = 2
    Const vbext_ct_Document = 100
    Const vbext_ct_MSForm = 3
    Const vbext_ct_StdModule = 1

    Select Case MyVBComp.Type
        Case vbext_ct_ClassModule
            GetFileExtension = ".cls"
        Case vbext_ct_Document
            GetFileExtension = ".cls"
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
           Case vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = ".bas"
    End Select

End Function