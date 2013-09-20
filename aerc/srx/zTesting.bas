Attribute VB_Name = "zTesting"
Option Explicit

Public Const FOLDER_WITH_VBA_PROJECT_FILES As String = ""   ' "C:\ae\aejqm\aexlpm\srx"

Public Sub ListExcelProperties()
' Ref: http://stackoverflow.com/questions/17406585/vba-set-custom-document-property
' Workbook.BuiltinDocumentProperties Property
' Ref: http://msdn.microsoft.com/en-us/library/bb220896.aspx

    Dim wb As Workbook
    Dim docProp As DocumentProperty
    Dim propExists As Boolean

    Set wb = Application.ThisWorkbook

    ' Ref: http://stackoverflow.com/questions/16642362/how-to-get-the-following-code-to-continue-on-error
    For Each docProp In wb.CustomDocumentProperties
        On Error Resume Next
        Debug.Print docProp.Name & ": " & docProp.Value
        On Error GoTo 0
    Next

    For Each docProp In ActiveWorkbook.BuiltinDocumentProperties
        On Error Resume Next
        Debug.Print docProp.Name & ": " & docProp.Value
        On Error GoTo 0
    Next

    Set wb = Nothing

End Sub
