Attribute VB_Name = "basWorkbook"
Option Explicit

Public Sub GetWorkbookInfo()

    Dim wkb As Workbook
    
    For Each wkb In Workbooks
        'Debug.Print wkb.Name
        If wkb.Name <> ThisWorkbook.Name Then
        End If
    Next wkb

End Sub

Public Function TheComputerName() As String

    Dim blnFlag As Boolean
    Dim strComputerName As String

    blnFlag = GetComputerName(strComputerName)
    TheComputerName = strComputerName

End Function

Public Sub GetTheComputerName()

    Dim blnFlag As Boolean
    Dim strComputerName As String

    blnFlag = GetComputerName(strComputerName)
    MsgBox strComputerName, vbInformation, "Computer Name"

End Sub

Public Sub ActivateSheet(strSheetName As String)
'
' Description: Activate Sheet
' Date: 10/19/2002
' Author: (c) Peter F. Ennis
'
    If SheetExists(strSheetName) Then
        Worksheets(strSheetName).Activate
    Else
        'MsgBox "Sheet '" & strSheetName & "' does not exist!", _
        '    vbCritical, "ActivateSheet"
    End If

End Sub

Public Function SheetExists(strSheetName As String) As Boolean

On Error GoTo Err_SheetExists

    Dim obj As Object
    
    For Each obj In Worksheets
        'Debug.Print Worksheets(obj.Name).Name
        If Worksheets(obj.Name).Name = strSheetName Then
            SheetExists = True
            Exit Function
        End If
    Next obj
    SheetExists = False

Exit_SheetExists:
    Exit Function
    
Err_SheetExists:
    MsgBox Err.Description & vbCrLf & strSheetName, vbCritical, "Err_SheetExists " & Err
    Resume Exit_SheetExists

End Function

Public Sub ProtectSheet(strSheetName As String)

    If gblnUnsecured Then Exit Sub
    Worksheets(strSheetName).Protect Password:=gstrPassword, DrawingObjects:=True, _
        Contents:=True, Scenarios:=True, UserInterfaceOnly:=True

End Sub

Public Sub UnprotectSheet(strSheetName As String)

    On Error Resume Next
    Worksheets(strSheetName).Unprotect gstrPassword
    
End Sub

Public Sub ProtectAllSheets()

    Dim obj As Object
    
    If gblnUnsecured Then Exit Sub
    For Each obj In Worksheets
        Worksheets(obj.Name).Protect Password:=gstrPassword, DrawingObjects:=True, _
            Contents:=True, Scenarios:=True
    Next obj
    
End Sub

Public Sub UnprotectAllSheets()

    Dim obj As Object
    
    For Each obj In Worksheets
        Worksheets(obj.Name).Unprotect gstrPassword
    Next obj
    
End Sub

Public Sub UnprotectWorkbook()

    ActiveWorkbook.Unprotect gstrPassword
    
End Sub

Public Sub ProtectWorkbook()

    If gblnUnsecured Then Exit Sub
    ActiveWindow.WindowState = xlMaximized
    ActiveWorkbook.Protect Structure:=True, Windows:=False, _
        Password:=gstrPassword
    
End Sub

Public Function GetFirstBlankCellNumber( _
            intStartRow As Integer, _
            intStartColumn As Integer) As Integer
'
' Description: Get First Blank Cell Number
' Date: 10/21/2002
' Author: (c) Peter F. Ennis
'
' Input:    intStartRow = the start row
'           intStartColumn = the start column
'               The function will progress down each row until
'               it finds a blank cell and then return that
'               row number.

    Dim strAddress As String
    Dim intColumn As Integer

    ' Find the first blank cell
    
    Application.GoTo Reference:="R" & intStartRow _
                                    & "C" & intStartColumn
    Do
        If IsEmpty(Application.ActiveCell.Value) Then
            'MsgBox "The Cell is BLANK"
            Exit Do
        Else
            'MsgBox "The Cell is NOT BLANK"
            ActiveCell.Offset(rowOffset:=1, columnoffset:=0).Activate
        End If
    Loop
    
    GetFirstBlankCellNumber = Application.Selection.Row

End Function






