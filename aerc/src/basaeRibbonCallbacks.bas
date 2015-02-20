Option Compare Database
Option Explicit

Public gobjaeRibbon As IRibbonUI

Public Sub OnRibbonLoad(ByVal ribbon As IRibbonUI)
    ' Callbackname in XML File "onLoad"
    On Error GoTo 0
    Set gobjaeRibbon = ribbon
End Sub

Public Sub OnActionButton(ByVal control As IRibbonControl)
    ' Callbackname in XML File "onAction"
    Select Case control.Id
        Case Else
            MsgBox "Button """ & control.Id & """ clicked!" & vbCrLf, vbInformation
    End Select
End Sub

Sub GetEnabled(ByVal control As IRibbonControl, ByRef enabled)
    ' Callbackname in XML File "getEnabled"
    On Error GoTo 0
    Select Case control.Id
        Case Else
            enabled = True
    End Select
End Sub

Sub GetVisible(ByVal control As IRibbonControl, ByRef visible)
    ' Callbackname in XML File "getVisible"
    On Error GoTo 0
    Select Case control.Id
        Case Else
            visible = True
    End Select
End Sub

Public Sub GetImages(ByVal control As IRibbonControl, ByRef Image)

    On Error GoTo 0
    Dim strPicture As String

    strPicture = getTheValue(control.Tag, "Pic")
    Set Image = getIconFromTable(strPicture)

End Sub

Private Function getTheValue(ByVal strTag As String, ByVal strValue As String) As String
    ' *************************************************************
    ' Created from     : Avenius
    ' Parameter        : Input String, SuchValue String
    ' Date created     : 05.01.2008
    '
    ' Sample:
    ' getTheValue("DefaultValue:=Test;Enabled:=0;Visible:=1", "DefaultValue")
    ' Return           : "Test"
    ' *************************************************************
      
   On Error Resume Next
      
   Dim workTb()     As String
   Dim Ele()        As String
   Dim myVariabs()  As String
   Dim i            As Integer

      workTb = Split(strTag, ";")
      
      ReDim myVariabs(LBound(workTb) To UBound(workTb), 0 To 1)
      For i = LBound(workTb) To UBound(workTb)
         Ele = Split(workTb(i), ":=")
         myVariabs(i, 0) = Ele(0)
         If UBound(Ele) = 1 Then
            myVariabs(i, 1) = Ele(1)
         End If
      Next
      
      For i = LBound(myVariabs) To UBound(myVariabs)
         If strValue = myVariabs(i, 0) Then
            getTheValue = myVariabs(i, 1)
         End If
      Next
      
End Function

Public Function getIconFromTable(ByVal strFileName As String) As Picture

    Dim lSize As Long
    Dim arrBin() As Byte
    Dim rst As DAO.Recordset

    On Error GoTo PROC_ERR

    Set rst = DBEngine(0)(0).OpenRecordset("tblBinary", dbOpenDynaset)
    rst.FindFirst "[FileName]='" & strFileName & "'"
    If rst.NoMatch Then
        Set getIconFromTable = Nothing
    Else
        lSize = rst.Fields("binary").FieldSize
        ReDim arrBin(lSize)
        arrBin = rst.Fields("binary").GetChunk(0, lSize)
        Set getIconFromTable = ArrayToPicture(arrBin)
    End If
    rst.Close

PROC_EXIT:
    Reset
    Erase arrBin
    Set rst = Nothing
    Exit Function

PROC_ERR:
    Resume PROC_EXIT

End Function