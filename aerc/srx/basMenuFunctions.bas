Attribute VB_Name = "basMenuFunctions"
Option Explicit

Public Sub aeXL_CheckTheMenuBits(lngNum As Long)
' What: Pass a Hex number from wksCmdSetup sheet to set menu buttons.
'       cmd0 not used as this bit would set negative numbers.
'       This leaves 31 command buttons available on the menu.
' Date: 04/04/2007
' Who: Peter F. Ennis
'
' Ref: http://www.xtremevbtalk.com/archive/index.php/t-32396.html
' To test individual bits, use the AND operator.
' If x And 1 Then 'bit 0 is set
' If x And 2 Then 'bit 1 is set
' If x And 4 Then 'bit 2 is set
' If x And 8 Then 'bit 3 is set
' etc.
'
' If you want to rotate through the number, a simple loop should do
' For i = 0 To 7
' If x And 1 Then
' do something because bit i is set
' End If
' x = x \ 2 'shift right
' Next i

On Error GoTo Err_aeXL_CheckTheMenuBits

1:    Dim i As Integer
      Debug.Print "lngNum=" & lngNum
2:    For i = 0 To 31
3:        If lngNum And 1 Then
4:            Debug.Print "bit=" & i & " is set", "cmd" & Abs(i - 31)
5:        End If
6:        lngNum = lngNum \ 2
7:    Next i

Exit_aeXL_CheckTheMenuBits:
    Exit Sub
    
Err_aeXL_CheckTheMenuBits:
    MsgBox Erl & " " & Err.Description, vbCritical, "aeXL_CheckTheMenuBits Err=" & Err
    Resume Exit_aeXL_CheckTheMenuBits

End Sub

Public Function aeControlTipText(strCmd As String) As String
' What: Return the command control tip text from wksCmdSetup
' Date: 04/11/2007
' Who:  Peter F. Ennis

On Error GoTo Err_aeWksName

      Dim strAddress As String
1:    strAddress = aeWksLookup(Range("CmdLookup"), strCmd)
2:    aeControlTipText = Worksheets("wksCmdSetup").Range(strAddress).Offset(-2, 0)

Exit_aeWksName:
    Exit Function
    
Err_aeWksName:
    If Err = 1004 Then      ' Ignore error on labels
        Resume Next
    Else
        MsgBox Erl & " " & Err.Description, vbCritical, "aeControlTipText Err=" & Err
        Resume Exit_aeWksName
    End If
    
End Function

Public Function aeWksName(strWks As String) As String
' What: Return the worksheet name from wksCmdSetup
' Date: 04/07/2007
' Who:  Peter F. Ennis

On Error GoTo Err_aeWksName

      Dim strAddress As String
1:    strAddress = aeWksLookup(Range("HexLookup"), strWks)
2:    aeWksName = Worksheets("wksCmdSetup").Range(strAddress).Offset(0, -1)

Exit_aeWksName:
    Exit Function
    
Err_aeWksName:
    If Err = 1004 Then
        Resume Next
    Else
        MsgBox Erl & " " & Err.Description, vbCritical, "aeWksName Err=" & Err
        Resume Exit_aeWksName
    End If
    
End Function

Public Function aeWksHex(strWks As String) As String
' What: Return the Hex value for menu display from wksCmdSetup sheet
' Date: 04/07/2007
' Who:  Peter F. Ennis

On Error GoTo Err_aeWksHex

    Dim strAddress As String
1:    strAddress = aeWksLookup(Range("HexLookup"), strWks)
      'Debug.Print "strAddress=" & strAddress
      'Debug.Print "Worksheets(""wksCmdSetup"").Range(strAddress)=" & Worksheets("wksCmdSetup").Range(strAddress)
      'Debug.Print "Worksheets(""wksCmdSetup"").Range(strAddress).Offset(0, 42)=" & Worksheets("wksCmdSetup").Range(strAddress).Offset(0, 42)
2:    aeWksHex = Worksheets("wksCmdSetup").Range(strAddress).Offset(0, 42)

Exit_aeWksHex:
    Exit Function
    
Err_aeWksHex:
    MsgBox Erl & " " & Err.Description, vbCritical, "aeWksHex Err=" & Err
    Resume Exit_aeWksHex
    
End Function

Public Function aeWksLookup(The_Range As Range, The_Sheet_Name As String) As String
' What: Return the address of The_Sheet_Name
' Date: 04/06/2007
' Who:  Peter F. Ennis
' Example:  ?aeWksLookup(Range("HexLookup"),"wks1")

    Dim cel As Range

    For Each cel In The_Range
        If cel.Value = The_Sheet_Name Then
            aeWksLookup = cel.Address
            'Debug.Print "aeWksLookup=" & aeWksLookup
            Exit For
        End If
    Next cel

End Function


