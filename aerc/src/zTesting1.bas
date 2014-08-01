Option Compare Database
Option Explicit

'zzzTmpTblQueries
Public Function RecordsetUpdatable(ByVal strSQL As String) As Boolean
' Ref: http://msdn.microsoft.com/en-us/library/office/ff193796(v=office.15).aspx

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim intPosition As Integer

    On Error GoTo PROC_ERR

    ' Initialize the function's return value to True.
    RecordsetUpdatable = True

    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(strSQL, dbOpenDynaset)

    ' If the entire dynaset isn't updatable, return False.
    If rst.Updatable = False Then
        RecordsetUpdatable = False
    Else
        ' If the dynaset is updatable, check if all fields in the
        ' dynaset are updatable. If one of the fields isn't updatable,
        ' return False.
        For intPosition = 0 To rst.Fields.Count - 1
            If rst.Fields(intPosition).DataUpdatable = False Then
                RecordsetUpdatable = False
                Exit For
            End If
        Next intPosition
    End If

PROC_EXIT:
    rst.Close
    dbs.Close
    Set rst = Nothing
    Set dbs = Nothing
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RecordsetUpdatable of Class aegit_expClass"
    Resume Next

End Function

Public Sub ApplicationInformation()
' Ref: http://msdn.microsoft.com/en-us/library/office/aa223101(v=office.11).aspx
' Ref: http://msdn.microsoft.com/en-us/library/office/aa173218(v=office.11).aspx
' Ref: http://msdn.microsoft.com/en-us/library/office/ff845735(v=office.15).aspx

    On Error GoTo 0
    Dim intProjType As Integer
    Dim strProjType As String
    Dim lng As Long

    intProjType = Application.CurrentProject.ProjectType

    Select Case intProjType
        Case 0 ' acNull
            strProjType = "acNull"
        Case 1 ' acADP
            strProjType = "acADP"
        Case 2 ' acMDB
            strProjType = "acMDB"
        Case Else
            MsgBox "Can't determine ProjectType"
    End Select

    Debug.Print Application.CurrentProject.FullName
    Debug.Print "Project Type", intProjType, strProjType
    lng = CodeLinesInProjectCount

End Sub

Public Sub TestRegKey()

    On Error GoTo 0
    Dim strKey As String

    ' Office 2010
    'strKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Common\General\ReportAddinCustomUIErrors"
    ' Office 2013
    strKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Common\General\ReportAddinCustomUIErrors"
    If RegKeyExists(strKey) Then
        Debug.Print strKey & " Exists!"
    Else
        Debug.Print strKey & " Does NOT Exist!"
    End If

End Sub

Public Function RegKeyExists(ByVal strRegKey As String) As Boolean
' Return True if the registry key i_RegKey was found and False if not
' Ref: http://vba-corner.livejournal.com/3054.html

    On Error GoTo ErrorHandler

    ' Use Windows scripting and try to read the registry key
    Dim myWS As Object
    Set myWS = CreateObject("WScript.Shell")

    myWS.RegRead strRegKey
    ' Key was found
    RegKeyExists = True
    Exit Function
  
ErrorHandler:
    ' Key was not found
    RegKeyExists = False

End Function

Public Function WuaVersion()
' Get current WUA version
    Dim oAgentInfo As Object
    Dim ProductVersion As String
    Dim ErrNum As Long
    On Error Resume Next
    Err.Clear
    Set oAgentInfo = CreateObject("Microsoft.Update.AgentInfo")
    If ErrNum = 0 Then
        WuaVersion = oAgentInfo.GetInfo("ProductVersionString")
    Else
        MsgBox "Error getting WUA version.", vbCritical, "WUA Version"
        WuaVersion = 0 ' Calling code can interpret 0 as an error
    End If
    On Error GoTo 0
End Function