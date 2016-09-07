Option Compare Database
Option Explicit

Private Function IsQryHidden(ByVal strQueryName As String) As Boolean
    'Debug.Print "IsQryHidden"
    On Error GoTo 0
    If IsNull(strQueryName) Or strQueryName = vbNullString Then
        IsQryHidden = False
        'Debug.Print "IsQryHidden Null Test", strQueryName, IsQryHidden
    Else
        IsQryHidden = GetHiddenAttribute(acQuery, strQueryName)
        'Debug.Print "IsQryHidden Attribute Test", strQueryName, IsQryHidden
    End If
End Function

Private Function IsIndex(ByVal tdf As DAO.TableDef, ByVal strField As String) As Boolean
    'Debug.Print "IsIndex"
    On Error GoTo 0

    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tdf.Indexes
        For Each fld In idx.Fields
            If strField = fld.Name Then
                IsIndex = True
                Exit Function
            End If
        Next fld
    Next idx
End Function

Private Function IsFK(ByVal tdf As DAO.TableDef, ByVal strField As String) As Boolean
    'Debug.Print "IsFK"
    On Error GoTo 0
    
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tdf.Indexes
        If idx.Foreign Then
            For Each fld In idx.Fields
                If strField = fld.Name Then
                    IsFK = True
                    Exit Function
                End If
            Next fld
        End If
    Next idx
End Function

Public Sub TestGetIndexDetails()

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim tdf As DAO.TableDef
    Set tdf = dbs.TableDefs("tblDummy3")
    Dim fld As DAO.Field
    Dim ndx As DAO.Index

    For Each ndx In tdf.Indexes
        Debug.Print ndx.Name, ndx.Fields, ndx.Foreign, ndx.IgnoreNulls, ndx.Primary, ndx.Properties.Count, ndx.Required, ndx.Unique
    Next
    Debug.Print "--------------------"

    For Each fld In tdf.Fields
        ' Testing show info for multi-field PK and multi-field index
        Debug.Print fld.Name, GetIndexDetails(tdf, fld.Name)
    Next

End Sub

Public Function GetIndexDetails(tdf As DAO.TableDef, strField As String) As String
' Ref: allenbrowne.com DescribeIndexField

    'Debug.Print "GetIndexDetails"
    Dim ind As DAO.Index
    Dim fld As DAO.Field
    Dim iCount As Integer
    Dim strReturn As String

    For Each ind In tdf.Indexes
        iCount = 0
        For Each fld In ind.Fields
            If fld.Name = strField Then
                If ind.Primary Then
                    strReturn = strReturn & IIf(iCount = 0, "P", "p")
                ElseIf ind.Unique Then
                    strReturn = strReturn & IIf(iCount = 0, "U", "u")
                Else
                    strReturn = strReturn & IIf(iCount = 0, "I", "i")
                End If
            End If
            iCount = iCount + 1
        Next
    Next
    GetIndexDetails = strReturn

End Function

Public Function DescribeIndexField(tdf As DAO.TableDef, strField As String) As String
' allenbrowne.com
'Purpose:   Indicate if the field is part of a primary key or unique index.
'Return:    String containing "P" if primary key, "U" if uniuqe index, "I" if non-unique index.
'           Lower case letters if secondary field in index. Can have multiple indexes.
'Arguments: tdf = the TableDef the field belongs to.
'           strField = name of the field to search the Indexes for.
    Dim ind As DAO.Index        'Each index of this table.
    Dim fld As DAO.Field        'Each field of the index
    Dim iCount As Integer
    Dim strReturn As String     'Return string
    
    For Each ind In tdf.Indexes
        iCount = 0
        For Each fld In ind.Fields
            If fld.Name = strField Then
                If ind.Primary Then
                    strReturn = strReturn & IIf(iCount = 0, "P", "p")
                ElseIf ind.Unique Then
                    strReturn = strReturn & IIf(iCount = 0, "U", "u")
                Else
                    strReturn = strReturn & IIf(iCount = 0, "I", "i")
                End If
            End If
            iCount = iCount + 1
        Next
    Next
    DescribeIndexField = strReturn
End Function

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

Public Function WUAversion() As String
' Get current WUA version
    Dim objAgentInfo As Object
    On Error Resume Next
    Err.Clear
    Set objAgentInfo = CreateObject("Microsoft.Update.AgentInfo")
    If Err = 0 Then
        WUAversion = objAgentInfo.GetInfo("ProductVersionString")
        Debug.Print , "wuapi.dll version: " & objAgentInfo.GetInfo("ProductVersionString")
        Debug.Print , "WUA version: " & objAgentInfo.GetInfo("ApiMajorVersion") & "." & objAgentInfo.GetInfo("ApiMinorVersion")
    Else
        MsgBox "Error getting WUA version.", vbCritical, "WUA Version"
        WUAversion = 0 ' Calling code can interpret 0 as an error
    End If
    On Error GoTo 0
End Function

Private Function zzzLongestTableName() As Integer
' ====================================================================
' Author:   Peter F. Ennis
' Date:     November 30, 2012
' Comment:  Return the length of the longest table name
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
' ====================================================================

    Dim tdf As DAO.TableDef
    Dim intTNLen As Integer

    Debug.Print "LongestTableName"
    On Error GoTo PROC_ERR

    intTNLen = 0
    For Each tdf In CurrentDb.TableDefs
        If Not (Left$(tdf.Name, 4) = "MSys" _
                Or Left$(tdf.Name, 4) = "~TMP" _
                Or Left$(tdf.Name, 3) = "zzz") Then
            If Len(tdf.Name) > intTNLen Then
                intTNLen = Len(tdf.Name)
            End If
        End If
    Next tdf

    zzzLongestTableName = intTNLen

PROC_EXIT:
    Set tdf = Nothing
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestTableName of Class aegit_expClass", vbCritical, "ERROR"
    zzzLongestTableName = 0
    Resume PROC_EXIT

End Function