Option Compare Database
Option Explicit

Public Function Test_IsSingleIndexField() As Boolean

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Set dbs = CurrentDb
    Dim FieldCountRes As Integer

    Set tdf = dbs.TableDefs("tblDummy3")
'    Debug.Print tdf.Name
    Debug.Print , IsSingleIndexField(tdf, FieldCountRes), FieldCountRes

End Function

Private Function IsSingleIndexField(ByVal tdf As DAO.TableDef, ByRef FieldCountResult As Integer) As Boolean

    Dim strIndexInfo As String
    strIndexInfo = SingleTableIndexSummary(tdf)
    Debug.Print strIndexInfo
    FieldCountResult = LCaseCountChar("I", strIndexInfo)
    If FieldCountResult = 1 Then
        IsSingleIndexField = True
        Debug.Print , "Single Field Index", "IsSingleIndexField is " & IsSingleIndexField
    ElseIf FieldCountResult > 1 Then
        IsSingleIndexField = False
        Debug.Print , "Multi Field Index", "IsSingleIndexField is " & IsSingleIndexField
    ElseIf FieldCountResult = 0 Then
        IsSingleIndexField = False
        Debug.Print , "No Index", "IsSingleIndexField is " & IsSingleIndexField
    End If

End Function

Public Function Test_IsSinglePrimaryField() As Boolean

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Set dbs = CurrentDb
    Dim IndexPrimaryFieldCount As Integer

    Set tdf = dbs.TableDefs("aeItems")
    Debug.Print tdf.Name
    Debug.Print , IsSinglePrimaryField(tdf, IndexPrimaryFieldCount), IndexPrimaryFieldCount

    Set tdf = dbs.TableDefs("tblDummy2")
    Debug.Print tdf.Name
    Debug.Print , IsSinglePrimaryField(tdf, IndexPrimaryFieldCount), IndexPrimaryFieldCount

    Set tdf = dbs.TableDefs("tblDummy3")
    Debug.Print tdf.Name
    Debug.Print , IsSinglePrimaryField(tdf, IndexPrimaryFieldCount), IndexPrimaryFieldCount

End Function

Private Function IsSinglePrimaryField(ByVal tdf As DAO.TableDef, ByRef PrimaryIndexFieldCount As Integer) As Boolean

    Dim strIndexInfo As String
    strIndexInfo = SingleTableIndexSummary(tdf)
    PrimaryIndexFieldCount = LCaseCountChar("P", strIndexInfo)
    If PrimaryIndexFieldCount = 1 Then
        IsSinglePrimaryField = True
        'Debug.Print , strIndexInfo, "Single Field Primary Key", IsSinglePrimaryField
    ElseIf PrimaryIndexFieldCount > 1 Then
        IsSinglePrimaryField = False
        Debug.Print , strIndexInfo, "Multi Field Primary Key"
    ElseIf PrimaryIndexFieldCount = 0 Then
        IsSinglePrimaryField = False
        'Debug.Print , strIndexInfo, "No Primary Key"
    End If

End Function

Public Function AllTablesIndexSummary() As Boolean

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strTableIndexInfo As String
    strTableIndexInfo = vbNullString
    Set dbs = CurrentDb
    For Each tdf In dbs.TableDefs
        ' Ignore system and temporary tables
        If Not (tdf.Name Like "MSys*" Or tdf.Name Like "~*") Then
            strTableIndexInfo = SingleTableIndexSummary(tdf)
            Debug.Print tdf.Name, strTableIndexInfo
        End If
    Next
    Set tdf = Nothing
    Set dbs = Nothing

End Function

Private Function SingleTableIndexSummary(ByVal tdf As DAO.TableDef) As String

    Dim strIndexFieldInfo As Variant
    Dim fld As DAO.Field

    'On Error Resume Next
    Debug.Print tdf.Name
'    For Each fld In tdf.Fields
        strIndexFieldInfo = aeDescribeIndexField(tdf)
'    Next

End Function

Private Function LCaseCountChar(ByVal searchChar As String, ByVal searchString As String) As Long
    Dim i As Long
    For i = 1 To Len(searchString)
        If Mid$(LCase$(searchString), i, 1) = LCase(searchChar) Then
        LCaseCountChar = LCaseCountChar + 1
    End If
    Next
End Function

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

Private Function aeDescribeIndexField(tdf As DAO.TableDef) As String()
    ' Based on work from allenbrowne.com
    ' Purpose:   Get details of all indexes in a table.
    ' Return:    Array of descriptors and field names.
    '            String containing "P" if primary key, "U" if unique index, "I" if non-unique index.
    '            Lower case letters if secondary field in index. Can have multiple indexes.
    ' Arguments: tdf = the TableDef the field belongs to.

    Dim idx As DAO.Index        ' Each index of this table
    Dim fld As DAO.Field        ' Each field of the index
    Dim iCount As Integer
    Dim iCountMax As Integer
    Dim iCountTotal As Integer
    Dim arrTemp() As String
    ReDim arrTemp(1, 0)         ' Ref: http://stackoverflow.com/questions/13183775/excel-vba-how-to-redim-a-2d-array
    Dim arrReturn() As String   ' Return array
    ReDim arrReturn(1, 0)       ' Ref: http://stackoverflow.com/questions/13183775/excel-vba-how-to-redim-a-2d-array
    Dim i As Integer
    Dim j As Integer
    Dim jFirst As Integer
    Dim jLast As Integer
    Dim blnNextIndex As Boolean

    jFirst = 0
    iCountTotal = 0
    For Each idx In tdf.Indexes
        iCount = iCountTotal
        Debug.Print ":", idx.Name
        For Each fld In idx.Fields
            If idx.Primary Then
                arrTemp(0, iCount) = arrTemp(0, iCount) & IIf(iCount = iCountTotal, "P", "p")
                arrTemp(1, iCount) = fld.Name
            ElseIf idx.Unique Then
                arrTemp(0, iCount) = arrTemp(0, iCount) & IIf(iCount = iCountTotal, "U", "u")
                arrTemp(1, iCount) = fld.Name
            Else
                arrTemp(0, iCount) = arrTemp(0, iCount) & IIf(iCount = iCountTotal, "I", "i")
                arrTemp(1, iCount) = fld.Name
            End If
            iCount = iCount + 1
            iCountMax = iCount
            ReDim Preserve arrTemp(1, iCount)
        Next
        Debug.Print "> ", "iCountMax = " & iCountMax
        jLast = iCountMax - 1
        iCountTotal = iCountTotal + iCountMax
        Debug.Print "> ", "iCountTotal = " & iCountTotal
        ReDim Preserve arrReturn(1, iCountTotal)
        ReDim Preserve arrTemp(1, iCountTotal)

        If blnNextIndex Then
            Debug.Print "blnNextIndex = " & blnNextIndex, "jFirst = " & jFirst, "jLast = " & jLast, "iCountMax = " & iCountMax, "iCountTotal = " & iCountTotal
            jFirst = iCountTotal - iCountMax
            jLast = jFirst + iCountMax - 1
            blnNextIndex = False
            Debug.Print "blnNextIndex = " & blnNextIndex, "jFirst = " & jFirst, "jLast = " & jLast, "iCountMax = " & iCountMax, "iCountTotal = " & iCountTotal
            'Stop
        End If
        For i = 0 To 1
            For j = jFirst To jLast
                'Debug.Print i, j, jFirst,
                arrReturn(i, j) = arrTemp(i, j)
                Debug.Print arrReturn(i, j),
            Next
            Debug.Print
        Next
        j = 0
        jFirst = iCountTotal - 1
        blnNextIndex = True
    Next
    Debug.Print ">> ", "iCountTotal = " & iCountTotal
    aeDescribeIndexField = arrReturn()
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