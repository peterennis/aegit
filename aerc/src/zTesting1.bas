Option Compare Database
Option Explicit

Public Sub Setup_Test_aeDescribeIndexField()

    Dim arrTest() As String
    ReDim arrTest(1, 0)
    Dim blnTest As Boolean
    Dim strField As String
    Dim strTestName As String

    ' Using tblDummy fields as test example
    strTestName = "T1:"
    arrTest(0, 0) = "P"
    strField = "eventId"
    arrTest(1, 0) = strField
    blnTest = IsSinglePrimaryField(arrTest, strField)
    Debug.Print strTestName
    ShowTestArray arrTest
    If blnTest Then
        Debug.Print , strTestName & " " & strField & " Is a Single Primary Field"
    Else
        Debug.Print , strTestName & " " & strField & " Is NOT a Single Primary Field"
    End If

End Sub

Private Sub ShowTestArray(arr() As String)
    Debug.Print , "LBound: " & CStr(LBound(arr, 2)), _
        "UBound: " & CStr(UBound(arr, 2)), _
        "NumElements: " & CStr(UBound(arr, 2) - LBound(arr, 2) + 1)
    Dim i As Integer
    For i = LBound(arr, 2) To UBound(arr, 2)
        Debug.Print , arr(0, i), arr(1, i)
    Next
End Sub

Public Sub Test_aeDescribeIndexField()

    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim tdf As DAO.TableDef
    Set tdf = dbs.TableDefs("tblDummy3")       ' List of test tables:
                                                            ' tblDummy3, Sheet1_DS1_DEMOG, dbo_rights, dbo_studentAttendances

    Dim i As Integer

    Debug.Print tdf.Name
    Dim arrIndexFieldInfo() As String
    arrIndexFieldInfo = aeDescribeIndexField(tdf)
    Debug.Print ":::arrIndexFieldInfo", "LBound: " & CStr(LBound(arrIndexFieldInfo, 2)), _
        "UBound: " & CStr(UBound(arrIndexFieldInfo, 2)), _
        "NumElements: " & CStr(UBound(arrIndexFieldInfo, 2) - LBound(arrIndexFieldInfo, 2) + 1)

    For i = LBound(arrIndexFieldInfo, 2) To UBound(arrIndexFieldInfo, 2) - LBound(arrIndexFieldInfo, 2)
        Debug.Print arrIndexFieldInfo(0, i), arrIndexFieldInfo(1, i)
    Next i

End Sub

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
    Dim iCountMaxMem As Integer
    Dim iCountMaxFirst As Integer
    Dim iCountIndex As Integer
    Dim iCountIndexFields As Integer
    Dim iArray As Integer
    Dim arrTemp() As String
    Dim arrReturn() As String   ' Return array
    Dim i As Integer
    Dim j As Integer
    Dim jFirst As Integer
    Dim jLast As Integer
    Dim blnNextIndex As Boolean

    ' ReDim the arrays size for all fields
    ' Ref: http://stackoverflow.com/questions/13183775/excel-vba-how-to-redim-a-2d-array
    iCountIndexFields = 0
    For Each idx In tdf.Indexes
        For Each fld In idx.Fields
        iCountIndexFields = iCountIndexFields + 1
        Next
    Next
    Debug.Print "iCountIndexFields = " & iCountIndexFields
    ReDim Preserve arrTemp(1, iCountIndexFields - 1)
    ReDim Preserve arrReturn(1, iCountIndexFields - 1)

    jFirst = 0
    jLast = 0
    iCountIndex = 0
    iArray = 0
    blnNextIndex = False

    For Each idx In tdf.Indexes
        iCount = 0
        arrTemp(0, iCount + jFirst) = vbNullString
        arrTemp(1, iCount + jFirst) = vbNullString
        For Each fld In idx.Fields
            If idx.Primary Then
                arrTemp(0, iCount + jFirst) = arrTemp(0, iCount + jFirst) & IIf(iCount = 0, "P", "p")
                arrTemp(1, iCount + jFirst) = fld.Name
            ElseIf idx.Unique Then
                arrTemp(0, iCount + jFirst) = arrTemp(0, iCount + jFirst) & IIf(iCount = 0, "U", "u")
                arrTemp(1, iCount + jFirst) = fld.Name
            Else
                arrTemp(0, iCount + jFirst) = arrTemp(0, iCount + jFirst) & IIf(iCount = 0, "I", "i")
                arrTemp(1, iCount + jFirst) = fld.Name
            End If
            If iCount = 0 Then iCountMaxFirst = iCountMaxFirst + iCountMaxMem
            iCount = iCount + 1
            iCountMaxMem = iCount
            Debug.Print "iArray = " & iArray, "iCountIndex = " & iCountIndex, "jFirst = " & jFirst, "iCount = " & iCount, "jLast = " & jLast, idx.Name, arrTemp(0, iCount - 1 + jFirst), arrTemp(1, iCount - 1 + jFirst)
            arrReturn(0, iArray) = arrTemp(0, iCount - 1 + jFirst)
            arrReturn(1, iArray) = arrTemp(1, iCount - 1 + jFirst)
            iArray = iArray + 1
            arrTemp(0, iCount - 1 + jFirst) = vbNullString
            arrTemp(1, iCount - 1 + jFirst) = vbNullString
        Next

        jFirst = iCountMaxFirst
        jLast = jFirst + iCountMaxMem - 1
        'Debug.Print "INDEX: " & idx.Name, "iCountIndex = " & iCountIndex, "jFirst = " & jFirst, "jLast = " & jLast, "iCountMaxFirst = " & iCountMaxFirst, "iCountMaxMem = " & iCountMaxMem
        iCountIndex = iCountIndex + 1
    Next

'        For i = 0 To 1
            For j = 0 To iCountIndexFields - 1
                Debug.Print i, j, arrReturn(0, j), arrReturn(1, j)
            Next
'        Next

    Debug.Print ":iCountIndexFields = " & iCountIndexFields
    ' Ref: http://stackoverflow.com/questions/26644231/vba-using-ubound-on-a-multidimensional-array
    Debug.Print "::arrReturn", "LBound: " & CStr(LBound(arrReturn, 2)), _
        "UBound: " & CStr(UBound(arrReturn, 2)), _
        "NumElements: " & CStr(UBound(arrReturn, 2) - LBound(arrReturn, 2) + 1)
    aeDescribeIndexField = arrReturn()
End Function

Private Function IsSingleIndexField(arr() As String, ByVal strFieldName As String) As Boolean
    Debug.Print "ADD TEST CODE FOR IsSingleIndexField"
End Function

Private Function IsSinglePrimaryField(arr() As String, ByVal strFieldName As String) As Boolean
    Debug.Print "ADD TEST CODE FOR IsSinglePrimaryField"
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

Public Function DescribeIndexField(tdf As DAO.TableDef, strField As String) As String
    ' allenbrowne.com
    ' Purpose:   Indicate if the field is part of a primary key or unique index.
    ' Return:    String containing "P" if primary key, "U" if uniuqe index, "I" if non-unique index.
    '           Lower case letters if secondary field in index. Can have multiple indexes.
    ' Arguments: tdf = the TableDef the field belongs to.
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