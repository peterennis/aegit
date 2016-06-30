Option Compare Database
Option Explicit

#Const conLateBinding = 0

Private mstrToParse As String

Public Sub GenerateLovefieldSchema()

    Dim strFileIn As String
    Dim strFileOut As String

    strFileIn = "C:\ae\aegit\aerc\src\OutputSchemaFile.txt.sql.only"
    strFileOut = ".\Out.txt"

    ReadInputWriteOutputLovefieldSchema strFileIn, strFileOut

End Sub

Public Sub ReadInputWriteOutputLovefieldSchema(ByVal strFileIn As String, ByVal strFileOut As String)

    'Debug.Print "ReadInputWriteOutputLovefieldSchema"
    On Error GoTo PROC_ERR

    Dim fleIn As Integer
    Dim fleOut As Integer
    Dim strIn As String
    Dim i As Integer
    Dim strLfCreateTable As String

    Dim dbs As DAO.Database
    Set dbs = CurrentDb()
    Dim strAppName As String
    strAppName = Application.VBE.ActiveVBProject.Name
    Dim strLfBegin As String
    strLfBegin = "// Begin schema creation" & vbCrLf & "var schemaBuilder = lf.schema.create('" & strAppName & "', 1);"

    fleOut = FreeFile()
    Open strFileOut For Output As #fleOut

    Dim arrSQL() As String
    i = 0
    fleIn = FreeFile()
    Open strFileIn For Input As #fleIn
    Do While Not EOF(fleIn)
        ReDim Preserve arrSQL(i)
        Line Input #fleIn, arrSQL(i)
        i = i + 1
    Loop
    Close fleIn

    'For i = 0 To UBound(arrSQL)
    '    Debug.Print i & ">", arrSQL(i)
    'Next

    Debug.Print strLfBegin

    For i = 0 To UBound(arrSQL)
        If Left$(arrSQL(i), 12) = "CREATE TABLE" Then
            ' Get the table name
            strLfCreateTable = "schemaBuilder.createTable('" & GetTableName(arrSQL(i)) & "')."
            Debug.Print i, strLfCreateTable
            mstrToParse = Right$(arrSQL(i), Len(arrSQL(i)) - InStr(arrSQL(i), "("))
            Do While mstrToParse <> vbNullString
                Debug.Print , GetFieldInfo(mstrToParse) ', mstrToParse
                'Stop
            Loop
        ElseIf Left$(arrSQL(i), 19) = "CREATE UNIQUE INDEX" Then
            ' Create the index
            'Print #fleOut, strSqlA
            Debug.Print i, arrSQL(i)
        End If
    Next
    Debug.Print "DONE !!!"

PROC_EXIT:
    Close fleIn
    Close fleOut
    Exit Sub

PROC_ERR:
    Select Case Err
        Case Else
            MsgBox "Erl=" & Erl & " Err=" & Err.Number & " (" & Err.Description & ") in procedure ReadInputWriteOutputLovefieldSchema of Class aegitClass"
            Resume PROC_EXIT
    End Select

End Sub

Public Sub GenerateLovefieldSchemaSample()
' Ref: https://github.com/google/lovefield/blob/master/docs/spec/01_schema.md

    Const APP_NAME As String = "aelfdb"
    Const LF_BEGIN As String = "// Begin schema creation" & vbCrLf & "var schemaBuilder = lf.schema.create('" & APP_NAME & "', 1);"
    
    Dim strTableName As String
    Dim strColumnName As String
    Dim strLfCreateTable As String

    strTableName = "Assets"
    strLfCreateTable = "schemaBuilder.createTable('" & strTableName & "')."

    Debug.Print LF_BEGIN
    Debug.Print strLfCreateTable
    strColumnName = "id"
    Debug.Print AddColumnString(strColumnName)
    strColumnName = "asset"
    Debug.Print AddColumnString(strColumnName)
    strColumnName = "timestamp"
    Debug.Print AddColumnInteger(strColumnName)
    strColumnName = "id"
    Debug.Print AddPrimaryKey(strColumnName)

End Sub

Private Function AddColumnString(ByVal strColName As String) As String
    AddColumnString = Space(4) & "addColumn('" & strColName & "', lf.Type.STRING)."
End Function

Private Function AddColumnInteger(ByVal strColName As String) As String
    AddColumnInteger = Space(4) & "addColumn('" & strColName & "', lf.Type.INTEGER)."
End Function

Private Function AddPrimaryKey(ByVal strColName As String) As String
    AddPrimaryKey = Space(4) & "addPrimaryKey('[" & strColName & "']);"
End Function

Public Function GetTableName(ByVal strSchemaLine As String) As String

    Dim strResult As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer

    intPos1 = InStr(1, strSchemaLine, "[")
    intPos2 = InStr(1, strSchemaLine, "]")
    GetTableName = Mid$(strSchemaLine, intPos1 + 1, intPos2 - intPos1 - 1)

End Function

Public Function GetFieldInfo(ByVal strSchemaLine As String) As String

    Dim strResult As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer

    intPos1 = InStr(1, strSchemaLine, "[")
    If InStr(1, strSchemaLine, ",") <> 0 Then
        intPos2 = InStr(1, strSchemaLine, ",")
    Else
        intPos2 = InStr(1, strSchemaLine, " )") + 1
    End If
    strResult = Mid$(strSchemaLine, intPos1 + 1, intPos2 - intPos1 - 1)
    GetFieldInfo = Replace(strResult, "]", vbNullString)
    ' Shorten the parse string by removing the found field
    mstrToParse = Right$(strSchemaLine, Len(strSchemaLine) - intPos2)
    If strSchemaLine = " )" Then mstrToParse = vbNullString
    
End Function

Public Sub TestOutputLovefieldFile()

    Dim strFileIn As String
    Dim strFileOut As String

    strFileIn = "C:\ae\aegit\aerc\src\OutputSchemaFile.txt.sql"
    strFileOut = ".\Out.txt"

    ReadInputWriteOutputSqlSchemaOnlyFile strFileIn, strFileOut

End Sub

Public Sub ReadInputWriteOutputSqlSchemaOnlyFile(ByVal strFileIn As String, ByVal strFileOut As String)

    'Debug.Print "ReadInputWriteOutputSqlSchemaOnlyFile"
    On Error GoTo PROC_ERR

    Dim fleIn As Integer
    Dim fleOut As Integer
    Dim strIn As String
    Dim i As Integer
    Dim strSqlA As String
    Dim strSqlB As String

    fleOut = FreeFile()
    Open strFileOut For Output As #fleOut

    Dim arrSQL() As String
    i = 0
    fleIn = FreeFile()
    Open strFileIn For Input As #fleIn
    Do While Not EOF(fleIn)
        ReDim Preserve arrSQL(i)
        Line Input #fleIn, arrSQL(i)
        i = i + 1
    Loop
    Close fleIn

    'For i = 0 To UBound(arrSQL)
    '    Debug.Print i & ">", arrSQL(i)
    'Next

    For i = 0 To UBound(arrSQL)
        If (i <> UBound(arrSQL)) Then
            If Left$(arrSQL(i + 1), 16) = "strSQL=strSQL & " Then
                If Left$(arrSQL(i), 7) = "strSQL=" Then
                    strSqlA = Right$(arrSQL(i), Len(arrSQL(i)) - 8)
                    strSqlA = Left$(strSqlA, Len(strSqlA) - 1)
                    Debug.Print i & ">", "strSqlA=" & strSqlA
                    strSqlB = Right$(arrSQL(i + 1), Len(arrSQL(i + 1)) - 17)
                    strSqlB = Left$(strSqlB, Len(strSqlB) - 1)
                    Debug.Print i & ">", "strSqlB=" & strSqlB
                    Print #fleOut, strSqlA & strSqlB
                    i = i + 1
                End If
            ElseIf Left$(arrSQL(i), 7) = "strSQL=" Then
                strSqlA = Right$(arrSQL(i), Len(arrSQL(i)) - 8)
                strSqlA = Left$(strSqlA, Len(strSqlA) - 1)
                Debug.Print i, strSqlA
                Print #fleOut, strSqlA
            End If
        Else
            If Left$(arrSQL(i), 7) = "strSQL=" Then
                strSqlA = Right$(arrSQL(i), Len(arrSQL(i)) - 8)
                strSqlA = Left$(strSqlA, Len(strSqlA) - 1)
                Debug.Print i, strSqlA
                Print #fleOut, strSqlA
            End If
        Debug.Print "UBound"
        End If
    Next
    Debug.Print "DONE !!!"

PROC_EXIT:
    Close fleIn
    Close fleOut
    Exit Sub

PROC_ERR:
    Select Case Err
    Case Else
        MsgBox "Erl=" & Erl & " Err=" & Err.Number & " (" & Err.Description & ") in procedure ReadInputWriteOutputSqlSchemaOnlyFile of Class aegitClass"
        Resume PROC_EXIT
    End Select

End Sub

'Private Function FoundSqlInLine(ByVal strLine As String, Optional ByVal varEnd As Variant) As Boolean
'
'    'Debug.Print "FoundSqlInLine"
'    On Error GoTo 0
'
'    FoundSqlInLine = False
'    If Not IsMissing(varEnd) Then
'        If InStr(1, strLine, "strSQL=strSQL & ", vbTextCompare) > 0 Then
'            FoundSqlInLine = True
'        Else
'            FoundSqlInLine = True
'        End If
'        Exit Function
'    ElseIf InStr(1, strLine, "strSQL=", vbTextCompare) > 0 Then
'        FoundSqlInLine = True
'        Exit Function
'    End If
'
'End Function

Public Function FileDelete(ByVal strFileName As String) As Boolean
    On Error GoTo 0
    If Len(Dir$(strFileName)) > 0 Then
        Kill strFileName
    End If
End Function

Public Sub TestGetSQLServerData()
    On Error GoTo 0
    Dim bln As Boolean
    bln = GetSQLServerData(".\SQLEXPRESS", "AdventureWorks2012")
End Sub

Public Function GetSQLServerData(ByVal strServer As String, ByVal strDatabase As String) As Boolean
' Ref: http://www.saplsmw.com/node/11
' Ref: http://www.eileenslounge.com/viewtopic.php?f=29&t=5886
' Ref: *** http://social.msdn.microsoft.com/Forums/office/en-US/00c3f331-15e6-44f2-9e6f-abede3a986d8/sql-server-data-connectivity-best-practices-for-ms-access-vba ***
' http://www.connectionstrings.com/sql-server-native-client-11-0-odbc-driver/
' oConn.Properties("Prompt") = adPromptAlways
' oConn.Open "Driver={SQL Server Native Client 11.0};Server=myServerAddress;Database=myDataBase;"

    On Error GoTo 0
    Dim strODBC As String
    strODBC = "DRIVER={SQL Server Native Client 11.0};SERVER=" & strServer & ";DATABASE=" & _
        strDatabase & ";Trusted_Connection=No"
    Debug.Print strODBC

#If conLateBinding = 1 Then
    Dim cnn As Object
    ' Ref: http://support.microsoft.com/kb/195982
    Const adPromptAlways As Integer = 1
    ' Ref: http://www.w3schools.com/ado/met_rs_open.asp#CommandTypeEnum
    Const adOpenDynamic As Integer = 2
    Const adLockOptimistic As Integer = 3
    Dim orst As Object
    Dim ofld As Object
    Dim ocat As Object
    Dim otbl As Object
    Dim oind As Object
    Dim ocol As Object
#Else
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    cnn.ConnectionString = strODBC
    cnn.Properties("Prompt") = adPromptAlways
    Dim orst As ADODB.Recordset
    Set orst = New ADODB.Recordset
    Dim ofld As ADODB.Field
    Dim ocat As ADOX.Catalog
    Set ocat = New ADOX.Catalog
    ocat.ActiveConnection = cnn
    Dim otbl As ADOX.Table
    Set otbl = New ADOX.Table
    Dim oind As ADOX.Index
    Dim ocol As ADOX.Column
#End If

    Dim vx As Variant
    Dim intSec As Integer

    'Stop

    ' Setup the local connection in this database
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim rst As DAO.Recordset

    ' Ref: http://technet.microsoft.com/en-us/library/ms189082(v=sql.105).aspx
    Debug.Print "Count of items in tables catalog=" & ocat.Tables.Count
    For Each otbl In ocat.Tables
        If otbl.Name = "AWBuildVersion" Then
        ElseIf otbl.Name = "DatabaseLog" Then
        ElseIf otbl.Name = "ErrorLog" Then
        Else
        Debug.Print otbl.Type, otbl.Name
        If otbl.Type = "TABLE" Then

            orst.Open "HumanResources." & otbl.Name, cnn, adOpenDynamic, adLockOptimistic
            For Each tdf In dbs.TableDefs
                'Debug.Print , tdf.Name
                If tdf.Name = otbl.Name Then
                    ' This table already exists!  Delete it.
                    dbs.TableDefs.Delete (otbl.Name)
                End If
            Next

            Set tdf = dbs.CreateTableDef(otbl.Name)
            For Each ocol In otbl.Columns
                Debug.Print "ocol=" & ocol.Name, "DefinedSize=" & ocol.DefinedSize
                If ocol.DefinedSize < 256 And ocol.DefinedSize <> 0 Then
                    Set fld = tdf.CreateField(ocol.Name, dbText, ocol.DefinedSize)
                ElseIf ocol.DefinedSize = 0 Then
                    Set fld = tdf.CreateField(ocol.Name, dbText, 255)
                Else
                    'MsgBox "ocol.DefinedSize=" & ocol.DefinedSize
                    Set fld = tdf.CreateField(ocol.Name, dbMemo, 255)
                    'Stop
                End If
                fld.AllowZeroLength = True
                fld.Required = False
                tdf.Fields.Append fld
            Next
            dbs.TableDefs.Append tdf
            dbs.TableDefs.Refresh

            Set rst = dbs.OpenRecordset(otbl.Name)
            Do While Not orst.EOF
                rst.AddNew
                For Each ofld In orst.Fields
                    ' Avoid type mismatch with trim and &""
                    Debug.Print "ofld.Name=" & ofld.Name
                    rst.Fields(ofld.Name).Value = Trim$(vbNullString & ofld.Value)
                Next
                rst.Update
                orst.MoveNext
                If Second(Now) <> intSec Then
                    ' update once each second
                    intSec = Second(Now)
                    vx = DoEvents()
                End If
            Loop
            rst.Close

            orst.Close
        End If
        End If
    Next
    dbs.Close
    cnn.Close
    GetSQLServerData = True

    ' Cleanup to prevent memory leaks
    Set cnn = Nothing
    Set ocat = Nothing
    Set otbl = Nothing
    Set oind = Nothing
    Set ocol = Nothing
    Set orst = Nothing
    Set ofld = Nothing
    Set dbs = Nothing
    Set dbs = Nothing
    Set tdf = Nothing
    Set fld = Nothing
    Set rst = Nothing

End Function