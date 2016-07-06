Option Compare Database
Option Explicit

#Const conLateBinding = 0

'Public Sub GenerateLovefieldSchemaSample()
'' Ref: https://github.com/google/lovefield/blob/master/docs/spec/01_schema.md
'
'    Const APP_NAME As String = "aelfdb"
'    Const LF_BEGIN As String = "// Begin schema creation" & vbCrLf & "var schemaBuilder = lf.schema.create('" & APP_NAME & "', 1);"
'
'    Dim strTableName As String
'    Dim strColumnName As String
'    Dim strLfCreateTable As String
'
'    strTableName = "Assets"
'    strLfCreateTable = "schemaBuilder.createTable('" & strTableName & "')."
'
'    Debug.Print LF_BEGIN
'    Debug.Print strLfCreateTable
'    strColumnName = "id"
'    Debug.Print AddColumnString(strColumnName)
'    strColumnName = "asset"
'    Debug.Print AddColumnString(strColumnName)
'    strColumnName = "timestamp"
'    Debug.Print AddColumnInteger(strColumnName)
'    strColumnName = "id"
'    Debug.Print AddPrimaryKey(strColumnName)
'
'End Sub
'
'Private Function AddColumnString(ByVal strColName As String) As String
'    AddColumnString = Space(4) & "addColumn('" & strColName & "', lf.Type.STRING)."
'End Function
'
'Private Function AddColumnInteger(ByVal strColName As String) As String
'    AddColumnInteger = Space(4) & "addColumn('" & strColName & "', lf.Type.INTEGER)."
'End Function
'
'Private Function AddPrimaryKey(ByVal strColName As String) As String
'    AddPrimaryKey = Space(4) & "addPrimaryKey('[" & strColName & "']);"
'End Function

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