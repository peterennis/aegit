Option Compare Database
Option Explicit

'?strTran("Dev" & vbCrLf & "ashish", vbCrLf, " ")
'Dev ashish
'
'Function strTran(ByVal sInString As String, _
                            sFindString As String, _
                            sReplaceString As String, _
                            Optional iCount As Variant) As String
'   Dim iSpot As Integer, iCtr As Integer
'   If IsMissing(iCount) Then iCount = 1
'   If iCount = 0 Then iCount = 1000
'   For iCtr = 1 To iCount
'     iSpot = InStr(1, sInString, sFindString)
'     If iSpot > 0 Then
'       sInString = Left(sInString, iSpot - 1) & _
'                         sReplaceString & _
'                         Mid(sInString, iSpot + Len(sFindString))
'     Else
'       Exit For
'     End If
'   Next
'   strTran = sInString
'
'End Function
'
'http://computer-programming-forum.com/1-vba/34d339bb6472eb9d.htm

Public Sub CatalogUserCreatedObjects()
' Ref: http://blogannath.blogspot.com/2010/03/microsoft-access-tips-tricks-working.html#ixzz3WCBJcxwc
' Ref: http://stackoverflow.com/questions/5286620/saving-a-query-via-access-vba-code

    Dim strSQL As String
    Const MY_QUERY_NAME = "zzzqryCatalogUserCreatedObjects"

    strSQL = "SELECT IIf(type = 1,""Table"", IIf(type = 6, ""Linked Table"", "
    strSQL = strSQL & vbCrLf & "IIf(type = 5,""Query"", IIf(type = -32768,""Form"", "
    strSQL = strSQL & vbCrLf & "IIf(type = -32764,""Report"", IIf(type=-32766,""Module"", "
    strSQL = strSQL & vbCrLf & "IIf(type = -32761,""Module"", ""Unknown""))))))) as [Object Type], "
    strSQL = strSQL & vbCrLf & "MSysObjects.Name, MSysObjects.DateCreate, MSysObjects.DateUpdate "
    strSQL = strSQL & vbCrLf & "FROM MSysObjects "
    strSQL = strSQL & vbCrLf & "WHERE Type IN (1, 5, 6, -32768, -32764, -32766, -32761) "
    strSQL = strSQL & vbCrLf & "AND Left(Name, 4) <> ""MSys"" AND Left(Name, 1) <> ""~"" "
    strSQL = strSQL & vbCrLf & "ORDER BY IIf(type=1,""Table"",IIf(type=6,""Linked Table"",IIf(type=5,""Query"",IIf(type=-32768,""Form"",IIf(type=-32764,""Report"",IIf(type=-32766,""Module"",IIf(type=-32761,""Module"",""Unknown""))))))), MSysObjects.Name;"

    Debug.Print strSQL
    
    ' Using a query name and sql string, if the query does not exist, ...
    If IsNull(DLookup("Name", "MsysObjects", "Name='" & MY_QUERY_NAME & "'")) Then
        ' create it ...
        CurrentDb.CreateQueryDef MY_QUERY_NAME, strSQL
    Else
        ' other wise, update the sql
        CurrentDb.QueryDefs(MY_QUERY_NAME).SQL = strSQL
    End If

    DoCmd.OpenQuery MY_QUERY_NAME

End Sub