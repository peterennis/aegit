Option Compare Database
Option Explicit
Option Private Module


' Validate the table field is a primary key
'
' Returns:
'
' TRUE - The field is a primary key
' FALSE - The field is NOT a primary key
Public Function zzzValidateIsPK(ByVal tdf As DAO.TableDef, ByVal strField As String) As Boolean

    Dim success As Boolean: success = False

    ' Validate the password is of correct length
    If IsPK(tdf, strField) Then
        success = True
    End If

    zzzValidateIsPK = success

End Function