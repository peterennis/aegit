'* * * * * * * * * * * * * * * * * * * *
'*                                     *  +--------------------------+
'*      Written by James Kauffman      *  |                          |
'*                                     *  |  http://www.saplsmw.com  |
'*     Ver 1.20 Updated 17Jun2010      *  |                          |
'*                                     *  +--------------------------+
'* * * * * * * * * * * * * * * * * * * *

Option Compare Database

Function FileDelete(strFileName As String) As Boolean
    'Use this function to delete a file
    If Len(Dir(strFileName)) > 0 Then
        Kill strFileName
    End If
End Function