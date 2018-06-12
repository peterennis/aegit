Option Compare Database
Option Explicit

' Ref: https://github.com/timabell/msaccess-vcs-integration/blob/master/MSAccess-VCS/VCS_File.bas
' Copyright © 2012 Brendan Kidwell et al
'
' Use of msaccess-vcs-integration and documentation are subject to the following
' BSD-style license:
'
' Permission to use, copy, modify, and/or distribute this software for any purpose
' with or without fee is hereby granted, provided that the above copyright notice
' and this permission notice appear in all copies.
'
' THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH
' REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY AND
' FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT,
' INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM LOSS
' OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER
' TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE OF
' THIS SOFTWARE.

' Updates by Peter Ennis 2017-18
' %001 - Only support VBA7
'
'

Private Declare PtrSafe Function getTempPath Lib "kernel32" _
    Alias "GetTempPathA" (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

Private Declare PtrSafe Function getTempFileName Lib "kernel32" _
    Alias "GetTempFileNameA" (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long

' Structure to track buffered reading or writing of binary files
Private Type BinFile
    file_num As Integer
    file_len As Long
    file_pos As Long
    buffer As String
    buffer_len As Integer
    buffer_pos As Integer
    at_eof As Boolean
    mode As String
End Type

Public Sub Test_UTF8Conversion()

    Dim strSourceFile As String
    Dim strDestinationFile As String
    
End Sub

Public Sub Test_aeExportToTextUnicode()
    On Error GoTo 0
    Dim bln As Boolean
    bln = ExportToTextUnicode("Items", "C:\Temp\ExportedItemsUnicode.txt")
End Sub

Public Function aeExportToTextUnicode(strTableName As String, strFileName As String, _
    Optional ByVal strDelim As String = vbTab) As Boolean
    ' Ref: https://saplsmw.com/Export_directly_to_UTF-8_from_Access_using_VBA
    ' Written by Jimbo at SAPLSMW.com
    ' Special thanks: accessblog.net/2007/06/how-to-write-out-unicode-text-files-in.html

    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim nCurrent As Long
    Dim nFieldCount As Long
    Dim nRecordCount As Long
    Dim RetVal As Variant
    Dim nCurRec As Long
    Dim dnow As Date
    Dim nCurSec As Long
    Dim nTotalSeconds As Long
    Dim nSecondsLeft As Long
    Dim strTest As String

    strSQL = "SELECT * FROM " & strTableName & ";"
    ' Check to see if strTableName is actually a query. If so, use its SQL query
    nCurrent = 0
    Do While nCurrent < CurrentDb.QueryDefs.Count
        If UCase(CurrentDb.QueryDefs(nCurrent).Name) = UCase(strTableName) Then
            strSQL = CurrentDb.QueryDefs(nCurrent).sql
        End If
        nCurrent = nCurrent + 1
    Loop
    Set rst = CurrentDb.OpenRecordset(strSQL)
    nFieldCount = rst.Fields.Count

    If Not rst.EOF Then
        ' Now find the *actual* record count--returns a value of 1 record if we don't do these moves
        rst.MoveLast
        rst.MoveFirst
    End If

    nRecordCount = rst.RecordCount
    RetVal = SysCmd(acSysCmdInitMeter, "Exporting " & strTableName & " to " & strFileName & ". . .", nRecordCount)
    ' Create a binary stream
    Dim UnicodeStream
    Set UnicodeStream = CreateObject("ADODB.Stream")
    UnicodeStream.Charset = "UTF-8"
    UnicodeStream.Open

    For nCurrent = 0 To nFieldCount - 1
        If Right(rst.Fields(nCurrent).Name, 1) = "_" Then
            UnicodeStream.WriteText Left(rst.Fields(nCurrent).Name, Len(rst.Fields(nCurrent).Name) - 1) & strDelim
        Else
            UnicodeStream.WriteText rst.Fields(nCurrent).Name & strDelim
        End If
    Next

    UnicodeStream.WriteText vbCrLf
    nCurSec = Second(Now())

    Do While Not rst.EOF
        nCurRec = nCurRec + 1
        If Second(Now()) <> nCurSec And nCurRec <> rst.RecordCount Then
            nCurSec = Second(Now())
            RetVal = SysCmd(acSysCmdUpdateMeter, nCurRec)
            RetVal = DoEvents()
        End If
        strTest = ""
        For nCurrent = 0 To nFieldCount - 1 ' Check for blank lines--no need to export those!
            strTest = strTest & IIf(IsNull(rst.Fields), "", rst.Fields(nCurrent))
        Next
        If Len(Trim(strTest)) > 0 Then      ' Check for blank lines--no need to export those!
            For nCurrent = 0 To nFieldCount - 1
                If Not IsNull(rst.Fields(nCurrent).Value) Then
                    UnicodeStream.WriteText Trim(rst.Fields(nCurrent).Value)
                End If
                If nCurrent = (nFieldCount - 1) Then
                    UnicodeStream.WriteText vbCrLf ' New line.
                Else
                    UnicodeStream.WriteText strDelim
                End If
            Next
        End If
        rst.MoveNext
    Loop

    ' Check to ensure that the file doesn't already exist
    If Len(Dir(strFileName)) > 0 Then
        Kill strFileName  ' The file exists, so we must delete it before it be created again
    End If
    UnicodeStream.SaveToFile strFileName
    UnicodeStream.Close
    rst.Close
    Set rst = Nothing
    aeExportToTextUnicode = True
    RetVal = SysCmd(acSysCmdRemoveMeter)

End Function

Private Function BinOpen(ByVal file_path As String, ByVal mode As String) As BinFile
    ' Open a binary file for reading (mode = 'r') or writing (mode = 'w')
    Dim f As BinFile

    f.file_num = FreeFile
    f.mode = LCase$(mode)
    If f.mode = "r" Then
        Open file_path For Binary Access Read As f.file_num
        f.file_len = LOF(f.file_num)
        f.file_pos = 0
        If f.file_len > &H4000 Then
            f.buffer = String$(&H4000, " ")
            f.buffer_len = &H4000
        Else
            f.buffer = String$(f.file_len, " ")
            f.buffer_len = f.file_len
        End If
        f.buffer_pos = 0
        Get f.file_num, f.file_pos + 1, f.buffer
    Else
        '''        VCS_DelIfExist file_path
        Open file_path For Binary Access Write As f.file_num
        f.file_len = 0
        f.file_pos = 0
        f.buffer = String$(&H4000, " ")
        f.buffer_len = 0
        f.buffer_pos = 0
    End If

    BinOpen = f
End Function

Private Function BinRead(ByRef f As BinFile) As Integer
    ' Buffered read one byte at a time from a binary file
    If f.at_eof = True Then
        BinRead = 0
        Exit Function
    End If

    BinRead = Asc(Mid$(f.buffer, f.buffer_pos + 1, 1))

    f.buffer_pos = f.buffer_pos + 1
    If f.buffer_pos >= f.buffer_len Then
        f.file_pos = f.file_pos + &H4000
        If f.file_pos >= f.file_len Then
            f.at_eof = True
            Exit Function
        End If
        If f.file_len - f.file_pos > &H4000 Then
            f.buffer_len = &H4000
        Else
            f.buffer_len = f.file_len - f.file_pos
            f.buffer = String$(f.buffer_len, " ")
        End If
        f.buffer_pos = 0
        Get f.file_num, f.file_pos + 1, f.buffer
    End If
End Function

Private Sub BinWrite(ByRef f As BinFile, b As Integer)
    ' Buffered write one byte at a time from a binary file
    Mid(f.buffer, f.buffer_pos + 1, 1) = Chr$(b)
    f.buffer_pos = f.buffer_pos + 1
    If f.buffer_pos >= &H4000 Then
        Put f.file_num, , f.buffer
        f.buffer_pos = 0
    End If
End Sub

Private Sub BinClose(ByRef f As BinFile)
    ' Close binary file
    If f.mode = "w" And f.buffer_pos > 0 Then
        f.buffer = Left$(f.buffer, f.buffer_pos)
        Put f.file_num, , f.buffer
    End If
    Close f.file_num
End Sub

Public Sub VCS_ConvertUcs2Utf8(ByVal Source As String, ByVal dest As String)
    ' Binary convert a UCS2-little-endian encoded file to UTF-8
    Dim f_in As BinFile
    Dim f_out As BinFile
    Dim in_low As Integer
    Dim in_high As Integer

    f_in = BinOpen(Source, "r")
    f_out = BinOpen(dest, "w")

    Do While Not f_in.at_eof
        in_low = BinRead(f_in)
        in_high = BinRead(f_in)
        If in_high = 0 And in_low < &H80 Then
            ' U+0000 - U+007F   0LLLLLLL
            BinWrite f_out, in_low
        ElseIf in_high < &H8 Then
            ' U+0080 - U+07FF   110HHHLL 10LLLLLL
            BinWrite f_out, &HC0 + ((in_high And &H7) * &H4) + ((in_low And &HC0) / &H40)
            BinWrite f_out, &H80 + (in_low And &H3F)
        Else
            ' U+0800 - U+FFFF   1110HHHH 10HHHHLL 10LLLLLL
            BinWrite f_out, &HE0 + ((in_high And &HF0) / &H10)
            BinWrite f_out, &H80 + ((in_high And &HF) * &H4) + ((in_low And &HC0) / &H40)
            BinWrite f_out, &H80 + (in_low And &H3F)
        End If
    Loop

    BinClose f_in
    BinClose f_out
End Sub