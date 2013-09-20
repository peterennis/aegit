Option Compare Database
Option Explicit

' Remove this after integration with aegitClass
Public Const THE_SOURCE_FOLDER = "C:\ae\aegit\aerc\src\"

Public Sub TestCreateDbScript()
    'CreateDbScript "C:\Temp\Schema.txt"
    Debug.Print "THE_SOURCE_FOLDER=" & THE_SOURCE_FOLDER
    CreateDbScript THE_SOURCE_FOLDER & "Schema.txt"
End Sub

Public Sub CreateDbScript(strScriptFile As String)
' Remou - Ref: http://stackoverflow.com/questions/698839/how-to-extract-the-schema-of-an-access-mdb-database/9910716#9910716

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim ndx As DAO.Index
    Dim strSQL As String
    Dim strFlds As String
    Dim strCn As String
    Dim strLinkedTablePath As String
    Dim fs As Object
    Dim f As Object

    Set db = CurrentDb
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.CreateTextFile(strScriptFile)

    strSQL = "Public Sub CreateTheDb()" & vbCrLf
    f.WriteLine strSQL
    strSQL = "Dim strSQL As String"
    f.WriteLine strSQL
    strSQL = "On Error GoTo ErrorTrap"
    f.WriteLine strSQL

    For Each tdf In db.TableDefs
        If Not (Left(tdf.Name, 4) = "MSys" _
                Or Left(tdf.Name, 4) = "~TMP" _
                Or Left(tdf.Name, 3) = "zzz") Then

            strLinkedTablePath = GetLinkedTableCurrentPath(tdf.Name)
            If strLinkedTablePath <> "" Then
                f.WriteLine vbCrLf & "'OriginalLink=>" & strLinkedTablePath
            Else
                f.WriteLine vbCrLf & "'Local Table"
            End If

            strSQL = "strSQL=""CREATE TABLE [" & tdf.Name & "] ("
            strFlds = ""

            For Each fld In tdf.Fields

                strFlds = strFlds & ",[" & fld.Name & "] "

                Select Case fld.Type
                    Case dbText
                        'No look-up fields
                        strFlds = strFlds & "Text (" & fld.Size & ")"
                    Case dbLong
                        If (fld.Attributes And dbAutoIncrField) = 0& Then
                            strFlds = strFlds & "Long"
                        Else
                            strFlds = strFlds & "Counter"
                        End If
                    Case dbBoolean
                        strFlds = strFlds & "YesNo"
                    Case dbByte
                        strFlds = strFlds & "Byte"
                    Case dbInteger
                        strFlds = strFlds & "Integer"
                    Case dbCurrency
                        strFlds = strFlds & "Currency"
                    Case dbSingle
                        strFlds = strFlds & "Single"
                    Case dbDouble
                        strFlds = strFlds & "Double"
                    Case dbDate
                        strFlds = strFlds & "DateTime"
                    Case dbBinary
                        strFlds = strFlds & "Binary"
                    Case dbLongBinary
                        strFlds = strFlds & "OLE Object"
                    Case dbMemo
                        If (fld.Attributes And dbHyperlinkField) = 0& Then
                            strFlds = strFlds & "Memo"
                        Else
                            strFlds = strFlds & "Hyperlink"
                        End If
                    Case dbGUID
                        strFlds = strFlds & "GUID"
                End Select

            Next

            strSQL = strSQL & Mid(strFlds, 2) & " )""" & vbCrLf & "Currentdb.Execute strSQL"
            f.WriteLine vbCrLf & strSQL

            'Indexes
            For Each ndx In tdf.Indexes

                If ndx.Unique Then
                    strSQL = "strSQL=""CREATE UNIQUE INDEX "
                Else
                    strSQL = "strSQL=""CREATE INDEX "
                End If

                strSQL = strSQL & "[" & ndx.Name & "] ON [" & tdf.Name & "] ("
                strFlds = ""

                For Each fld In tdf.Fields
                    strFlds = ",[" & fld.Name & "]"
                Next

                strSQL = strSQL & Mid(strFlds, 2) & ") "
                strCn = ""

                If ndx.Primary Then
                    strCn = " PRIMARY"
                End If

                If ndx.Required Then
                    strCn = strCn & " DISALLOW NULL"
                End If

                If ndx.IgnoreNulls Then
                    strCn = strCn & " IGNORE NULL"
                End If

                If Trim(strCn) <> vbNullString Then
                    strSQL = strSQL & " WITH" & strCn & " "
                End If

                f.WriteLine vbCrLf & strSQL & """" & vbCrLf & "Currentdb.Execute strSQL"
            Next
        End If
    Next

    'strSQL = vbCrLf & "Debug.Print " & """" & "Done" & """"
    'f.WriteLine strSQL
    f.WriteLine
    f.WriteLine "'Access 2010 - Compact And Repair"
    strSQL = "SendKeys " & """" & "%F{END}{ENTER}%F{TAB}{TAB}{ENTER}" & """" & ", False"
    f.WriteLine strSQL
    strSQL = "Exit Sub"
    f.WriteLine strSQL
    strSQL = "ErrorTrap:"
    f.WriteLine strSQL
    'MsgBox "Erl=" & Erl & vbCrLf & "Err.Number=" & Err.Number & vbCrLf & "Err.Description=" & Err.Description
    strSQL = "MsgBox " & """" & "Erl=" & """" & " & vbCrLf & " & _
                """" & "Err.Number=" & """" & " & Err.Number & vbCrLf & " & _
                """" & "Err.Description=" & """" & " & Err.Description"
    f.WriteLine strSQL & vbCrLf
    strSQL = "End Sub"
    f.WriteLine strSQL

    f.Close
    Debug.Print "Done"

End Sub

Public Function GetLinkedTableCurrentPath(MyLinkedTable As String) As String
' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=198057
' To test in the Immediate window:       ? getcurrentpath("Const")
'====================================================================
' Procedure : GetLinkedTableCurrentPath
' DateTime  : 08/23/2010
' Author    : Rx
' Purpose   : Returns Current Path of a Linked Table in Access
' Updates   : Peter F. Ennis
' Updated   : All notes moved to change log
' History   : See comment details, basChangeLog, commit messages on github
'====================================================================
    On Error GoTo PROC_ERR
    GetLinkedTableCurrentPath = Mid(CurrentDb.TableDefs(MyLinkedTable).Connect, InStr(1, CurrentDb.TableDefs(MyLinkedTable).Connect, "=") + 1)
        ' non-linked table returns blank - the Instr removes the "Database="

PROC_EXIT:
    On Error Resume Next
    Exit Function

PROC_ERR:
    Select Case Err.Number
        'Case ###         ' Add your own error management or log error to logging table
        Case Else
            'a custom log usage function commented out
            'function LogUsage(ByVal strFormName As String, strCallingProc As String, Optional ControlName) As Boolean
            'call LogUsage Err.Number, "basRelinkTables", "GetCurrentPath" ()
    End Select
    Resume PROC_EXIT
End Function

' Ref: http://www.utteraccess.com/forum/lofiversion/index.php/t1995627.html
'-------------------------------------------------------------------------------------------------
' Procedure : ExecSQL
' DateTime  : 30/03/2009 10:19
' Author    : Dial222
' Purpose   : Execute SQL Select statements in the Immediate window
' Context   : Module basSQL2IMM
' Notes     : No error trapping whatsover - this is a 1.0 technology!
'             Max out at 194 data rows since immediate only displays 100!
'
' Usage     : in the immediate pane: ?execsql("select * from zstblprofile","|")
'
' Revision History
' Version   Date        Who             What
' 01        30/03/2009  Dial222         Function 'ExecSQL' Created
' 02        30/03/2009  Dial222         Added code for left/right align of text/numeric data
'                                       Added MaxRowLen and vbCrLF parsing functionality
'                                       Uprated cMaxRows to 194
'-------------------------------------------------------------------------------------------------
'

Public Function ExecSQL(strSQL As String, Optional strColumDelim As String = "|") As Boolean

    Dim rs              As DAO.Recordset
    Dim aintLen()       As Integer
    Dim i               As Integer
    Dim str             As String
    Dim lngRowCOunt     As Long

    Const cMaxRows      As Integer = 194
    Const cMaxRowLen    As Integer = 1023  ' Max width of immediate pane in characters, truncate after this.

    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)

    With rs
        .MoveLast
        .MoveFirst

        lngRowCOunt = .RecordCount
        If lngRowCOunt > 0 Then
            If lngRowCOunt > cMaxRows Then
                Debug.Print "Too many rows to return, will only print first " & cMaxRows & " rows."
            End If

            ReDim Preserve aintLen(.Fields.Count)

            For i = 0 To .Fields.Count - 1
                ' Initialise field len to field name len
                aintLen(i) = Len(.Fields(i).Name) + 3
            Next i

            ' On this pass just get length of field data for formatting
            Do Until .EOF
                If .AbsolutePosition = cMaxRows Then
                    ' Stop at the magic number
                    Exit Do
                Else
                    For i = 0 To rs.Fields.Count - 1
                        ' Test and update field len
                        If Len(CStr(Nz(.Fields(i).Value, ""))) > aintLen(i) Then
                            aintLen(i) = Len(CStr(.Fields(i).Value)) + 3
                        End If
                    Next i
                End If
                .MoveNext
            Loop

            ' Print Column Headers
            str = "Row " & strColumDelim & " "
            For i = 0 To rs.Fields.Count - 1
                ' Initialise field len to field name len
                str = str & Left(.Fields(i).Name & Space(aintLen(i)), aintLen(i)) & " " & strColumDelim & " "
            Next i

            ' Print the header row
            Debug.Print Left(str, cMaxRowLen)
            str = Space(Len(str))
            str = Replace(str, " ", "-")

            ' print underscores
            Debug.Print Left(str, cMaxRowLen)
            str = ""

            ' Start over for the data
            .MoveFirst

            Do Until .EOF
                If .AbsolutePosition = cMaxRows Then
                    Exit Do
                Else
                    str = Left(.AbsolutePosition + 1 & Space(3), 3) & " " & strColumDelim & " "
                    For i = 0 To .Fields.Count - 1
                        Select Case .Fields(i).Type
                            Case Is = 3, 4, 5, 6, 7, 8, 16, 19, 20, 21, 22, 23 ' The numeric DataTypeEnums
                                str = str & Right(Space(aintLen(i)) & .Fields(i).Value, aintLen(i)) & " " & strColumDelim & " "
                            Case Else
                                ' Is it number stored as text
                                If IsNumeric(.Fields(i).Value) Then
                                    ' Right align
                                    str = str & Right(Space(aintLen(i)) & .Fields(i).Value, aintLen(i)) & " " & strColumDelim & " "
                                Else
                                    ' Left align
                                    str = str & Left(.Fields(i).Value & Space(aintLen(i)), aintLen(i)) & " " & strColumDelim & " "
                                End If
                        End Select
                    Next i
                End If

                ' Parse out vbCrLf and dump data row to immediate
                Debug.Print Left(Replace(Replace(str, Chr(13), " "), Chr(10), " "), cMaxRowLen)
                .MoveNext
                str = ""
            Loop

            ExecSQL = True
        Else
            Debug.Print "No rows returned"
        End If
    End With

    Set rs = Nothing

End Function

Public Function SpFolder(SpName)

    Dim objShell As Object
    Dim objFolder As Object
    Dim objFolderItem As Object

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(SpName)

    Set objFolderItem = objFolder.Self

    SpFolder = objFolderItem.Path

End Function
   
Public Sub AllCodeToDesktop()
' Ref: http://wiki.lessthandot.com/index.php/Code_and_Code_Windows
' Ref: http://stackoverflow.com/questions/2794480/exporting-code-from-microsoft-access
' The reference for the FileSystemObject Object is Windows Script Host Object Model
' but it not necessary to add the reference for this procedure.

    Const Desktop = &H10&
    Const MyDocuments = &H5&

    Dim fs As Object
    Dim f As Object
    Dim strMod As String
    Dim mdl As Object
    Dim i As Integer
    Dim strTxtFile As String

    Set fs = CreateObject("Scripting.FileSystemObject")

    'Set up the file
    Debug.Print "CurrentProject.Name = " & CurrentProject.Name
    strTxtFile = SpFolder(Desktop) & "\" & Replace(CurrentProject.Name, ".", " ") & ".txt"
    Debug.Print "strTxtFile = " & strTxtFile
    Set f = fs.CreateTextFile(SpFolder(Desktop) & "\" _
        & Replace(CurrentProject.Name, ".", " ") & ".txt")

    'For each component in the project ...
    For Each mdl In VBE.ActiveVBProject.VBComponents
        'using the count of lines ...
        i = VBE.ActiveVBProject.VBComponents(mdl.Name).codemodule.CountOfLines
        'put the code in a string ...
        If i > 0 Then
            strMod = VBE.ActiveVBProject.VBComponents(mdl.Name).codemodule.Lines(1, i)
        End If
        'and then write it to a file, first marking the start with
        'some equal signs and the component name.
        f.WriteLine String(15, "=") & vbCrLf & mdl.Name _
            & vbCrLf & String(15, "=") & vbCrLf & strMod
    Next
       
    'Close eveything
    f.Close
    Set fs = Nothing

End Sub

Public Function PropertyExists(obj As Object, strPropertyName As String) As Boolean
' Ref: http://www.utteraccess.com/forum/Description-property-Mic-t552348.html
' e.g. ? PropertyExists(CurrentDB. ("The Name Of Your Table"), "Description")
    Dim var As Variant

    On Error Resume Next
    Set var = obj.Properties(strPropertyName)
    If Err.Number > 0 Then
        PropertyExists = False
    Else
        PropertyExists = True
    End If

End Function

Public Sub GetPropertyDescription()
' Ref: http://www.dbforums.com/microsoft-access/1620765-read-ms-access-table-properties-using-vba.html

    Dim dbs As DAO.Database
    Dim obj As Object
    Dim prp As Property

    Set dbs = Application.CurrentDb
    Set obj = dbs.Containers("modules").Documents("aegitClass")

    On Error Resume Next
    For Each prp In obj.Properties
        Debug.Print prp.Name, prp.Value
    Next prp

    Set obj = Nothing
    Set dbs = Nothing

End Sub

Public Sub UTF2TXT_TestFunction()
' Access 2013 is saving text as UTF-16 (FFFE at the start of file)
' This is a test for fixing it
' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=241996
  
  If aeReadWriteStream("C:\ae\aegit\aerc\src\qry_HiddenDummy.qry") = True Then
      MsgBox "aeReadWriteStream succeeded"
    Else
      MsgBox "aeReadWriteStream failed"
  End If

End Sub

Public Function fn_ReadWriteStream(pFileName As String) As Boolean

    Dim fname As String
    Dim fname2 As String
    Dim fnr As Integer
    Dim fnr2 As Integer
    Dim tstring As String * 1
    Dim i As Integer
    Dim ByteCount As Integer

    fn_ReadWriteStream = False

    fname = pFileName
    fname2 = pFileName & ".clean.txt"

    fnr2 = FreeFile()
    Open fname2 For Binary Lock Read Write As #fnr2
    fnr = FreeFile()
    Open fname For Binary Access Read As #fnr
    Do
        Get #fnr, , tstring
        If EOF(fnr) Then Exit Do

        ByteCount = ByteCount + 1

        'If ByteCount < 10 Then
        '   MsgBox Asc(tstring)
        'End If

        If Asc(tstring) = 254 Or _
            Asc(tstring) = 255 Or _
            Asc(tstring) = 0 Then
        Else
            Put #fnr2, , tstring
        End If
    Loop

    Close #fnr
    Close #fnr2
    fn_ReadWriteStream = True

End Function

Public Function aeReadWriteStream(strPathFileName As String) As Boolean

    Dim fname As String
    Dim fname2 As String
    Dim fnr As Integer
    Dim fnr2 As Integer
    Dim tstring As String * 1
    Dim i As Integer
    Dim ByteCount As Integer

    aeReadWriteStream = False

    ' If the file has no Byte Order Mark (BOM)
    ' Ref: http://msdn.microsoft.com/en-us/library/windows/desktop/dd374101%28v=vs.85%29.aspx
    ' then do nothing
    fname = strPathFileName
    fname2 = strPathFileName & ".clean.txt"

    fnr = FreeFile()
    Open fname For Binary Access Read As #fnr
    Get #fnr, , tstring
    ' #FFFE, #FFFF, #0000
    ' If no BOM then it is a txt file and header stripping is not needed
    If Asc(tstring) <> 254 And Asc(tstring) <> 255 And _
                Asc(tstring) <> 0 Then
        Close #fnr
        aeReadWriteStream = True
        Exit Function
    End If

    fnr2 = FreeFile()
    Open fname2 For Binary Lock Read Write As #fnr2

    While Not EOF(fnr)

        'ByteCount = ByteCount + 1
        Get #fnr, , tstring

        'If ByteCount < 10 Then
        '   MsgBox Asc(tstring)
        'End If

        If Asc(tstring) = 254 Or Asc(tstring) = 255 Or _
                Asc(tstring) = 0 Then
        Else
            Put #fnr2, , tstring
        End If
    Wend

    Close #fnr
    Close #fnr2
    aeReadWriteStream = True

End Function