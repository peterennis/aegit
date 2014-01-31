Option Compare Database
Option Explicit

' Ref: http://www.cpearson.com/excel/sizestring.htm
' This enum is used by SizeString to indicate whether the supplied text
' appears on the left or right side of result string.
Private Enum SizeStringSide
    TextLeft = 1
    TextRight = 2
End Enum

Private aeintFNLen As Long
Private aestrLFN As String
Private aestrLFNTN As String
Private aeintFDLen As Long
Private aestrLFD As String
Private aeintFTLen As Long
Private aestrLFT As String
Private Const aestr4 As String = "    "
Private Const aeintFSize As Long = 4

' aeDocumentTables "debug"
Public Sub aeDocumentTables(Optional varDebug As Variant)
' Ref: http://www.tek-tips.com/faqs.cfm?fid=6905
' Document the tables, fields, and relationships
' Tables, field type, primary keys, foreign keys, indexes
' Relationships in the database with table, foreign table, primary keys, foreign keys
' Ref: http://allenbrowne.com/func-06.html

    Dim strDocument As String
    Dim tblDef As DAO.TableDef
    Dim fld As DAO.Field
    Dim blnDebug As Boolean
    Dim blnResult As Boolean
    Dim intFailCount As Integer
    Dim strFile As String

    ' Use a call stack and global error handler
    'If gcfHandleErrors Then On Error GoTo PROC_ERR
    'PushCallStack "aeDocumentTables"

    On Error GoTo PROC_ERR

    intFailCount = 0
    
    LongestFieldPropsName
    Debug.Print "Longest Field Name=" & aestrLFN
    Debug.Print "Longest Field Name Length=" & aeintFNLen
    Debug.Print "Longest Field Name Table Name=" & aestrLFNTN
    Debug.Print "Longest Field Description=" & aestrLFD
    Debug.Print "Longest Field Description Length=" & aeintFDLen
    Debug.Print "Longest Field Type=" & aestrLFT
    Debug.Print "Longest Field Type Length=" & aeintFTLen

    ' Reset values
    aestrLFN = ""
    aeintFNLen = 11     ' Minimum required by design
    'aestrLFNTN = ""
    'aestrLFD = ""
    aeintFDLen = 0
    'aestrLFT = ""
    'aeintFTLen = 0

    Debug.Print "aeDocumentTables"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeDocumentTables is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeDocumentTables is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    'strFile = aestrSourceLocation & aeTblTxtFile
    '
    'If Dir(strFile) <> "" Then
    '    ' The file exists
    '    If Not FileLocked(strFile) Then KillProperly (strFile)
    '    Open strFile For Append As #1
    'Else
    '    If Not FileLocked(strFile) Then Open strFile For Append As #1
    'End If

    For Each tblDef In CurrentDb.TableDefs
        If Not (Left(tblDef.Name, 4) = "MSys" _
                Or Left(tblDef.Name, 4) = "~TMP" _
                Or Left(tblDef.Name, 3) = "zzz") Then
            If blnDebug Then
                blnResult = TableInfo(tblDef.Name, "WithDebugging")
                If Not blnResult Then intFailCount = intFailCount + 1
                If blnDebug And aeintFDLen <> 11 Then Debug.Print "aeintFDLen=" & aeintFDLen
            Else
                blnResult = TableInfo(tblDef.Name)
                If Not blnResult Then intFailCount = intFailCount + 1
            End If
            'Debug.Print
            aeintFDLen = 0
        End If
    Next tblDef

    'If intFailCount > 0 Then
    '    aeDocumentTables = False
    'Else
    '    aeDocumentTables = True
    'End If
    If blnDebug Then
        Debug.Print "intFailCount = " & intFailCount
        'Debug.Print "aeDocumentTables = " & aeDocumentTables
    End If

    'aeDocumentTables = True

PROC_EXIT:
    Set fld = Nothing
    Set tblDef = Nothing
    Close 1
    'PopCallStack
    Exit Sub

PROC_ERR:
    MsgBox "erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTables of Class aegitClass"
    If blnDebug Then Debug.Print ">>>erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTables of Class aegitClass"
    'aeDocumentTables = False
    'GlobalErrHandler
    Resume PROC_EXIT

End Sub

Public Function TableInfo(strTableName As String, Optional varDebug As Variant) As Boolean
' Ref: http://allenbrowne.com/func-06.html
'====================================================================
' Purpose:  Display the field names, types, sizes and descriptions for a table
' Argument: Name of a table in the current database
' Updates:  Peter F. Ennis
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim sLen As Long
    Dim strLinkedTablePath As String
    
    Dim blnDebug As Boolean

    ' Use a call stack and global error handler
    'If gcfHandleErrors Then On Error GoTo PROC_ERR
    'PushCallStack "TableInfo"

    On Error GoTo PROC_ERR

    strLinkedTablePath = ""

    If IsMissing(varDebug) Then
        blnDebug = False
        'Debug.Print , "varDebug IS missing so blnDebug of TableInfo is set to False"
        'Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        'Debug.Print , "varDebug IS NOT missing so blnDebug of TableInfo is set to True"
        'Debug.Print , "NOW DEBUGGING..."
    End If

    Set dbs = CurrentDb()
    Set tdf = dbs.TableDefs(strTableName)
    sLen = Len("TABLE: ") + Len(strTableName)

    strLinkedTablePath = GetLinkedTableCurrentPath(strTableName)

    aeintFDLen = LongestTableDescription(tdf.Name)

    If aeintFDLen < Len("DESCRIPTION") Then aeintFDLen = Len("DESCRIPTION")

    If blnDebug Then
    'If blnDebug And aeintFDLen <> 11 Then
        Debug.Print SizeString("-", sLen, TextLeft, "-")
        Debug.Print SizeString("TABLE: " & strTableName, sLen, TextLeft, " ")
        Debug.Print SizeString("-", sLen, TextLeft, "-")
        If strLinkedTablePath <> "" Then
            Debug.Print strLinkedTablePath
        End If
        Debug.Print SizeString("FIELD NAME", aeintFNLen, TextLeft, " ") _
                        & aestr4 & SizeString("FIELD TYPE", aeintFTLen, TextLeft, " ") _
                        & aestr4 & SizeString("SIZE", aeintFSize, TextLeft, " ") _
                        & aestr4 & SizeString("DESCRIPTION", aeintFDLen, TextLeft, " ")
        Debug.Print SizeString("=", aeintFNLen, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFTLen, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFSize, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFDLen, TextLeft, "=")
    End If

    'Print #1, SizeString("-", sLen, TextLeft, "-")
    'Print #1, SizeString("TABLE: " & strTableName, sLen, TextLeft, " ")
    'Print #1, SizeString("-", sLen, TextLeft, "-")
    'If strLinkedTablePath <> "" Then
    '    Print #1, "Linked=>" & strLinkedTablePath
    'End If
    'Print #1, SizeString("FIELD NAME", aeintFNLen, TextLeft, " ") _
                        & aestr4 & SizeString("FIELD TYPE", aeintFTLen, TextLeft, " ") _
                        & aestr4 & SizeString("SIZE", aeintFSize, TextLeft, " ") _
                        & aestr4 & SizeString("DESCRIPTION", aeintFDLen, TextLeft, " ")
    'Print #1, SizeString("=", aeintFNLen, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFTLen, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFSize, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFDLen, TextLeft, "=")
    strLinkedTablePath = ""

    For Each fld In tdf.Fields
        If blnDebug Then
        'If blnDebug And aeintFDLen <> 11 Then
            Debug.Print SizeString(fld.Name, aeintFNLen, TextLeft, " ") _
                & aestr4 & SizeString(FieldTypeName(fld), aeintFTLen, TextLeft, " ") _
                & aestr4 & SizeString(fld.size, aeintFSize, TextLeft, " ") _
                & aestr4 & SizeString(GetDescrip(fld), aeintFDLen, TextLeft, " ")
        End If
        'Print #1, SizeString(fld.Name, aeintFNLen, TextLeft, " ") _
            & aestr4 & SizeString(FieldTypeName(fld), aeintFTLen, TextLeft, " ") _
            & aestr4 & SizeString(fld.Size, aeintFSize, TextLeft, " ") _
            & aestr4 & SizeString(GetDescrip(fld), aeintFDLen, TextLeft, " ")
    Next
    If blnDebug Then Debug.Print
    'If blnDebug And aeintFDLen <> 11 Then Debug.Print
    'Print #1, vbCrLf

    TableInfo = True

PROC_EXIT:
    Set fld = Nothing
    Set tdf = Nothing
    Set dbs = Nothing
    'PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TableInfo of Class aegitClass"
    If blnDebug Then Debug.Print ">>>erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TableInfo of Class aegitClass"
    TableInfo = False
    'GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function GetLinkedTableCurrentPath(MyLinkedTable As String) As String
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
        ' Case ###         ' Add your own error management or log error to logging table
        Case Else
            ' A custom log usage example function commented out:
            ' function LogUsage(ByVal strFormName As String, strCallingProc As String, Optional ControlName) As Boolean
            ' call LogUsage err.Number, "basRelinkTables", "GetCurrentPath" ()
    End Select
    Resume PROC_EXIT
End Function

Private Function SizeString(Text As String, Length As Long, _
    Optional ByVal TextSide As SizeStringSide = TextLeft, _
    Optional PadChar As String = " ") As String
' Ref: http://www.cpearson.com/excel/sizestring.htm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SizeString
' This procedure creates a string of a specified length. Text is the original string
' to include, and Length is the length of the result string. TextSide indicates whether
' Text should appear on the left (in which case the result is padded on the right with
' PadChar) or on the right (in which case the string is padded on the left). When padding on
' either the left or right, padding is done using the PadChar. character. If PadChar is omitted,
' a space is used. If PadChar is longer than one character, the left-most character of PadChar
' is used. If PadChar is an empty string, a space is used. If TextSide is neither
' TextLeft or TextRight, the procedure uses TextLeft.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim sPadChar As String

    If Len(Text) >= Length Then
        ' if the source string is longer than the specified length, return the
        ' Length left characters
        SizeString = Left(Text, Length)
        Exit Function
    End If

    If Len(PadChar) = 0 Then
        ' PadChar is an empty string. use a space.
        sPadChar = " "
    Else
        ' use only the first character of PadChar
        sPadChar = Left(PadChar, 1)
    End If

    If (TextSide <> TextLeft) And (TextSide <> TextRight) Then
        ' if TextSide was neither TextLeft nor TextRight, use TextLeft.
        TextSide = TextLeft
    End If

    If TextSide = TextLeft Then
        ' if the text goes on the left, fill out the right with spaces
        SizeString = Text & String(Length - Len(Text), sPadChar)
    Else
        ' otherwise fill on the left and put the Text on the right
        SizeString = String(Length - Len(Text), sPadChar) & Text
    End If

End Function

Private Function FieldTypeName(fld As DAO.Field) As String
' Ref: http://allenbrowne.com/func-06.html
' Purpose: Converts the numeric results of DAO Field.Type to text
    Dim strReturn As String    ' Name to return

    Select Case CLng(fld.Type) ' fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            '  1
        Case dbByte: strReturn = "Byte"                 '  2
        Case dbInteger: strReturn = "Integer"           '  3
        Case dbLong                                     '  4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         '  5
        Case dbSingle: strReturn = "Single"             '  6
        Case dbDouble: strReturn = "Double"             '  7
        Case dbDate: strReturn = "Date/Time"            '  8
        Case dbBinary: strReturn = "Binary"             '  9 (no interface)
        Case dbText                                     ' 10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        ' (no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     ' 11
        Case dbMemo                                     ' 12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 ' 15

        ' Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        ' 16
        Case dbVarBinary: strReturn = "VarBinary"       ' 17
        Case dbChar: strReturn = "Char"                 ' 18
        Case dbNumeric: strReturn = "Numeric"           ' 19
        Case dbDecimal: strReturn = "Decimal"           ' 20
        Case dbFloat: strReturn = "Float"               ' 21
        Case dbTime: strReturn = "Time"                 ' 22
        Case dbTimeStamp: strReturn = "Time Stamp"      ' 23

        ' Constants for complex types don't work
        ' prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"             ' dbAttachment
        Case 102&: strReturn = "Complex Byte"           ' dbComplexByte
        Case 103&: strReturn = "Complex Integer"        ' dbComplexInteger
        Case 104&: strReturn = "Complex Long"           ' dbComplexLong
        Case 105&: strReturn = "Complex Single"         ' dbComplexSingle
        Case 106&: strReturn = "Complex Double"         ' dbComplexDouble
        Case 107&: strReturn = "Complex GUID"           ' dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"        ' dbComplexDecimal
        Case 109&: strReturn = "Complex Text"           ' dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn

End Function

Private Function GetDescrip(obj As Object) As String
    On Error Resume Next
    GetDescrip = obj.Properties("Description")
End Function

Public Function LongestTableDescription(strTblName As String) As Integer
' ?LongestTableDescription("tblCaseManager")

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim strLFD As String

    ' Use a call stack and global error handler
    'If gcfHandleErrors Then On Error GoTo PROC_ERR
    'PushCallStack "LongestTableDescription"

    Set dbs = CurrentDb()
    Set tdf = dbs.TableDefs(strTblName)

    For Each fld In tdf.Fields
        If Len(GetDescrip(fld)) > aeintFDLen Then
            strLFD = GetDescrip(fld)
            aeintFDLen = Len(GetDescrip(fld))
        End If
    Next

    LongestTableDescription = aeintFDLen

PROC_EXIT:
    Set fld = Nothing
    Set tdf = Nothing
    Set dbs = Nothing
    'PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestFieldPropsName of Class aegitClass"
    LongestTableDescription = -1
    'GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function LongestFieldPropsName() As Boolean
'====================================================================
' Author:   Peter F. Ennis
' Date:     December 5, 2012
' Comment:  Return length of field properties for text output alignment
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim dbs As DAO.Database
    Dim tblDef As DAO.TableDef
    Dim fld As DAO.Field

    ' Use a call stack and global error handler
    'If gcfHandleErrors Then On Error GoTo PROC_ERR
    'PushCallStack "LongestFieldPropsName"

    On Error GoTo PROC_ERR

    aeintFNLen = 0
    aeintFTLen = 0
    aeintFDLen = 0

    Set dbs = CurrentDb()

    For Each tblDef In CurrentDb.TableDefs
        If Not (Left(tblDef.Name, 4) = "MSys" _
                Or Left(tblDef.Name, 4) = "~TMP" _
                Or Left(tblDef.Name, 3) = "zzz") Then
            For Each fld In tblDef.Fields
                If Len(fld.Name) > aeintFNLen Then
                    aestrLFNTN = tblDef.Name
                    aestrLFN = fld.Name
                    aeintFNLen = Len(fld.Name)
                End If
                If Len(FieldTypeName(fld)) > aeintFTLen Then
                    aestrLFT = FieldTypeName(fld)
                    aeintFTLen = Len(FieldTypeName(fld))
                End If
                If Len(GetDescrip(fld)) > aeintFDLen Then
                    aestrLFD = GetDescrip(fld)
                    aeintFDLen = Len(GetDescrip(fld))
                End If
            Next
        End If
    Next tblDef

    LongestFieldPropsName = True

PROC_EXIT:
    Set fld = Nothing
    Set tblDef = Nothing
    Set dbs = Nothing
    'PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestFieldPropsName of Class aegitClass"
    LongestFieldPropsName = False
    'GlobalErrHandler
    Resume PROC_EXIT
    
End Function

Public Sub ExportRibbon()

    Dim strDB As String
    Dim lngPath As Long
    Dim lngRev As Long
    Dim strLeft As String

    strDB = Application.CurrentDb.Name
    lngPath = Len(strDB)
    lngRev = InStrRev(strDB, "\")
    strLeft = Left(strDB, lngPath - (lngPath - lngRev))

    DoCmd.OutputTo acOutputTable, "listview", acFormatTXT, strLeft & "OutputRibbon.txt"

End Sub

Public Function LoadRibbons() As Boolean
' Load ribbons from XML file into the database
' Ref: http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Load_Ribbons_Into_The_Database:..._From_XML_File

    On Error GoTo PROC_ERR
    
    Dim f As Long
    Dim strText As String
    Dim strOut As String

    f = FreeFile
    Open "C:\Folder\Ribbon\AccRibbon.xml" For Input As f
    ' C:\Folder\... has to be replaced by your folder/filename.
    Do While Not EOF(f)
        Line Input #f, strText
        strOut = strOut & strText
    Loop
    Application.LoadCustomUI "AppRibbon_1", strOut

PROC_EXIT:
    On Error Resume Next
    Close f
    'PopCallStack
    Exit Function

PROC_ERR:
    Select Case Err
        Case 32609
        ' Ribbon already loaded
    Case Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LoadRibbons of Class aegitClass"
    End Select
    LoadRibbons = False
    'GlobalErrHandler
    Resume PROC_EXIT

End Function

Public Sub CreateRibbon()
' Ref: http://www.nullskull.com/q/10320914/change-ribbon-programatically.aspx
' Ref: http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Load_Ribbons_Into_The_Database:..._From_XML_File

    On Error Resume Next
    CodeDb.Properties.Append CodeDb.CreateProperty("aeRibbonID", dbText, "adaept1")
    CodeDb.Properties("aeRibbonID") = "adaept1"

End Sub