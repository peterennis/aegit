Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright (c) 2011 Peter F. Ennis
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation;
'version 3.0.
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, visit
'http://www.gnu.org/licenses/lgpl-3.0.txt

Option Compare Database
Option Explicit

' Ref: http://www.di-mgt.com.au/cl_Simple.html
'=======================================================================
' Author:   Peter F. Ennis
' Date:     February 24, 2011
' Comment:  Create class for revision control
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'=======================================================================

Private Const aegitVERSION As String = "0.3.9"
Private Const aegitVERSION_DATE As String = "Jul7 2, 2013"
Private Const THE_DRIVE As String = "C"

Private Const gcfHandleErrors As Boolean = True

' Ref: http://www.cpearson.com/excel/sizestring.htm
' This enum is used by SizeString to indicate whether the supplied text
' appears on the left or right side of result string.
Private Enum SizeStringSide
    TextLeft = 1
    TextRight = 2
End Enum

Private Type mySetupType
    SourceFolder As String
    ImportFolder As String
    UseImportFolder As Boolean
End Type

' Current pointer to the array element of the call stack
Private mintStackPointer As Integer
' Array of procedure names in the call stack
Private mastrCallStack() As String
' The number of elements to increase the array
Private Const mcintIncrementStackSize As Integer = 10
Private mfInErrorHandler As Boolean

Private aegitType As mySetupType
Private aegitSourceFolder As String
Private aegitImportFolder As String
Private aegitUseImportFolder As Boolean
Private aegitblnCustomSourceFolder As Boolean
Private aestrSourceLocation As String
Private aestrImportLocation As String
Private aeintLTN As Long
Private aeintFNLen As Long
Private aeintFTLen As Long
Private Const aeintFSize As Long = 4
Private aeintFDLen As Long
Private Const aestr4 As String = "    "
Private Const aeSqlTxtFile = "SqlCodeForQueries.txt"
Private Const aeTblTxtFile = "TblSetupForTables.txt"
Private Const aeRefTxtFile = "ReferencesSetup.txt"
Private Const aeRelTxtFile = "RelationsSetup.txt"
Private Const aePrpTxtFile = "PropertiesBuiltIn.txt"
'

Private Sub Class_Initialize()
' Ref: http://www.cadalyst.com/cad/autocad/programming-with-class-part-1-5050
' Ref: http://www.bigresource.com/Tracker/Track-vb-cyJ1aJEyKj/
' Ref: http://stackoverflow.com/questions/1731052/is-there-a-way-to-overload-the-constructor-initialize-procedure-for-a-class-in

    ' provide a default value for the SourceFolder and ImportFolder properties
    aegitSourceFolder = "default"
    aegitImportFolder = "default"
    aegitUseImportFolder = False
    aegitType.SourceFolder = "C:\ae\aegit\aerc\src\"
    aegitType.ImportFolder = "C:\ae\aegit\aerc\imp\"
    aegitType.UseImportFolder = False
    aeintLTN = LongestTableName
    LongestFieldPropsName

    Debug.Print "Class_Initialize"
    Debug.Print , "Default for aegitSourceFolder = " & aegitSourceFolder
    Debug.Print , "Default for aegitImportFolder = " & aegitImportFolder
    Debug.Print , "Default for aegitType.SourceFolder = " & aegitType.SourceFolder
    Debug.Print , "Default for aegitType.ImportFolder = " & aegitType.ImportFolder
    Debug.Print , "Default for aegitType.UseImportFolder = " & aegitType.UseImportFolder
    Debug.Print , "aeintLTN = " & aeintLTN
    Debug.Print , "aeintFNLen = " & aeintFNLen
    Debug.Print , "aeintFTLen = " & aeintFTLen
    Debug.Print , "aeintFSize = " & aeintFSize
    Debug.Print , "aeintFDLen = " & aeintFDLen
End Sub

Private Sub Class_Terminate()
    Debug.Print
    Debug.Print "Class_Terminate"
    Debug.Print , "aegit VERSION: " & aegitVERSION
    Debug.Print , "aegit VERSION_DATE: " & aegitVERSION_DATE
End Sub

Property Get SourceFolder() As String
    SourceFolder = aegitSourceFolder
End Property

Property Let SourceFolder(ByVal strSourceFolder As String)
    ' Ref: http://www.techrepublic.com/article/build-your-skills-using-class-modules-in-an-access-database-solution/5031814
    ' Ref: http://www.utteraccess.com/wiki/index.php/Classes
    aegitSourceFolder = strSourceFolder
End Property

Property Get ImportFolder() As String
    ImportFolder = aegitImportFolder
End Property

Property Let ImportFolder(ByVal strImportFolder As String)
    aegitImportFolder = strImportFolder
End Property

Property Let UseImportFolder(ByVal blnUseImportFolder As Boolean)
    aegitUseImportFolder = blnUseImportFolder
End Property

Property Get DocumentTheDatabase(Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get DocumentTheDatabase"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentTheDatabase = aeDocumentTheDatabase
    Else
        Debug.Print "Get DocumentTheDatabase"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeDocumentTheDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTheDatabase = aeDocumentTheDatabase(DebugTheCode)
    End If
End Property

Property Get Exists(strAccObjType As String, _
                        strAccObjName As String, _
                        Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get Exists"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeExists"
        Debug.Print , "DEBUGGING IS OFF"
        Exists = aeExists(strAccObjType, strAccObjName)
    Else
        Debug.Print "Get Exists"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeExists"
        Debug.Print , "DEBUGGING TURNED ON"
        Exists = aeExists(strAccObjType, strAccObjName, DebugTheCode)
    End If
End Property

Property Get ReadDocDatabase(Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get ReadDocDatabase"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeReadDocDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        ReadDocDatabase = aeReadDocDatabase
    Else
        Debug.Print "Get ReadDocDatabase"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeReadDocDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
        ReadDocDatabase = aeReadDocDatabase(DebugTheCode)
    End If
End Property

Property Get GetReferences(Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get GetReferences"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeGetReferences"
        Debug.Print , "DEBUGGING IS OFF"
        GetReferences = aeGetReferences
    Else
        Debug.Print "Get GetReferences"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeGetReferences"
        Debug.Print , "DEBUGGING TURNED ON"
        GetReferences = aeGetReferences(DebugTheCode)
    End If
End Property

Property Get DocumentRelations(Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get DocumentRelations"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeDocumentRelations"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentRelations = aeDocumentRelations
    Else
        Debug.Print "Get DocumentRelations"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeDocumentRelations"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentRelations = aeDocumentRelations(DebugTheCode)
    End If
End Property

Property Get DocumentTables(Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get DocumentTables"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeDocumentTables"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentTables = aeDocumentTables
    Else
        Debug.Print "Get DocumentTables"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeDocumentTables"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTables = aeDocumentTables(DebugTheCode)
    End If
End Property

Property Get CompactAndRepair(Optional varTrueFalse As Variant) As Boolean
' Automation for Compact and Repair

    Dim blnRun As Boolean

    Debug.Print "CompactAndRepair"
    If Not IsMissing(varTrueFalse) Then
        blnRun = False
        Debug.Print , "varTrueFalse IS NOT MISSING so blnRun of CompactAndRepair is set to False"
        Debug.Print , "RUN CompactAndRepair IS OFF"
    Else
        blnRun = True
        Debug.Print , "varTrueFalse IS MISSING so blnRun of CompactAndRepair is set to True"
        Debug.Print , "RUN CompactAndRepair IS ON..."
    End If

' TableDefs not refreshed after create
' Ref: http://support.microsoft.com/kb/104339
' So force a compact and repair
' Ref: http://msdn.microsoft.com/en-us/library/office/aa202943(v=office.10).aspx
' Not a "good practice" but for this use it is simple and works
' From the Access window
' Access 2003: SendKeys "%(TDC)", False
' Access 2007: SendKeys "%(FMC)", False
' Access 2010: SendKeys "%(YC)", False
' From the Immediate window
    
    If blnRun Then
        ' Close VBA
        SendKeys "%F{END}{ENTER}", False
        ' Run Compact and Repair
        SendKeys "%F{TAB}{TAB}{ENTER}", False
        CompactAndRepair = True
    Else
        CompactAndRepair = False
    End If
    
End Property

Private Function aeGetReferences(Optional varDebug As Variant) As Boolean
' Ref: http://vbadud.blogspot.com/2008/04/get-references-of-vba-project.html
' Ref: http://www.pcreview.co.uk/forums/type-property-reference-object-vbulletin-project-t3793816.html
' Ref: http://www.cpearson.com/excel/missingreferences.aspx
' Ref: http://allenbrowne.com/ser-38.html
' Ref: http://access.mvps.org/access/modules/mdl0022.htm (References Wizard)
' Ref: http://www.accessmvp.com/djsteele/AccessReferenceErrors.html
'====================================================================
' Author:   Peter F. Ennis
' Date:     November 28, 2012
' Comment:  Added and adapted from aeladdin (tm) code
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim i As Integer
    Dim RefName As String
    Dim RefDesc As String
    Dim blnRefBroken As Boolean
    Dim blnDebug As Boolean
    Dim strFile As String

    Dim vbaProj As Object
    Set vbaProj = Application.VBE.ActiveVBProject

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeGetReferences"

    Debug.Print "aeGetReferences"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeGetReferences is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeGetReferences is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    strFile = aestrSourceLocation & aeRefTxtFile
    
    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    If blnDebug Then
        Debug.Print ">==> aeGetReferences >==>"
        Debug.Print , "vbaProj.Name = " & vbaProj.Name
        Debug.Print , "vbaProj.Type = '" & vbaProj.Type & "'"
        ' Display the versions of Access, ADO and DAO
        Debug.Print , "Access version = " & Application.Version
        Debug.Print , "ADO (ActiveX Data Object) version = " & CurrentProject.Connection.Version
        Debug.Print , "DAO (DbEngine)  version = " & Application.DBEngine.Version
        Debug.Print , "DAO (CodeDb)    version = " & Application.CodeDb.Version
        Debug.Print , "DAO (CurrentDb) version = " & Application.CurrentDb.Version
        Debug.Print , "<@_@>"
        Debug.Print , "     " & "References:"
    End If

        Print #1, ">==> The Project References >==>"
        Print #1, , "vbaProj.Name = " & vbaProj.Name
        Print #1, , "vbaProj.Type = '" & vbaProj.Type & "'"
        ' Display the versions of Access, ADO and DAO
        Print #1, , "Access version = " & Application.Version
        Print #1, , "ADO (ActiveX Data Object) version = " & CurrentProject.Connection.Version
        Print #1, , "DAO (DbEngine)  version = " & Application.DBEngine.Version
        Print #1, , "DAO (CodeDb)    version = " & Application.CodeDb.Version
        Print #1, , "DAO (CurrentDb) version = " & Application.CurrentDb.Version
        Print #1, , "<@_@>"
        Print #1, , "     " & "References:"

    For i = 1 To vbaProj.References.Count

        blnRefBroken = False

        ' Get the Name of the Reference
        RefName = vbaProj.References(i).Name

        ' Get the Description of Reference
        RefDesc = vbaProj.References(i).Description

        If blnDebug Then Debug.Print , , vbaProj.References(i).Name, vbaProj.References(i).Description
        If blnDebug Then Debug.Print , , , vbaProj.References(i).FullPath

        Print #1, , , vbaProj.References(i).Name, vbaProj.References(i).Description
        Print #1, , , , vbaProj.References(i).FullPath

        ' Returns a Boolean value indicating whether or not the Reference object points to a valid reference in the registry. Read-only.
        If Application.VBE.ActiveVBProject.References(i).IsBroken = True Then
              blnRefBroken = True
              If blnDebug Then Debug.Print , , vbaProj.References(i).Name, "blnRefBroken=" & blnRefBroken
              Print #1, , , vbaProj.References(i).Name, "blnRefBroken=" & blnRefBroken
        End If
    Next
    If blnDebug Then Debug.Print , "<*_*>"
    If blnDebug Then Debug.Print "<==<"

    Print #1, , "<*_*>"
    Print #1, "<==<"

    aeGetReferences = True

PROC_EXIT:
    Set vbaProj = Nothing
    Close 1
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeGetReferences of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeGetReferences of Class aegitClass"
    aeGetReferences = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function LongestTableName() As Integer
'====================================================================
' Author:   Peter F. Ennis
' Date:     November 30, 2012
' Comment:  Return the length of the longest table name
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim dbs As DAO.Database
    Dim tblDef As DAO.TableDef
    Dim intTNLen As Integer

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "LongestTableName"

    intTNLen = 0
    Set dbs = CurrentDb()
    For Each tblDef In CurrentDb.TableDefs
        If Not (Left(tblDef.Name, 4) = "MSys" _
                Or Left(tblDef.Name, 4) = "~TMP" _
                Or Left(tblDef.Name, 3) = "zzz") Then
            If Len(tblDef.Name) > intTNLen Then
                intTNLen = Len(tblDef.Name)
            End If
        End If
    Next tblDef

    LongestTableName = intTNLen

PROC_EXIT:
    Set tblDef = Nothing
    Set dbs = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestTableName of Class aegitClass"
    LongestTableName = 0
    GlobalErrHandler
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
    Dim strLFN As String
    Dim strLFT As String
    Dim strLFD As String

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "LongestFieldPropsName"
    
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
                    strLFN = fld.Name
                    aeintFNLen = Len(fld.Name)
                End If
                If Len(FieldTypeName(fld)) > aeintFTLen Then
                    strLFT = FieldTypeName(fld)
                    aeintFTLen = Len(FieldTypeName(fld))
                End If
                If Len(GetDescrip(fld)) > aeintFDLen Then
                    strLFD = GetDescrip(fld)
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
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestFieldPropsName of Class aegitClass"
    LongestFieldPropsName = False
    GlobalErrHandler
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
        'Case ###         ' Add your own error management or log error to logging table
        Case Else
            'a custom log usage function commented out
            'function LogUsage(ByVal strFormName As String, strCallingProc As String, Optional ControlName) As Boolean
            'call LogUsage Err.Number, "basRelinkTables", "GetCurrentPath" ()
    End Select
    Resume PROC_EXIT
End Function

Private Function FileLocked(strFileName As String) As Boolean
' Ref: http://support.microsoft.com/kb/209189
    On Error Resume Next
    ' If the file is already opened by another process,
    ' and the specified type of access is not allowed,
    ' the Open operation fails and an error occurs.
    Open strFileName For Binary Access Read Write Lock Read Write As #1
    Close 1
    ' If an error occurs, the document is currently open.
    If Err.Number <> 0 Then
        ' Display the error number and description.
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FileLocked of Class aegitClass"
        FileLocked = True
        Err.Clear
    End If
End Function

Private Function TableInfo(strTableName As String, Optional varDebug As Variant) As Boolean
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
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "TableInfo"

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

    If aeintFDLen < Len("DESCRIPTION") Then aeintFDLen = Len("DESCRIPTION")

    If blnDebug Then
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

    Print #1, SizeString("-", sLen, TextLeft, "-")
    Print #1, SizeString("TABLE: " & strTableName, sLen, TextLeft, " ")
    Print #1, SizeString("-", sLen, TextLeft, "-")
    If strLinkedTablePath <> "" Then
        Print #1, "Linked=>" & strLinkedTablePath
    End If
    Print #1, SizeString("FIELD NAME", aeintFNLen, TextLeft, " ") _
                        & aestr4 & SizeString("FIELD TYPE", aeintFTLen, TextLeft, " ") _
                        & aestr4 & SizeString("SIZE", aeintFSize, TextLeft, " ") _
                        & aestr4 & SizeString("DESCRIPTION", aeintFDLen, TextLeft, " ")
    Print #1, SizeString("=", aeintFNLen, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFTLen, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFSize, TextLeft, "=") _
                        & aestr4 & SizeString("=", aeintFDLen, TextLeft, "=")
    strLinkedTablePath = ""

    For Each fld In tdf.Fields
        If blnDebug Then
            Debug.Print SizeString(fld.Name, aeintFNLen, TextLeft, " ") _
                & aestr4 & SizeString(FieldTypeName(fld), aeintFTLen, TextLeft, " ") _
                & aestr4 & SizeString(fld.Size, aeintFSize, TextLeft, " ") _
                & aestr4 & SizeString(GetDescrip(fld), aeintFDLen, TextLeft, " ")
        End If
        Print #1, SizeString(fld.Name, aeintFNLen, TextLeft, " ") _
            & aestr4 & SizeString(FieldTypeName(fld), aeintFTLen, TextLeft, " ") _
            & aestr4 & SizeString(fld.Size, aeintFSize, TextLeft, " ") _
            & aestr4 & SizeString(GetDescrip(fld), aeintFDLen, TextLeft, " ")
    Next
    If blnDebug Then Debug.Print
    Print #1, vbCrLf

    TableInfo = True

PROC_EXIT:
    Set fld = Nothing
    Set tdf = Nothing
    Set dbs = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TableInfo of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TableInfo of Class aegitClass"
    TableInfo = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function GetDescrip(obj As Object) As String
    On Error Resume Next
    GetDescrip = obj.Properties("Description")
End Function

Private Function FieldTypeName(fld As DAO.Field) As String
' Ref: http://allenbrowne.com/func-06.html
' Purpose: Converts the numeric results of DAO Field.Type to text
    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work
        'prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"             'dbAttachment
        Case 102&: strReturn = "Complex Byte"           'dbComplexByte
        Case 103&: strReturn = "Complex Integer"        'dbComplexInteger
        Case 104&: strReturn = "Complex Long"           'dbComplexLong
        Case 105&: strReturn = "Complex Single"         'dbComplexSingle
        Case 106&: strReturn = "Complex Double"         'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"           'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"        'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"           'dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn

End Function

Private Function aeDocumentTables(Optional varDebug As Variant) As Boolean
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
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeDocumentTables"
    
    intFailCount = 0
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

    strFile = aestrSourceLocation & aeTblTxtFile
    
    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    For Each tblDef In CurrentDb.TableDefs
        If Not (Left(tblDef.Name, 4) = "MSys" _
                Or Left(tblDef.Name, 4) = "~TMP" _
                Or Left(tblDef.Name, 3) = "zzz") Then
            If blnDebug Then
                blnResult = TableInfo(tblDef.Name, "WithDebugging")
                If Not blnResult Then intFailCount = intFailCount + 1
            Else
                blnResult = TableInfo(tblDef.Name)
                If Not blnResult Then intFailCount = intFailCount + 1
            End If
        End If
    Next tblDef

    If intFailCount > 0 Then
        aeDocumentTables = False
    Else
        aeDocumentTables = True
    End If
    If blnDebug Then
        Debug.Print "intFailCount = " & intFailCount
        Debug.Print "aeDocumentTables = " & aeDocumentTables
    End If

    aeDocumentTables = True

PROC_EXIT:
    Set fld = Nothing
    Set tblDef = Nothing
    Close 1
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTables of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTables of Class aegitClass"
    aeDocumentTables = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function isPK(tblDef As DAO.TableDef, strField As String) As Boolean
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tblDef.Indexes
        If idx.Primary Then
            For Each fld In idx.Fields
                If strField = fld.Name Then
                    isPK = True
                    Exit Function
                End If
            Next fld
        End If
    Next idx
End Function

Private Function isIndex(tblDef As DAO.TableDef, strField As String) As Boolean
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tblDef.Indexes
        For Each fld In idx.Fields
            If strField = fld.Name Then
                isIndex = True
                Exit Function
            End If
        Next fld
    Next idx
End Function

Private Function isFK(tblDef As DAO.TableDef, strField As String) As Boolean
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tblDef.Indexes
        If idx.Foreign Then
            For Each fld In idx.Fields
                If strField = fld.Name Then
                    isFK = True
                    Exit Function
                End If
            Next fld
        End If
    Next idx
End Function

Private Function aeDocumentRelations(Optional varDebug As Variant) As Boolean
' Ref: http://www.tek-tips.com/faqs.cfm?fid=6905
  
    Dim strDocument As String
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim prop As DAO.Property
    Dim strFile As String
    
    Dim blnDebug As Boolean

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeDocumentRelations"

    Debug.Print "aeDocumentRelations"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeDocumentRelations is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeDocumentRelations is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    strFile = aestrSourceLocation & aeRelTxtFile
    
    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    For Each rel In CurrentDb.Relations
        If Not (Left(rel.Name, 4) = "MSys" _
                        Or Left(rel.Name, 4) = "~TMP" _
                        Or Left(rel.Name, 3) = "zzz") Then
            strDocument = strDocument & vbCrLf & "Name: " & rel.Name & vbCrLf
            strDocument = strDocument & "  " & "Table: " & rel.Table & vbCrLf
            strDocument = strDocument & "  " & "Foreign Table: " & rel.ForeignTable & vbCrLf
            For Each fld In rel.Fields
                strDocument = strDocument & "  PK: " & fld.Name & "   FK:" & fld.ForeignName
                strDocument = strDocument & vbCrLf
            Next fld
        End If
    Next rel
    If blnDebug Then Debug.Print strDocument
    Print #1, strDocument
    
    aeDocumentRelations = True

PROC_EXIT:
    Set prop = Nothing
    Set idx = Nothing
    Set fld = Nothing
    Set rel = Nothing
    Close 1
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentRelations of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentRelations of Class aegitClass"
    aeDocumentRelations = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function OutputQueriesSqlText() As Boolean
' Ref: http://www.pcreview.co.uk/forums/export-sql-saved-query-into-text-file-t2775525.html
'====================================================================
' Author:   Peter F. Ennis
' Date:     December 3, 2012
' Comment:  Output the sql code of all queries to a text file
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim dbs As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strFile As String

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "OutputQueriesSqlText"

    strFile = aestrSourceLocation & aeSqlTxtFile
    
    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    Set dbs = CurrentDb
    For Each qdf In dbs.QueryDefs
        If Not (Left(qdf.Name, 4) = "MSys" Or Left(qdf.Name, 4) = "~sq_" _
                        Or Left(qdf.Name, 4) = "~TMP" _
                        Or Left(qdf.Name, 3) = "zzz") Then
            Print #1, "<<<" & qdf.Name & ">>>" & vbCrLf & qdf.SQL
        End If
    Next

    OutputQueriesSqlText = True

PROC_EXIT:
    Set qdf = Nothing
    Set dbs = Nothing
    Close 1
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputQueriesSqlText of Class aegitClass"
    OutputQueriesSqlText = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Sub KillProperly(Killfile As String)
' Ref: http://word.mvps.org/faqs/macrosvba/DeleteFiles.htm
    If Len(Dir(Killfile)) > 0 Then
        SetAttr Killfile, vbNormal
        Kill Killfile
    End If
End Sub

Private Function GetPropEnum(typeNum As Long) As String
' Ref: http://msdn.microsoft.com/en-us/library/bb242635.aspx
 
    Select Case typeNum
        Case 1
            GetPropEnum = "dbBoolean"
        Case 2
            GetPropEnum = "dbByte"
        Case 3
            GetPropEnum = "dbInteger"
        Case 4
            GetPropEnum = "dbLong"
        Case 5
            GetPropEnum = "dbCurrency"
        Case 6
            GetPropEnum = "dbSingle"
        Case 7
            GetPropEnum = "dbDouble"
        Case 8
            GetPropEnum = "dbDate"
        Case 9
            GetPropEnum = "dbBinary"
        Case 10
            GetPropEnum = "dbText"
        Case 11
            GetPropEnum = "dbLongBinary"
        Case 12
            GetPropEnum = "dbMemo"
        Case 15
            GetPropEnum = "dbGUID"
        Case 16
            GetPropEnum = "dbBigInt"
        Case 17
            GetPropEnum = "dbVarBinary"
        Case 18
            GetPropEnum = "dbChar"
        Case 19
            GetPropEnum = "dbNumeric"
        Case 20
            GetPropEnum = "dbDecimal"
        Case 21
            GetPropEnum = "dbFloat"
        Case 22
            GetPropEnum = "dbTime"
        Case 23
            GetPropEnum = "dbTimeStamp"
        Case 101
            GetPropEnum = "dbAttachment"
        Case 102
            GetPropEnum = "dbComplexByte"
        Case 103
            GetPropEnum = "dbComplexInteger"
        Case 104
            GetPropEnum = "dbComplexLong"
        Case 105
            GetPropEnum = "dbComplexSingle"
        Case 106
            GetPropEnum = "dbComplexDouble"
        Case 107
            GetPropEnum = "dbComplexGUID"
        Case 108
            GetPropEnum = "dbComplexDecimal"
        Case 109
            GetPropEnum = "dbComplexText"
    End Select

End Function

Private Function GetPrpValue(obj As Object) As String
    'On Error Resume Next
    GetPrpValue = obj.Properties("Value")
End Function
 
Private Function OutputBuiltInPropertiesText() As Boolean
' Ref: http://www.jpsoftwaretech.com/listing-built-in-access-database-properties/

    Dim dbs As DAO.Database
    Dim prps As DAO.Properties
    Dim prp As DAO.Property
    Dim varPropValue As Variant
    Dim strFile As String
    Dim strError As String

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "OutputBuiltInPropertiesText"

    strFile = aestrSourceLocation & aePrpTxtFile

    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If
 
    Set dbs = CurrentDb
    Set prps = dbs.Properties

    Debug.Print "OutputBuiltInPropertiesText"

    For Each prp In prps
        Print #1, "Name: " & prp.Name
        Print #1, "Type: " & GetPropEnum(prp.Type)
        ' Fixed for error 3251
        varPropValue = GetPrpValue(prp)
        Print #1, "Value: " & varPropValue
        Print #1, "Inherited: " & prp.Inherited & ";" & strError
        strError = ""
        Print #1, "---"
    Next prp

    OutputBuiltInPropertiesText = True

PROC_EXIT:
    Set prp = Nothing
    Set prps = Nothing
    Set dbs = Nothing
    Close 1
    PopCallStack
    Exit Function

PROC_ERR:
     Select Case Err.Number
        Case 3251
            strError = Err.Number & "," & Err.Description
            varPropValue = Null
            Resume Next
        Case Else
            'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputBuiltInPropertiesText of Class aegitClass"
            'If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputBuiltInPropertiesText of Class aegitClass"
            OutputBuiltInPropertiesText = False
            GlobalErrHandler
            Resume PROC_EXIT
    End Select

End Function
 
Private Function DocumentTheContainer(strContainerType As String, strExt As String, Optional varDebug As Variant) As Boolean
' strContainerType: Forms, Reports, Scripts (Macros), Modules

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "DocumentTheContainer"

    Dim dbs As DAO.Database
    Dim cnt As DAO.Container
    Dim doc As DAO.Document
    Dim i As Integer
    Dim intAcObjType As Integer
    Dim blnDebug As Boolean

    Set dbs = CurrentDb() ' use CurrentDb() to refresh Collections

    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of DocumentTheContainer is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of DocumentTheContainer is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    i = 0
    Set cnt = dbs.Containers(strContainerType)

    Select Case strContainerType
        Case "Forms": intAcObjType = 2   'acForm
        Case "Reports": intAcObjType = 3 'acReport
        Case "Scripts": intAcObjType = 4 'acMacro
        Case "Modules": intAcObjType = 5 'acModule
        Case Else
            MsgBox "Wrong Case Select in DocumentTheContainer"
    End Select

    If blnDebug Then Debug.Print UCase(strContainerType)

    For Each doc In cnt.Documents
        If blnDebug Then Debug.Print , doc.Name
        If Not (Left(doc.Name, 3) = "zzz" Or Left(doc.Name, 4) = "~TMP") Then
            i = i + 1
            ' NOTE: Err 2220 is intermittent. Seems to happen more after compact and repair.
            ' If some code is added it goes away so this break is included for detection to
            ' hopefully find some solution...
            If Err.Number = 2220 Then Stop
            Application.SaveAsText intAcObjType, doc.Name, aestrSourceLocation & doc.Name & "." & strExt
        End If
    Next doc

    If blnDebug Then
        Debug.Print , i & " EXPORTED!"
        Debug.Print , cnt.Documents.Count & " EXISTING!"
    End If

    DocumentTheContainer = True

PROC_EXIT:
    Set doc = Nothing
    Set cnt = Nothing
    Set dbs = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    DocumentTheContainer = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function
 
Private Function aeDocumentTheDatabase(Optional varDebug As Variant) As Boolean
' Based on sample code from Arvin Meyer (MVP) June 2, 1999
' Ref: http://www.accessmvp.com/Arvin/DocDatabase.txt
' Ref: http://www.tek-tips.com/faqs.cfm?fid=6905
'====================================================================
' Author:   Peter F. Ennis
' Date:     February 8, 2011
' Comment:  Uses the undocumented [Application.SaveAsText] syntax
'           To reload use the syntax [Application.LoadFromText]
'           Add explicit references for DAO
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim dbs As DAO.Database
    Dim cnt As DAO.Container
    Dim doc As DAO.Document
    Dim qdf As DAO.QueryDef
    Dim strFile As String
    Dim i As Integer
    Dim blnDebug As Boolean

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeDocumentTheDatabase"

    Debug.Print "aeDocumentTheDatabase"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeDocumentTheDatabase is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeDocumentTheDatabase is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If
    
    If aegitSourceFolder = "default" Then
        aestrSourceLocation = aegitType.SourceFolder
    Else
        aestrSourceLocation = aegitSourceFolder
    End If

    If aegitImportFolder = "default" Then
            aestrImportLocation = aegitType.ImportFolder
    End If
    If aegitUseImportFolder Then
        aestrImportLocation = aegitImportFolder
    End If
 
    ' Delete all the files in a given directory:
    ' Loop through all the files in the directory by using Dir$ function
    strFile = Dir(aestrSourceLocation & "*.*")
    Do While strFile <> ""
        KillProperly aestrSourceLocation & strFile
        ' Need to specify full path again because a file was deleted
        strFile = Dir(aestrSourceLocation & "*.*")
    Loop

    If blnDebug Then
        Debug.Print , ">==> aeDocumentTheDatabase >==>"
        Debug.Print , "SourceFolder = " & aestrSourceLocation
        Debug.Print , "ImportFolder = " & aestrImportLocation
    End If
    
    If blnDebug Then
        DocumentTheContainer "Forms", "frm", "WithDebugging"
        DocumentTheContainer "Reports", "rpt", "WithDebugging"
        DocumentTheContainer "Scripts", "mac", "WithDebugging"
        DocumentTheContainer "Modules", "bas", "WithDebugging"
    Else
        DocumentTheContainer "Forms", "frm"
        DocumentTheContainer "Reports", "rpt"
        DocumentTheContainer "Scripts", "mac"
        DocumentTheContainer "Modules", "bas"
    End If
    
    ListContainers ("ListOfContainers.txt")
    'Stop

    Set dbs = CurrentDb() ' use CurrentDb() to refresh Collections

    '=============
    ' QUERIES
    '=============
    i = 0
    If blnDebug Then Debug.Print "QUERIES"

    ' Delete all TEMP queries ...
    For Each qdf In CurrentDb.QueryDefs
        If Left(qdf.Name, 1) = "~" Then
            CurrentDb.QueryDefs.Delete qdf.Name
            CurrentDb.QueryDefs.Refresh
        End If
    Next qdf

    For Each qdf In CurrentDb.QueryDefs
        If blnDebug Then Debug.Print , qdf.Name
        If Not (Left(qdf.Name, 4) = "MSys" Or Left(qdf.Name, 4) = "~sq_" _
                        Or Left(qdf.Name, 4) = "~TMP" _
                        Or Left(qdf.Name, 3) = "zzz") Then
            i = i + 1
            Application.SaveAsText acQuery, qdf.Name, aestrSourceLocation & qdf.Name & ".qry"
        End If
    Next qdf
    
    If blnDebug Then
        If i = 1 Then
            Debug.Print , "1 Query EXPORTED!"
        Else
            Debug.Print , i & " Queries EXPORTED!"
        End If
        
        If CurrentDb.QueryDefs.Count = 1 Then
            Debug.Print , "1 Query EXISTING!"
        Else
            Debug.Print , CurrentDb.QueryDefs.Count & " Queries EXISTING!"
        End If
    End If

    OutputQueriesSqlText
    OutputBuiltInPropertiesText
    
    aeDocumentTheDatabase = True
    
PROC_EXIT:
    Set qdf = Nothing
    Set doc = Nothing
    Set cnt = Nothing
    Set dbs = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTheDatabase of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTheDatabase of Class aegitClass"
    aeDocumentTheDatabase = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function BuildTheDirectory(FSO As Scripting.FileSystemObject, _
                                        Optional varDebug As Variant) As Boolean
' Ref: http://msdn.microsoft.com/en-us/library/ebkhfaaz(v=vs.85).aspx
'====================================================================
' Author:   Peter F. Ennis
' Date:     February 8, 2011
' Comment:  Add optional debug parameter
' Requires: Reference to Microsoft Scripting Runtime
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim objImportFolder As Object
    Dim blnDebug As Boolean

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "BuildTheDirectory"

    Debug.Print "BuildTheDirectory"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of BuildTheDirectory is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of BuildTheDirectory is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    If blnDebug Then Debug.Print , ">==> BuildTheDirectory >==>"

    ' Bail out if (a) the drive does not exist, or if (b) the directory already exists.

    If blnDebug Then Debug.Print , , "THE_DRIVE = " & THE_DRIVE
    If blnDebug Then Debug.Print , , "FSO.DriveExists(THE_DRIVE) = " & FSO.DriveExists(THE_DRIVE)
    If Not FSO.DriveExists(THE_DRIVE) Then
        If blnDebug Then Debug.Print , , "FSO.DriveExists(THE_DRIVE) = FALSE - The drive DOES NOT EXIST !!!"
        BuildTheDirectory = False
        Exit Function
    End If
    If blnDebug Then Debug.Print , , "The drive EXISTS !!!"
    If blnDebug Then Debug.Print , , "aegitUseImportFolder = " & aegitUseImportFolder
    
    If aegitImportFolder = "default" Then
        aestrImportLocation = aegitType.ImportFolder
    End If
    If aegitUseImportFolder And aegitImportFolder <> "default" Then
        aestrImportLocation = aegitImportFolder
    End If
        
    If blnDebug Then Debug.Print , , "The import directory is: " & aestrImportLocation
   
    If FSO.FolderExists(aestrImportLocation) Then
        If blnDebug Then Debug.Print , , "FSO.FolderExists(aestrImportLocation) = TRUE - The directory EXISTS !!!"
        BuildTheDirectory = False
        Exit Function
    End If
    If blnDebug Then Debug.Print , , "The import directory does NOT EXIST !!!"

    If aegitUseImportFolder Then
        Set objImportFolder = FSO.CreateFolder(aestrImportLocation)
        If blnDebug Then Debug.Print , , aestrImportLocation & " has been CREATED !!!"
    End If

    BuildTheDirectory = True

PROC_EXIT:
    Set objImportFolder = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure BuildTheDirectory of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure BuildTheDirectory of Class aegitClass"
    BuildTheDirectory = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function aeReadDocDatabase(Optional varDebug As Variant) As Boolean
' VBScript makes use of ADOX (Microsoft's Active Data Objects Extensions for Data Definition Language and Security)
' to create a query on a Microsoft Access database
' Ref: http://stackoverflow.com/questions/859530/alternative-to-application-loadfromtext-for-ms-access-queries
' Microsoft Access Stored Queries and VBscript: How to Create and Edit a Stored Database Query
' Ref: http://www.suite101.com/content/microsoft-access-stored-queries-and-vbscript-a87978#ixzz1D32Vqbso
' Using WScript
' Ref: http://www.codeforexcelandoutlook.com/vba/shell-scripting-using-vba-and-the-windows-script-host-object-model/
' vbscript get file extension
' Ref: http://www.experts-exchange.com/Programming/Languages/Visual_Basic/Q_24297896.html
'
'====================================================================
' Author:   Peter F. Ennis
' Date:     February 8, 2011
' Comment:  Add explicit references for objects, wscript, fso
' Requires: Reference to Microsoft Scripting Runtime
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim MyFile As Object
    Dim strFileType As String
    Dim strFileBaseName As String
    Dim bln As Boolean
    Dim blnDebug As Boolean

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeReadDocDatabase"

    Debug.Print "aeReadDocDatabase"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeReadDocDatabase is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeReadDocDatabase is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    Const acQuery = 1

    If aegitSourceFolder = "default" Then
        aestrSourceLocation = aegitType.SourceFolder
    Else
        aestrSourceLocation = aegitSourceFolder
    End If

    If aegitImportFolder = "default" Then
        aestrImportLocation = aegitType.ImportFolder
    End If
    If aegitUseImportFolder And aegitImportFolder <> "default" Then
        aestrImportLocation = aegitImportFolder
    End If

    If blnDebug Then
        Debug.Print ">==> aeReadDocDatabase >==>"
        Debug.Print , "aegit VERSION: " & aegitVERSION
        Debug.Print , "aegit VERSION_DATE: " & aegitVERSION_DATE
        Debug.Print , "SourceFolder = " & aestrSourceLocation
        Debug.Print , "UseImportFolder = " & aegitUseImportFolder
        Debug.Print , "ImportFolder = " & aestrImportLocation
        'Stop
    End If

    ' Create needed objects
    Dim wsh As Object  ' As Object if late-bound
    Set wsh = CreateObject("WScript.Shell")

    wsh.CurrentDirectory = aestrImportLocation
    If blnDebug Then Debug.Print , "wsh.CurrentDirectory = " & wsh.CurrentDirectory
    ' CurDir Function
    If blnDebug Then Debug.Print , "CurDir = " & CurDir
    
    Dim FSO As Scripting.FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

    If blnDebug Then
        bln = BuildTheDirectory(FSO, "WithDebugging")
        Debug.Print , "<==<"
    Else
        bln = BuildTheDirectory(FSO)
    End If

    If aegitUseImportFolder Then
        Dim objFolder As Object
        Set objFolder = FSO.GetFolder(aegitType.ImportFolder)

        For Each MyFile In objFolder.Files
            If blnDebug Then Debug.Print "myFile = " & MyFile
            If blnDebug Then Debug.Print "myFile.Name = " & MyFile.Name
            strFileBaseName = FSO.GetBaseName(MyFile.Name)
            strFileType = FSO.GetExtensionName(MyFile.Name)
            If blnDebug Then Debug.Print strFileBaseName & " (" & strFileType & ")"

            If (strFileType = "frm") Then
                If Exists("FORMS", strFileBaseName) Then
                    MsgBox "Skipping: FORM " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                    If blnDebug Then Debug.Print "Skipping: FORM " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
                Else
                    Application.LoadFromText acForm, strFileBaseName, MyFile.Path
                End If
            ElseIf (strFileType = "rpt") Then
                If Exists("REPORTS", strFileBaseName) Then
                    MsgBox "Skipping: REPORT " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                    If blnDebug Then Debug.Print "Skipping: REPORT " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
                Else
                    Application.LoadFromText acReport, strFileBaseName, MyFile.Path
                End If
            ElseIf (strFileType = "bas") Then
                If Exists("MODULES", strFileBaseName) Then
                    MsgBox "Skipping: MODULE " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                    If blnDebug Then Debug.Print "Skipping: MODULE " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
                Else
                    Application.LoadFromText acModule, strFileBaseName, MyFile.Path
                End If
            ElseIf (strFileType = "mac") Then
                If Exists("MACROS", strFileBaseName) Then
                    MsgBox "Skipping: MACRO " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                    If blnDebug Then Debug.Print "Skipping: MACRO " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
                Else
                    Application.LoadFromText acMacro, strFileBaseName, MyFile.Path
                End If
            ElseIf (strFileType = "qry") Then
                If Exists("QUERIES", strFileBaseName) Then
                    MsgBox "Skipping: QUERY " & strFileBaseName & " exists in the current database.", vbInformation, "EXISTENCE IS REAL !!!"
                    If blnDebug Then Debug.Print "Skipping: QUERY " & strFileBaseName & " exists in the current database.", "EXISTENCE IS REAL !!!"
                Else
                    Application.LoadFromText acQuery, strFileBaseName, MyFile.Path
                End If
            End If
        Next
    End If

    If blnDebug Then Debug.Print "<==<"
    
    aeReadDocDatabase = True

PROC_EXIT:
    Set MyFile = Nothing
    'Set ojbFolder = Nothing
    Set FSO = Nothing
    Set wsh = Nothing
    aeReadDocDatabase = True
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadDocDatabase of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadDocDatabase of Class aegitClass"
    aeReadDocDatabase = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function aeExists(strAccObjType As String, _
                        strAccObjName As String, Optional varDebug As Variant) As Boolean
' Ref: http://vbabuff.blogspot.com/2010/03/does-access-object-exists.html
'
'====================================================================
' Author:     Peter F. Ennis
' Date:       February 18, 2011
' Comment:    Return True if the object exists
' Parameters:
'             strAccObjType: "Tables", "Queries", "Forms",
'                            "Reports", "Macros", "Modules"
'             strAccObjName: The name of the object
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'====================================================================

    Dim objType As Object
    Dim obj As Variant
    Dim blnDebug As Boolean
    
    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeExists"

    Debug.Print "aeExists"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of aeExists is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeExists is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    If blnDebug Then Debug.Print ">==> aeExists >==>"

    Select Case strAccObjType
        Case "Tables"
            Set objType = CurrentDb.TableDefs
        Case "Queries"
            Set objType = CurrentDb.QueryDefs
        Case "Forms"
            Set objType = CurrentProject.AllForms
        Case "Reports"
            Set objType = CurrentProject.AllReports
        Case "Macros"
            Set objType = CurrentProject.AllMacros
        Case "Modules"
            Set objType = CurrentProject.AllModules
        Case Else
            MsgBox "Wrong option!", vbCritical, "in procedure aeExists of Class aegitClass"
            If blnDebug Then
                Debug.Print , "strAccObjType = >" & strAccObjType & "< is  a false value"
                Debug.Print , "Option allowed is one of 'Tables', 'Queries', 'Forms', 'Reports', 'Macros', 'Modules'"
                Debug.Print "<==<"
            End If
            aeExists = False
            Set obj = Nothing
            Exit Function
    End Select

    If blnDebug Then Debug.Print , "strAccObjType = " & strAccObjType
    If blnDebug Then Debug.Print , "strAccObjName = " & strAccObjName

    For Each obj In objType
        If blnDebug Then Debug.Print , obj.Name, strAccObjName
        If obj.Name = strAccObjName Then
            If blnDebug Then
                Debug.Print , strAccObjName & " EXISTS!"
                Debug.Print "<==<"
            End If
            aeExists = True
            Set obj = Nothing
            Exit Function ' Found it!
        Else
            aeExists = False
        End If
    Next
    If blnDebug Then
        Debug.Print , strAccObjName & " DOES NOT EXIST!"
        Debug.Print "<==<"
    End If

    aeExists = True

PROC_EXIT:
    Set obj = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class aegitClass"
    aeExists = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function

'==================================================
' Global Error Handler Routines
' Ref: http://msdn.microsoft.com/en-us/library/office/ee358847(v=office.12).aspx#odc_ac2007_ta_ErrorHandlingAndDebuggingTipsForAccessVBAndVBA_WritingCodeForDebugging
'==================================================

Private Sub ResetWorkspace()
    Dim intCounter As Integer

    On Error Resume Next

    Application.MenuBar = ""
    DoCmd.SetWarnings False
    DoCmd.Hourglass False
    DoCmd.Echo True

    ' Clean up workspace by closing open forms and reports
    For intCounter = 0 To Forms.Count - 1
        DoCmd.Close acForm, Forms(intCounter).Name
    Next intCounter

    For intCounter = 0 To Reports.Count - 1
        DoCmd.Close acReport, Reports(intCounter).Name
    Next intCounter
End Sub

Private Sub GlobalErrHandler()
' Main procedure to handle errors that occur.

    Dim strError As String
    Dim lngError As Long
    Dim intErl As Integer
    Dim strMsg As String

    ' Variables to preserve error information
    strError = Err.Description
    lngError = Err.Number
    intErl = Erl

    ' Reset workspace, close open objects
    ResetWorkspace

    ' Prompt the user with information on the error:
    strMsg = "Procedure: " & CurrentProcName() & vbCrLf & _
             "Line: " & intErl & vbCrLf & _
             "Error: (" & lngError & ")" & strError & vbCrLf & _
             "Application Quit is turned OFF !!!"
    MsgBox strMsg, vbCritical, "GlobalErrHandler"

    ' Write error to file:
    WriteErrorToFile intErl, lngError, CurrentProcName(), strError

    ' Exit Access without saving any changes
    ' (you might want to change this to save all changes)

    'Application.Quit acExit
End Sub

Private Function CurrentProcName() As String
    CurrentProcName = mastrCallStack(mintStackPointer - 1)
End Function

Private Sub PushCallStack(strProcName As String)
' Add the current procedure name to the Call Stack.
' Should be called whenever a procedure is called

    On Error Resume Next

    ' Verify the stack array can handle the current array element
    If mintStackPointer > UBound(mastrCallStack) Then
    ' If array has not been defined, initialize the error handler
        If Err.Number = 9 Then
            ErrorHandlerInit
        Else
            ' Increase the size of the array to not go out of bounds
            ReDim Preserve mastrCallStack(UBound(mastrCallStack) + _
            mcintIncrementStackSize)
        End If
    End If

    On Error GoTo 0

    mastrCallStack(mintStackPointer) = strProcName

    ' Increment pointer to next element
    mintStackPointer = mintStackPointer + 1
End Sub

Private Sub ErrorHandlerInit()
    mfInErrorHandler = False
    mintStackPointer = 1
    ReDim mastrCallStack(1 To mcintIncrementStackSize)
End Sub

Private Sub PopCallStack()
' Remove a procedure name from the call stack

    If mintStackPointer <= UBound(mastrCallStack) Then
        mastrCallStack(mintStackPointer) = ""
    End If

    ' Reset pointer to previous element
    mintStackPointer = mintStackPointer - 1
End Sub

Private Sub WriteErrorToFile(intTheErl As Integer, lngTheErrorNum As Long, _
                strCurrentProcName As String, strErrorDescription As String)
    
    Dim strFilePath As String
    Dim lngFileNum As Long
    
    On Error Resume Next

    ' Write to a text file called aegitErrorLog in the MyDocuments folder
    strFilePath = CreateObject("WScript.Shell").SpecialFolders("MYDOCUMENTS") & "\aegitErrorLog.txt"

    lngFileNum = FreeFile
    Open strFilePath For Append Access Write Lock Write As lngFileNum
        Print #lngFileNum, Now(), intTheErl, lngTheErrorNum, strCurrentProcName, strErrorDescription
    Close lngFileNum

End Sub

Private Sub WriteStringToFile(lngFileNum As Long, strTheString As String, strTheAbsoluteFileName As String)
  
    On Error Resume Next

    Open strTheAbsoluteFileName For Append Access Write Lock Write As lngFileNum
        Print #lngFileNum, strTheString
    Close lngFileNum

End Sub

Public Function ListContainers(strTheFileName As String) As Boolean
' Ref: http://www.susandoreydesigns.com/software/AccessVBATechniques.pdf
' Ref: http://msdn.microsoft.com/en-us/library/office/bb177484(v=office.12).aspx

    Dim dbs As DAO.Database
    Dim conItem As DAO.Container
    Dim prpLoop As DAO.Property
    Dim strName As String
    Dim strOwner As String
    Dim strText As String
    Dim strFile As String
    Dim lngFileNum As Long

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "ListContainers"

    Set dbs = CurrentDb
    lngFileNum = FreeFile

    strFile = aestrSourceLocation & strTheFileName

    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As lngFileNum
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As lngFileNum
    End If

    With dbs
        ' Enumerate Containers collection.
        For Each conItem In .Containers
            Debug.Print "Properties of " & conItem.Name _
                & " container"
            WriteStringToFile lngFileNum, "Properties of " & conItem.Name _
                & " container", strFile
            
            ' Enumerate Properties collection of each Container object.
            For Each prpLoop In conItem.Properties
                Debug.Print "  " & prpLoop.Name _
                    & " = "; prpLoop
                WriteStringToFile lngFileNum, "  " & prpLoop.Name _
                    & " = " & prpLoop, strFile
            Next prpLoop
        Next conItem
        .Close
    End With

    ListContainers = True

PROC_EXIT:
    Set prpLoop = Nothing
    Set conItem = Nothing
    Set dbs = Nothing
    Close lngFileNum
    PopCallStack
    Exit Function

PROC_ERR:
    ListContainers = False
    GlobalErrHandler
    Resume PROC_EXIT

End Function