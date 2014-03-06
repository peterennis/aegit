Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

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
'
' Ref: http://www.di-mgt.com.au/cl_Simple.html
'=======================================================================
' Author:   Peter F. Ennis
' Date:     February 24, 2011
' Comment:  Create class for revision control
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'=======================================================================

Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Private Const aegit_impVERSION As String = "0.8.5"
Private Const aegit_impVERSION_DATE As String = "February 5, 2014"
Private Const THE_DRIVE As String = "C"

Private Const gcfHandleErrors As Boolean = True
Private Const gblnOutputPrinterInfo = False

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
    XMLFolder As String
End Type

' Current pointer to the array element of the call stack
Private mintStackPointer As Integer
' Array of procedure names in the call stack
Private mastrCallStack() As String
' The number of elements to increase the array
Private Const mcintIncrementStackSize As Integer = 10
Private mfInErrorHandler As Boolean

Private aegitSetup As Boolean
Private aegitType As mySetupType
Private aegitSourceFolder As String
Private aegitImportFolder As String
Private aegitXMLFolder As String
Private aegitDataXML() As Variant
Private aegitExportDataToXML As Boolean
'''x Private aegitUseImportFolder As Boolean
'''x Private aestrSourceLocation As String
Private aestrImportLocation As String
'''x Private aestrXMLLocation As String
Private aeintLTN As Long                        ' Longest Table Name
Private aestrLFN As String                      ' Longest Field Name
Private aestrLFNTN As String
Private aeintFNLen As Long
Private aestrLFT As String                      ' Longest Field Type
Private aeintFTLen As Long                      ' Field Type Length
Private Const aeintFSize As Long = 4
Private aeintFDLen As Long
Private aestrLFD As String
Private Const aestr4 As String = "    "
'''x Private Const aeSqlTxtFile = "OutputSqlCodeForQueries.txt"
'''x Private Const aeTblTxtFile = "OutputTblSetupForTables.txt"
'''x Private Const aeRefTxtFile = "OutputReferencesSetup.txt"
'''x Private Const aeRelTxtFile = "OutputRelationsSetup.txt"
'''x Private Const aePrpTxtFile = "OutputPropertiesBuiltIn.txt"
'''x Private Const aeFLkCtrFile = "OutputFieldLookupControlTypeList.txt"
'''x Private Const aeSchemaFile = "OutputSchemaFile.txt"
'''x Private Const aePrnterInfo = "OutputPrinterInfo.txt"
'

Private Sub Class_Initialize()
' Ref: http://www.cadalyst.com/cad/autocad/programming-with-class-part-1-5050
' Ref: http://www.bigresource.com/Tracker/Track-vb-cyJ1aJEyKj/
' Ref: http://stackoverflow.com/questions/1731052/is-there-a-way-to-overload-the-constructor-initialize-procedure-for-a-class-in

    ' provide a default value for the SourceFolder, ImportFolder and other properties
    aegitSourceFolder = "default"
    aegitImportFolder = "default"
    aegitXMLFolder = "default"
    ReDim Preserve aegitDataXML(1 To 1)
    aegitDataXML(1) = "aetlkpStates"
'''x    aegitUseImportFolder = False
    aegitExportDataToXML = False
    aegitType.SourceFolder = "C:\ae\aegit\aerc\src\"
    aegitType.ImportFolder = "C:\ae\aegit\aerc\src\imp\"
    aegitType.UseImportFolder = False
    aegitType.XMLFolder = "C:\ae\aegit\aerc\src\xml\"
    aeintLTN = LongestTableName
    LongestFieldPropsName

    Debug.Print "Class_Initialize"
    Debug.Print , "Default for aegitSourceFolder = " & aegitSourceFolder
    Debug.Print , "Default for aegitImportFolder = " & aegitImportFolder
    Debug.Print , "Default for aegitType.SourceFolder = " & aegitType.SourceFolder
    Debug.Print , "Default for aegitType.ImportFolder = " & aegitType.ImportFolder
    Debug.Print , "Default for aegitType.UseImportFolder = " & aegitType.UseImportFolder
    Debug.Print , "Default for aegitType.XMLFolder = " & aegitType.XMLFolder
    Debug.Print , "aeintLTN = " & aeintLTN
    Debug.Print , "aeintFNLen = " & aeintFNLen
    Debug.Print , "aeintFTLen = " & aeintFTLen
    Debug.Print , "aeintFSize = " & aeintFSize
    'Debug.Print , "aeintFDLen = " & aeintFDLen

End Sub

Private Sub Class_Terminate()
    Dim strFile As String
    strFile = aegitSourceFolder & "export.ini"
    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
    End If
    Debug.Print
    Debug.Print "Class_Terminate"
    Debug.Print , "aegit VERSION: " & aegit_impVERSION
    Debug.Print , "aegit VERSION_DATE: " & aegit_impVERSION_DATE
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

'''x Property Let UseImportFolder(ByVal blnUseImportFolder As Boolean)
'''x     aegitUseImportFolder = blnUseImportFolder
'''x End Property

Property Get XMLFolder() As String
    XMLFolder = aegitXMLFolder
End Property

Property Let XMLFolder(ByVal strXMLFolder As String)
    aegitXMLFolder = strXMLFolder
End Property

Property Let TablesExportToXML(ByRef avarTables() As Variant)
    MsgBox "Let TablesExportToXML: LBound(aegitDataXML())=" & LBound(aegitDataXML()) & _
        vbCrLf & "UBound(aegitDataXML())=" & UBound(aegitDataXML()), vbInformation, "CHECK"
    'aegitDataXML = avarTables
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

Property Get ReadDocDatabase(blnImport As Boolean, Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get ReadDocDatabase"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to aeReadDocDatabase"
        Debug.Print , "DEBUGGING IS OFF"
        ReadDocDatabase = aeReadDocDatabase(blnImport)
    Else
        Debug.Print "Get ReadDocDatabase"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to aeReadDocDatabase"
        Debug.Print , "DEBUGGING TURNED ON"
        ReadDocDatabase = aeReadDocDatabase(blnImport, DebugTheCode)
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

Private Function aeReadWriteStream(strPathFileName As String) As Boolean

    'Debug.Print "aeReadWriteStream Entry strPathFileName=" & strPathFileName

    ' Use a call stack and global error handler
    'If gcfHandleErrors Then On Error GoTo PROC_ERR
    'PushCallStack "aeReadWriteStream"

    On Error GoTo PROC_ERR

    Dim fName As String
    Dim fname2 As String
    Dim fnr As Integer
    Dim fnr2 As Integer
    Dim tstring As String * 1

    aeReadWriteStream = False

    ' If the file has no Byte Order Mark (BOM)
    ' Ref: http://msdn.microsoft.com/en-us/library/windows/desktop/dd374101%28v=vs.85%29.aspx
    ' then do nothing
    fName = strPathFileName
    fname2 = strPathFileName & ".clean.txt"

    fnr = FreeFile()
    Open fName For Binary Access Read As #fnr
    Get #fnr, , tstring
    ' #FFFE, #FFFF, #0000
    ' If no BOM then it is a txt file and header stripping is not needed
    If Asc(tstring) <> 254 And Asc(tstring) <> 255 And _
                Asc(tstring) <> 0 Then
        Close #fnr
        aeReadWriteStream = False
        Exit Function
    End If

    fnr2 = FreeFile()
    Open fname2 For Binary Lock Read Write As #fnr2

    While Not EOF(fnr)
        Get #fnr, , tstring
        If Asc(tstring) = 254 Or Asc(tstring) = 255 Or _
                Asc(tstring) = 0 Then
        Else
            Put #fnr2, , tstring
        End If
    Wend

PROC_EXIT:
    Close #fnr
    Close #fnr2
    aeReadWriteStream = True
    'PopCallStack
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 9
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadWriteStream of Class aegitClass" & _
                    vbCrLf & "aeReadWriteStream Entry strPathFileName=" & strPathFileName, vbCritical, "aeReadWriteStream ERROR=9"
            'If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadWriteStream of Class aegitClass"
            'GlobalErrHandler
            Resume Next
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeReadWriteStream of Class aegitClass"
            'GlobalErrHandler
            Resume Next
    End Select

End Function

Private Function Pause(NumberOfSeconds As Variant)
' Ref: http://www.access-programmers.co.uk/forums/showthread.php?p=952355

    On Error GoTo PROC_ERR

    Dim PauseTime As Variant
    Dim Start As Variant

    PauseTime = NumberOfSeconds
    Start = Timer
    Do While Timer < Start + PauseTime
        Sleep 100
        DoEvents
    Loop

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Pause of Class aegitClass"
    Resume PROC_EXIT

End Function

Private Sub WaitSeconds(intSeconds As Integer)
' Comments: Waits for a specified number of seconds
' Params  : intSeconds      Number of seconds to wait
' Source  : Total Visual SourceBook
' Ref     : http://www.fmsinc.com/MicrosoftAccess/modules/examples/AvoidDoEvents.asp

    On Error GoTo PROC_ERR

    Dim datTime As Date

    datTime = DateAdd("s", intSeconds, Now)

    Do
        ' Yield to other programs (better than using DoEvents which eats up all the CPU cycles)
        Sleep 100
        DoEvents
    Loop Until Now >= datTime

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure WaitSeconds of Class aegitClass"
    Resume PROC_EXIT
End Sub

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

'''x     strFile = aestrSourceLocation & aeRefTxtFile
    
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
        If blnDebug Then Debug.Print , , , vbaProj.References(i).Guid

        Print #1, , , vbaProj.References(i).Name, vbaProj.References(i).Description
        Print #1, , , , vbaProj.References(i).FullPath
        Print #1, , , , vbaProj.References(i).Guid

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
    Dim tdf As DAO.TableDef
    Dim intTNLen As Integer

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "LongestTableName"

    intTNLen = 0
    Set dbs = CurrentDb()
    For Each tdf In CurrentDb.TableDefs
        If Not (Left(tdf.Name, 4) = "MSys" _
                Or Left(tdf.Name, 4) = "~TMP" _
                Or Left(tdf.Name, 3) = "zzz") Then
            If Len(tdf.Name) > intTNLen Then
                intTNLen = Len(tdf.Name)
            End If
        End If
    Next tdf

    LongestTableName = intTNLen

PROC_EXIT:
    Set tdf = Nothing
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
'=======================================================================
' Author:   Peter F. Ennis
' Date:     December 5, 2012
' Comment:  Return length of field properties for text output alignment
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'=======================================================================

    Dim dbs As DAO.Database
    Dim tblDef As DAO.TableDef
    Dim fld As DAO.Field

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "LongestFieldPropsName"

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
'=========================================================================
' Procedure : GetLinkedTableCurrentPath
' DateTime  : 08/23/2010
' Author    : Rx
' Purpose   : Returns Current Path of a Linked Table in Access
' Updates   : Peter F. Ennis
' Updated   : All notes moved to change log
' History   : See comment details, basChangeLog, commit messages on github
'=========================================================================
    On Error GoTo PROC_ERR
    GetLinkedTableCurrentPath = Mid(CurrentDb.TableDefs(MyLinkedTable).Connect, InStr(1, CurrentDb.TableDefs(MyLinkedTable).Connect, "=") + 1)
        ' Non-linked table returns blank - Instr removes the "Database="

PROC_EXIT:
    On Error Resume Next
    Exit Function

PROC_ERR:
    Select Case Err.Number
        ' Case ###         ' Add your own error management or log error to logging table
        Case Else
            ' Add your own custom log usage function
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
'=============================================================================
' Purpose:  Display the field names, types, sizes and descriptions for a table
' Argument: Name of a table in the current database
' Updates:  Peter F. Ennis
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'=============================================================================

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim sLen As Long
    Dim strLinkedTablePath As String
    
    Dim blnDebug As Boolean

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "TableInfo"

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
  
    'Debug.Print ">>>", SizeString("-", sLen, TextLeft, "-")
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
        'If blnDebug And aeintFDLen <> 11 Then
            Debug.Print SizeString(fld.Name, aeintFNLen, TextLeft, " ") _
                & aestr4 & SizeString(FieldTypeName(fld), aeintFTLen, TextLeft, " ") _
                & aestr4 & SizeString(fld.size, aeintFSize, TextLeft, " ") _
                & aestr4 & SizeString(GetDescrip(fld), aeintFDLen, TextLeft, " ")
        End If
        Print #1, SizeString(fld.Name, aeintFNLen, TextLeft, " ") _
            & aestr4 & SizeString(FieldTypeName(fld), aeintFTLen, TextLeft, " ") _
            & aestr4 & SizeString(fld.size, aeintFSize, TextLeft, " ") _
            & aestr4 & SizeString(GetDescrip(fld), aeintFDLen, TextLeft, " ")
    Next
    If blnDebug Then Debug.Print
    'If blnDebug And aeintFDLen <> 11 Then Debug.Print
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

Private Function LongestTableDescription(strTblName As String) As Integer
' ?LongestTableDescription("tblCaseManager")

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim strLFD As String

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "LongestTableDescription"

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
    PopCallStack
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LongestTableDescription of Class aegitClass"
    LongestTableDescription = -1
    GlobalErrHandler
    Resume PROC_EXIT

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

Private Function aeDocumentTables(Optional varDebug As Variant) As Boolean
' Ref: http://www.tek-tips.com/faqs.cfm?fid=6905
' Document the tables, fields, and relationships
' Tables, field type, primary keys, foreign keys, indexes
' Relationships in the database with table, foreign table, primary keys, foreign keys
' Ref: http://allenbrowne.com/func-06.html

    'Dim strDoc As String
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim blnDebug As Boolean
    Dim blnResult As Boolean
    Dim intFailCount As Integer
    Dim strFile As String

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "aeDocumentTables"

    On Error GoTo PROC_ERR

    intFailCount = 0
    
    LongestFieldPropsName
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Name=" & aestrLFN
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Name Length=" & aeintFNLen
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Name Table Name=" & aestrLFNTN
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Description=" & aestrLFD
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Description Length=" & aeintFDLen
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Type=" & aestrLFT
    If Not IsMissing(varDebug) Then Debug.Print "Longest Field Type Length=" & aeintFTLen

    ' Reset values
    aestrLFN = ""
    If aeintFNLen < 11 Then aeintFNLen = 11     ' Minimum required by design
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

'''x     strFile = aestrSourceLocation & aeTblTxtFile
    
    If Dir(strFile) <> "" Then
        ' The file exists
        If Not FileLocked(strFile) Then KillProperly (strFile)
        Open strFile For Append As #1
    Else
        If Not FileLocked(strFile) Then Open strFile For Append As #1
    End If

    For Each tdf In CurrentDb.TableDefs
        If Not (Left(tdf.Name, 4) = "MSys" _
                Or Left(tdf.Name, 4) = "~TMP" _
                Or Left(tdf.Name, 3) = "zzz") Then
            If blnDebug Then
                blnResult = TableInfo(tdf.Name, "WithDebugging")
                If Not blnResult Then intFailCount = intFailCount + 1
                If blnDebug And aeintFDLen <> 11 Then Debug.Print "aeintFDLen=" & aeintFDLen
            Else
                blnResult = TableInfo(tdf.Name)
                If Not blnResult Then intFailCount = intFailCount + 1
            End If
            'Debug.Print
            aeintFDLen = 0
        End If
    Next tdf

    'If intFailCount > 0 Then
    '    aeDocumentTables = False
    'Else
    '    aeDocumentTables = True
    'End If
    If blnDebug Then
        Debug.Print "intFailCount = " & intFailCount
        'Debug.Print "aeDocumentTables = " & aeDocumentTables
    End If

    aeDocumentTables = True

PROC_EXIT:
    Set fld = Nothing
    Set tdf = Nothing
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

Private Function isPK(tdf As DAO.TableDef, strField As String) As Boolean
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tdf.Indexes
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

Private Function isIndex(tdf As DAO.TableDef, strField As String) As Boolean
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tdf.Indexes
        For Each fld In idx.Fields
            If strField = fld.Name Then
                isIndex = True
                Exit Function
            End If
        Next fld
    Next idx
End Function

Private Function isFK(tdf As DAO.TableDef, strField As String) As Boolean
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    For Each idx In tdf.Indexes
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

'''x     strFile = aestrSourceLocation & aeRelTxtFile
    
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

Private Sub KillProperly(Killfile As String)
' Ref: http://word.mvps.org/faqs/macrosvba/DeleteFiles.htm

    ' Use a call stack and global error handler
    'If gcfHandleErrors Then On Error GoTo PROC_ERR
    'PushCallStack "KillProperly"

    On Error GoTo PROC_ERR

TryAgain:
    If Len(Dir(Killfile)) > 0 Then
        SetAttr Killfile, vbNormal
        Kill Killfile
    End If

PROC_EXIT:
    'PopCallStack
    Exit Sub

PROC_ERR:
    If Err = 70 Or Err = 75 Then
        Pause (0.25)
        Resume TryAgain
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " Killfile=" & Killfile & " (" & Err.Description & ") in procedure KillProperly of Class aegitClass"
    'GlobalErrHandler
    Resume PROC_EXIT

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
 
Private Function IsFileLocked(PathFileName As String) As Boolean
' Ref: http://accessexperts.com/blog/2012/03/06/checking-if-files-are-locked/

    'Debug.Print "IsFileLocked Entry PathFileName=" & PathFileName

    ' Use a call stack and global error handler
    'If gcfHandleErrors Then On Error GoTo PROC_ERR
    'PushCallStack "IsFileLocked"

    On Error GoTo PROC_ERR

    Dim i As Integer

    'Debug.Print , Len(Dir$(PathFileName))
    If Len(Dir$(PathFileName)) Then
        i = FreeFile()
        Open PathFileName For Random Access Read Write Lock Read Write As #i
        Lock i 'Redundant but let's be 100% sure
        Unlock i
        Close i
    Else
        'Err.Raise 53
    End If

PROC_EXIT:
    On Error GoTo 0
    'PopCallStack
    Exit Function

PROC_ERR:
    Select Case Err.Number
        Case 70 'Unable to acquire exclusive lock
            MsgBox "A:Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegitClass"
            IsFileLocked = True
        Case 9
            MsgBox "B:Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegitClass" & _
                    vbCrLf & "IsFileLocked Entry PathFileName=" & PathFileName, vbCritical, "ERROR=9"
            IsFileLocked = False
            'If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegitClass"
            'GlobalErrHandler
            Resume PROC_EXIT
        Case Else
            MsgBox "C:Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegitClass"
            'If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFileLocked of Class aegitClass"
            'GlobalErrHandler
            Resume PROC_EXIT
    End Select
    Resume

End Function

Private Sub KillAllFiles(strLoc As String, Optional varDebug As Variant)

    Dim strFile As String
    Dim blnDebug As Boolean

    ' Use a call stack and global error handler
    If gcfHandleErrors Then On Error GoTo PROC_ERR
    PushCallStack "KillAllFiles"

    Debug.Print "aeDocumentTheDatabase"
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of KillAllFiles is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of KillAllFiles is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If

    If strLoc = "src" Then
        ' Delete all the exported src files
'''x         strFile = Dir(aestrSourceLocation & "*.*")
        Do While strFile <> ""
'''x             KillProperly (aestrSourceLocation & strFile)
            ' Need to specify full path again because a file was deleted
'''x             strFile = Dir(aestrSourceLocation & "*.*")
        Loop
'''x         strFile = Dir(aestrSourceLocation & "xml\" & "*.*")
        Do While strFile <> ""
'''x             KillProperly (aestrSourceLocation & "xml\" & strFile)
            ' Need to specify full path again because a file was deleted
'''x             strFile = Dir(aestrSourceLocation & "xml\" & "*.*")
        Loop
    ElseIf strLoc = "xml" Then
        ' Delete files in xml location
        If aegitSetup Then
'''x             strFile = Dir(aestrXMLLocation & "*.*")
            Do While strFile <> ""
'''x                 KillProperly (aestrXMLLocation & strFile)
                ' Need to specify full path again because a file was deleted
'''x                 strFile = Dir(aestrXMLLocation & "*.*")
            Loop
        End If
    Else
        MsgBox "Bad strLoc", vbCritical, "STOP"
        Stop
    End If

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    If Err = 70 Then    ' Permission denied
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure KillAllFiles of Class aegitClass" _
            & vbCrLf & vbCrLf & _
            "Manually delete the files from git, compact and repair database, then try again!", vbCritical, "STOP"
        Stop
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure KillAllFiles of Class aegitClass"
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure KillAllFiles of Class aegitClass"
    GlobalErrHandler
    Resume PROC_EXIT

End Sub

Private Function FolderExists(strPath As String) As Boolean
' Ref: http://allenbrowne.com/func-11.html
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Private Sub ListOrCloseAllOpenQueries(Optional strCloseAll As Variant)
' Ref: http://msdn.microsoft.com/en-us/library/office/aa210652(v=office.11).aspx

    Dim obj As AccessObject
    Dim dbs As Object
    Set dbs = Application.CurrentData

    If IsMissing(strCloseAll) Then
        ' Search for open AccessObject objects in AllQueries collection.
        For Each obj In dbs.AllQueries
            If obj.IsLoaded = True Then
                ' Print name of obj
                Debug.Print obj.Name
            End If
        Next obj
    Else
        For Each obj In dbs.AllQueries
            If obj.IsLoaded = True Then
                ' Close obj
                DoCmd.Close acQuery, obj.Name, acSaveYes
                Debug.Print "Closed query " & obj.Name
            End If
        Next obj
    End If

End Sub

Private Function BuildTheDirectory(fso As Object, _
                                        Optional varDebug As Variant) As Boolean
'Private Function BuildTheDirectory(FSO As Scripting.FileSystemObject, _
                                        Optional varDebug As Variant) As Boolean
'*** Requires reference to "Microsoft Scripting Runtime"
'
' Ref: http://msdn.microsoft.com/en-us/library/ebkhfaaz(v=vs.85).aspx
'=======================================================================
' Author:   Peter F. Ennis
' Date:     February 8, 2011
' Comment:  Add optional debug parameter
' Requires: Reference to Microsoft Scripting Runtime
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'=======================================================================

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
    If blnDebug Then Debug.Print , , "FSO.DriveExists(THE_DRIVE) = " & fso.DriveExists(THE_DRIVE)
    If Not fso.DriveExists(THE_DRIVE) Then
        If blnDebug Then Debug.Print , , "FSO.DriveExists(THE_DRIVE) = FALSE - The drive DOES NOT EXIST !!!"
        BuildTheDirectory = False
        Exit Function
    End If
    If blnDebug Then Debug.Print , , "The drive EXISTS !!!"
'''x     If blnDebug Then Debug.Print , , "aegitUseImportFolder = " & aegitUseImportFolder
    
    If aegitImportFolder = "default" Then
        aestrImportLocation = aegitType.ImportFolder
    End If
'''x     If aegitUseImportFolder And aegitImportFolder <> "default" Then
'''x         aestrImportLocation = aegitImportFolder
'''x     End If
        
    If blnDebug Then Debug.Print , , "The import directory is: " & aestrImportLocation
   
    If fso.FolderExists(aestrImportLocation) Then
        If blnDebug Then Debug.Print , , "FSO.FolderExists(aestrImportLocation) = TRUE - The directory EXISTS !!!"
        BuildTheDirectory = False
        Exit Function
    End If
    If blnDebug Then Debug.Print , , "The import directory does NOT EXIST !!!"

'''x     If aegitUseImportFolder Then
'''x         Set objImportFolder = fso.CreateFolder(aestrImportLocation)
'''x         If blnDebug Then Debug.Print , , aestrImportLocation & " has been CREATED !!!"
'''x     End If

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

Private Function aeReadDocDatabase(blnImport As Boolean, Optional varDebug As Variant) As Boolean
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
'=======================================================================
' Author:   Peter F. Ennis
' Date:     February 8, 2011
' Comment:  Add explicit references for objects, wscript, fso
' Requires: Reference to Microsoft Scripting Runtime
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'=======================================================================

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

    If Not blnImport Then
        Debug.Print , "blnImport IS FALSE so exit aeReadDocDatabase"
        aeReadDocDatabase = False
        Exit Function
    End If
    
    Const acQuery = 1

    If aegitSourceFolder = "default" Then
'''x         aestrSourceLocation = aegitType.SourceFolder
    Else
'''x         aestrSourceLocation = aegitSourceFolder
    End If

    If aegitImportFolder = "default" Then
        aestrImportLocation = aegitType.ImportFolder
    End If
'''x     If aegitUseImportFolder And aegitImportFolder <> "default" Then
'''x         aestrImportLocation = aegitImportFolder
'''x     End If

    If blnDebug Then
        Debug.Print ">==> aeReadDocDatabase >==>"
        Debug.Print , "aegit VERSION: " & aegit_impVERSION
        Debug.Print , "aegit VERSION_DATE: " & aegit_impVERSION_DATE
'''x         Debug.Print , "SourceFolder = " & aestrSourceLocation
'''x         Debug.Print , "UseImportFolder = " & aegitUseImportFolder
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

    ' Create needed objects
    Dim fso As Object
'    Dim FSO As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    If blnDebug Then
        bln = BuildTheDirectory(fso, "WithDebugging")
        Debug.Print , "<==<"
    Else
        bln = BuildTheDirectory(fso)
    End If

'''x     If aegitUseImportFolder Then
        Dim objFolder As Object
        Set objFolder = fso.GetFolder(aegitType.ImportFolder)

        For Each MyFile In objFolder.Files
            If blnDebug Then Debug.Print "myFile = " & MyFile
            If blnDebug Then Debug.Print "myFile.Name = " & MyFile.Name
            strFileBaseName = fso.GetBaseName(MyFile.Name)
            strFileType = fso.GetExtensionName(MyFile.Name)
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
'''x     End If

    If blnDebug Then Debug.Print "<==<"
    
    aeReadDocDatabase = True

PROC_EXIT:
    Set MyFile = Nothing
    'Set ojbFolder = Nothing
    Set fso = Nothing
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
'=======================================================================
' Author:     Peter F. Ennis
' Date:       February 18, 2011
' Comment:    Return True if the object exists
' Parameters:
'             strAccObjType: "Tables", "Queries", "Forms",
'                            "Reports", "Macros", "Modules"
'             strAccObjName: The name of the object
' Updated:  All notes moved to change log
' History:  See comment details, basChangeLog, commit messages on github
'=======================================================================

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
    If blnDebug And aeExists = False Then
        Debug.Print , strAccObjName & " DOES NOT EXIST!"
        Debug.Print "<==<"
    End If

PROC_EXIT:
    Set obj = Nothing
    PopCallStack
    Exit Function

PROC_ERR:
    If Err = 3011 Then
        aeExists = False
        Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class aegitClass"
        If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class aegitClass"
        aeExists = False
    End If
    GlobalErrHandler
    Resume PROC_EXIT

End Function

Private Function GetType(Value As Long) As String
' Ref: http://bytes.com/topic/access/answers/557780-getting-string-name-enum

    Select Case Value
        Case acCheckBox
            GetType = "CheckBox"
        Case acTextBox
            GetType = "TextBox"
        Case acListBox
            GetType = "ListBox"
        Case acComboBox
            GetType = "ComboBox"
        Case Else
    End Select

End Function

Private Function fListGUID(strTableName As String) As String
' Ref: http://stackoverflow.com/questions/8237914/how-to-get-the-guid-of-a-table-in-microsoft-access
' e.g. ?fListGUID("tblThisTableHasSomeReallyLongNameButItCouldBeMuchLonger")

    Dim i As Integer
    Dim arrGUID8() As Byte
    Dim strArrGUID8(8) As String
    Dim strGuid As String

    strGuid = ""
    arrGUID8 = CurrentDb.TableDefs(strTableName).Properties("GUID").Value
    For i = 1 To 8
        If Len(Hex(arrGUID8(i))) = 1 Then
            strArrGUID8(i) = "0" & CStr(Hex(arrGUID8(i)))
        Else
            strArrGUID8(i) = Hex(arrGUID8(i))
        End If
    Next

    For i = 1 To 8
        strGuid = strGuid & strArrGUID8(i) & "-"
    Next
    fListGUID = Left(strGuid, 23)

End Function

Public Sub PrettyXML(strPathFileName As String, Optional varDebug As Variant)

    ' Beautify XML in VBA with MSXML6 only
    ' Ref: http://social.msdn.microsoft.com/Forums/en-US/409601d4-ca95-448a-aafc-aa0ee1ad67cd/beautify-xml-in-vba-with-msxml6-only?forum=xmlandnetfx
    Dim objXMLStyleSheet As Object
    Dim strXMLStyleSheet As String
    Dim objXMLDOMDoc As Object

    Dim fle As Integer
    fle = FreeFile()

    strXMLStyleSheet = "<xsl:stylesheet" & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "  xmlns:xsl=""http://www.w3.org/1999/XSL/Transform""" & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "  version=""1.0"">" & vbCrLf & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "<xsl:output method=""xml"" indent=""yes""/>" & vbCrLf & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "<xsl:template match=""@* | node()"">" & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "  <xsl:copy>" & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "    <xsl:apply-templates select=""@* | node()""/>" & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "  </xsl:copy>" & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "</xsl:template>" & vbCrLf & vbCrLf
    strXMLStyleSheet = strXMLStyleSheet & "</xsl:stylesheet>"

    Set objXMLStyleSheet = CreateObject("Msxml2.DOMDocument.6.0")

    With objXMLStyleSheet
        ' Turn off Async I/O
        .async = False
        .validateOnParse = False
        .resolveExternals = False
    End With

    objXMLStyleSheet.LoadXML (strXMLStyleSheet)
    If objXMLStyleSheet.parseError.errorCode <> 0 Then
        Debug.Print "Some Error..."
        Exit Sub
    End If

    Set objXMLDOMDoc = CreateObject("Msxml2.DOMDocument.6.0")
    With objXMLDOMDoc
        ' Turn off Async I/O
        .async = False
        .validateOnParse = False
        .resolveExternals = False
    End With

    ' Ref: http://msdn.microsoft.com/en-us/library/ms762722(v=vs.85).aspx
    ' Ref: http://msdn.microsoft.com/en-us/library/ms754585(v=vs.85).aspx
    ' Ref: http://msdn.microsoft.com/en-us/library/aa468547.aspx
    objXMLDOMDoc.Load (strPathFileName)

    Dim strXMLResDoc
    Set strXMLResDoc = CreateObject("Msxml2.DOMDocument.6.0")

    objXMLDOMDoc.transformNodeToObject objXMLStyleSheet, strXMLResDoc
    strXMLResDoc = strXMLResDoc.XML
    strXMLResDoc = Replace(strXMLResDoc, vbTab, Chr(32) & Chr(32), , , vbBinaryCompare)
    Debug.Print "Pretty XML Sample Output"
    If Not IsMissing(varDebug) Then Debug.Print strXMLResDoc

    ' Rewrite the file as pretty xml
    Open strPathFileName For Output As #fle
    Print #fle, strXMLResDoc
    Close fle

    Set objXMLDOMDoc = Nothing
    Set objXMLStyleSheet = Nothing

End Sub

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
             "Line: " & Erl & vbCrLf & _
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