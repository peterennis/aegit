Option Compare Database
Option Explicit

Public Sub ExportRibbon()

    Dim strDb As String
    Dim lngPath As Long
    Dim lngRev As Long
    Dim strLeft As String

    strDb = Application.CurrentDb.Name
    lngPath = Len(strDb)
    lngRev = InStrRev(strDb, "\")
    strLeft = Left(strDb, lngPath - (lngPath - lngRev))

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

Public Function OutputTableDataMacros() As Boolean
' Ref: http://stackoverflow.com/questions/9206153/how-to-export-access-2010-data-macros
'====================================================================
' Author:   Peter F. Ennis
' Date:     February 16, 2014
'====================================================================

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef

    ' Use a call stack and global error handler
    'If gcfHandleErrors Then On Error GoTo PROC_ERR
    'PushCallStack "OutputTableDataMacros"

    On Error GoTo PROC_ERR

    OutputTableDataMacros = True

    Set dbs = CurrentDb()
    For Each tdf In CurrentDb.TableDefs
        If Not (Left(tdf.Name, 4) = "MSys" _
                Or Left(tdf.Name, 4) = "~TMP" _
                Or Left(tdf.Name, 3) = "zzz") Then
            Debug.Print tdf.Name
            SaveAsText acTableDataMacro, tdf.Name, "C:\Temp\table_" & tdf.Name & "_DataMacro.xml"
        End If
    Next tdf

PROC_EXIT:
    Set tdf = Nothing
    Set dbs = Nothing
    'PopCallStack
    Exit Function

PROC_ERR:
    If Err = 2950 Then ' Reserved Error
        Resume Next
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputTableDataMacros of Class aegitClass"
    OutputTableDataMacros = False
    'GlobalErrHandler
    Resume PROC_EXIT

End Function