Option Compare Database
Option Explicit

Public Function LoadRibbons() As Boolean
    ' Load ribbons from XML file into the database
    ' Ref: http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Load_Ribbons_Into_The_Database:..._From_XML_File

    On Error GoTo PROC_ERR
    
    Dim fle As Long
    Dim strText As String
    Dim strOut As String

    fle = FreeFile
    Open "C:\Folder\Ribbon\AccRibbon.xml" For Input As fle
    ' C:\Folder\... has to be replaced by your folder/filename.
    Do While Not EOF(fle)
        Line Input #fle, strText
        strOut = strOut & strText
    Loop
    Application.LoadCustomUI "AppRibbon_1", strOut

PROC_EXIT:
    On Error Resume Next
    Close fle
    Exit Function

PROC_ERR:
    Select Case Err
        Case 32609
            ' Ribbon already loaded
        Case Else
            MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LoadRibbons of Class aegitClass"
    End Select
    LoadRibbons = False
    Resume PROC_EXIT

End Function

Public Sub CreateRibbon()
    ' Ref: http://www.nullskull.com/q/10320914/change-ribbon-programatically.aspx
    ' Ref: http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Load_Ribbons_Into_The_Database:..._From_XML_File

    On Error Resume Next
    CodeDb.Properties.Append CodeDb.CreateProperty("aeRibbonID", dbText, "adaept")
    CodeDb.Properties("aeRibbonID") = "adaept"

End Sub