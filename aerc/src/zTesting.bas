Option Compare Database
Option Explicit

'''x Private Const FOR_READING = 1
Public Const Desktop = &H10&
Public Const MyDocuments = &H5&
'''x Private aestrXMLLocation As String
'

Public Sub ObjectCounts()
 
    Dim qry As DAO.QueryDef
    Dim cnt As DAO.Container
 
    'Delete all TEMP queries ...
    For Each qry In CurrentDb.QueryDefs
        If Left(qry.Name, 1) = "~" Then
            CurrentDb.QueryDefs.Delete qry.Name
            CurrentDb.QueryDefs.Refresh
        End If
    Next qry
 
    'Print the values to the immediate window
    With CurrentDb
 
        Debug.Print "--- From the DAO.Database ---"
        Debug.Print "-----------------------------"
        Debug.Print "Tables (Inc. System tbls): " & .TableDefs.Count
        Debug.Print "Querys: " & .QueryDefs.Count & vbCrLf
 
        For Each cnt In .Containers
            Debug.Print cnt.Name & ":" & cnt.Documents.Count
        Next cnt
 
    End With
 
    'Use the "Project" collections to get the counts of objects
    With CurrentProject
        Debug.Print vbCrLf & "--- From the Access 'Project' ---"
        Debug.Print "---------------------------------"
        Debug.Print "Forms: " & .AllForms.Count
        Debug.Print "Reports: " & .AllReports.Count
        Debug.Print "DataAccessPages: " & .AllDataAccessPages.Count
        Debug.Print "Modules: " & .AllModules.Count
        Debug.Print "Macros (aka Scripts): " & .AllMacros.Count
    End With
 
End Sub

Public Function SpFolder(SpName)

    Dim objShell As Object
    Set objShell = CreateObject("Shell.Application")
    Dim objFolder As Object
    Set objFolder = objShell.Namespace(SpName)
    Dim objFolderItem As Object
    Set objFolderItem = objFolder.Self

    SpFolder = objFolderItem.Path

End Function
   
Public Sub ExportAllModulesToFile()
' Ref: http://wiki.lessthandot.com/index.php/Code_and_Code_Windows
' Ref: http://stackoverflow.com/questions/2794480/exporting-code-from-microsoft-access
' The reference for the FileSystemObject Object is Windows Script Host Object Model
' but it not necessary to add the reference for this procedure.

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fil As Object
    Dim strMod As String
    Dim mdl As Object
    Dim i As Integer
    Dim strTxtFile As String


    ' Set up the file
    Debug.Print "CurrentProject.Name = " & CurrentProject.Name
    strTxtFile = SpFolder(Desktop) & "\" & Replace(CurrentProject.Name, ".", "_") & ".txt"
    Debug.Print "strTxtFile = " & strTxtFile
    Set fil = fso.CreateTextFile(SpFolder(Desktop) & "\" _
            & Replace(CurrentProject.Name, ".", " ") & ".txt")

    ' For each component in the project ...
    For Each mdl In VBE.ActiveVBProject.VBComponents
        ' using the count of lines ...
        If Left(mdl.Name, 3) <> "zzz" Then
            Debug.Print mdl.Name
            i = VBE.ActiveVBProject.VBComponents(mdl.Name).CodeModule.CountOfLines
            ' put the code in a string ...
            If i > 0 Then
                strMod = VBE.ActiveVBProject.VBComponents(mdl.Name).CodeModule.Lines(1, i)
            End If
            ' and then write it to a file, first marking the start with
            ' some equal signs and the component name.
            fil.WriteLine String(15, "=") & vbCrLf & mdl.Name _
                & vbCrLf & String(15, "=") & vbCrLf & strMod
        End If
    Next
       
    'Close eveything
    fil.Close
    Set fso = Nothing

End Sub

Public Sub SaveTableMacros()

    ' Export Table Data to XML
    ' Ref: http://technet.microsoft.com/en-us/library/ee692914.aspx
    Application.ExportXML acExportTable, "Items", "C:\Temp\ItemsData.xml"

    ' Save table macros as XML
    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=99179
    Application.SaveAsText acTableDataMacro, "Items", "C:\Temp\Items.xml"
    Debug.Print , "Items table macros saved to C:\Temp\Items.xml"

    ' Beautify XML in VBA with MSXML6 only
    ' Ref: http://social.msdn.microsoft.com/Forums/en-US/409601d4-ca95-448a-aafc-aa0ee1ad67cd/beautify-xml-in-vba-with-msxml6-only?forum=xmlandnetfx
    Dim objXMLStyleSheet As Object
    Dim strXMLStyleSheet As String
    Dim objXMLDOMDoc As Object

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

'''x    Dim objXMLResDoc As Object
    Set objXMLStyleSheet = CreateObject("Msxml2.DOMDocument.6.0")

    With objXMLStyleSheet
        'Turn off Async I/O
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
        'Turn off Async I/O
        .async = False
        .validateOnParse = False
        .resolveExternals = False
    End With

    ' Ref: http://msdn.microsoft.com/en-us/library/ms762722(v=vs.85).aspx
    ' Ref: http://msdn.microsoft.com/en-us/library/ms754585(v=vs.85).aspx
    ' Ref: http://msdn.microsoft.com/en-us/library/aa468547.aspx
    objXMLDOMDoc.Load ("C:\Temp\Items.xml")

    Dim strXMLResDoc
    Set strXMLResDoc = CreateObject("Msxml2.DOMDocument.6.0")

    objXMLDOMDoc.transformNodeToObject objXMLStyleSheet, strXMLResDoc
    strXMLResDoc = strXMLResDoc.XML
    strXMLResDoc = Replace(strXMLResDoc, vbTab, Chr(32) & Chr(32), , , vbBinaryCompare)
    Debug.Print "Pretty XML Sample Output"
    Debug.Print strXMLResDoc

    Set objXMLDOMDoc = Nothing
    Set objXMLStyleSheet = Nothing

End Sub

Public Sub SetRefToLibrary()
' http://www.exceltoolset.com/setting-a-reference-to-the-vba-extensibility-library-by-code/
' Adjusted for Microsoft Access
' Create a reference to the VBA Extensibility library
    On Error Resume Next        ' in case the reference already exits
    Access.Application.VBE.ActiveVBProject.References _
                  .AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 0
End Sub

Public Sub ApplicationInformation()
' Ref: http://msdn.microsoft.com/en-us/library/office/aa223101(v=office.11).aspx
' Ref: http://msdn.microsoft.com/en-us/library/office/aa173218(v=office.11).aspx
' Ref: http://msdn.microsoft.com/en-us/library/office/ff845735(v=office.15).aspx

    Dim intProjType As Integer
    Dim strProjType As String
    Dim lng As Long

    intProjType = Application.CurrentProject.ProjectType

    Select Case intProjType
        Case 0 ' acNull
            strProjType = "acNull"
        Case 1 ' acADP
            strProjType = "acADP"
        Case 2 ' acMDB
            strProjType = "acMDB"
        Case Else
            MsgBox "Can't determine ProjectType"
    End Select

    Debug.Print Application.CurrentProject.FullName
    Debug.Print "Project Type", intProjType, strProjType
    lng = CodeLinesInProjectCount

End Sub

Public Sub TestRegKey()

    Dim strKey As String

    ' Office 2010
    'strKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Common\General\ReportAddinCustomUIErrors"
    ' Office 2013
    strKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Common\General\ReportAddinCustomUIErrors"
    If RegKeyExists(strKey) Then
        Debug.Print strKey & " Exists!"
    Else
        Debug.Print strKey & " Does NOT Exist!"
    End If

End Sub

Public Function RegKeyExists(strRegKey As String) As Boolean
' Return True if the registry key i_RegKey was found and False if not
' Ref: http://vba-corner.livejournal.com/3054.html

    On Error GoTo ErrorHandler

    ' Use Windows scripting and try to read the registry key
    Dim myWS As Object
    Set myWS = CreateObject("WScript.Shell")

    myWS.RegRead strRegKey
    ' Key was found
    RegKeyExists = True
    Exit Function
  
ErrorHandler:
    ' Key was not found
    RegKeyExists = False

End Function