Option Compare Database
Option Explicit

Private Const TEST_FILE_PATH As String = "C:\TEMP\"
Private Const FOR_READING = 1
Public Const Desktop = &H10&
Public Const MyDocuments = &H5&
Private aestrXMLLocation As String
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
    Dim objFolder As Object
    Dim objFolderItem As Object

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(SpName)

    Set objFolderItem = objFolder.Self

    SpFolder = objFolderItem.Path

End Function
   
Public Sub ExportAllModulesToFile()
' Ref: http://wiki.lessthandot.com/index.php/Code_and_Code_Windows
' Ref: http://stackoverflow.com/questions/2794480/exporting-code-from-microsoft-access
' The reference for the FileSystemObject Object is Windows Script Host Object Model
' but it not necessary to add the reference for this procedure.

    Dim fso As Object
    Dim fil As Object
    Dim strMod As String
    Dim mdl As Object
    Dim i As Integer
    Dim strTxtFile As String

    Set fso = CreateObject("Scripting.FileSystemObject")

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

Public Sub TestPropertiesOutput()
' Ref: http://www.everythingaccess.com/tutorials.asp?ID=Accessing-detailed-file-information-provided-by-the-Operating-System
' Ref: http://www.techrepublic.com/article/a-simple-solution-for-tracking-changes-to-access-data/
' Ref: http://social.msdn.microsoft.com/Forums/office/en-US/480c17b3-e3d1-4f98-b1d6-fa16b23c6a0d/please-help-to-edit-the-table-query-form-and-modules-modified-date
'
' Ref: http://perfectparadigm.com/tip001.html
'SELECT MSysObjects.DateCreate, MSysObjects.DateUpdate,
'MSysObjects.Name , MSysObjects.Type
'FROM MSysObjects;

    Debug.Print ">>>frm_Dummy"
    Debug.Print "DateCreated", DBEngine(0)(0).Containers("Forms")("frm_Dummy").Properties("DateCreated").Value
    Debug.Print "LastUpdated", DBEngine(0)(0).Containers("Forms")("frm_Dummy").Properties("LastUpdated").Value

' *** Ref: http://support.microsoft.com/default.aspx?scid=kb%3Ben-us%3B299554 ***
'When the user initially creates a new Microsoft Access specific-object, such as a form), the database engine still
'enters the current date and time into the DateCreate and DateUpdate columns in the MSysObjects table. However, when
'the user modifies and saves the object, Microsoft Access does not notify the database engine; therefore, the
'DateUpdate column always stays the same.

' Ref: http://questiontrack.com/how-can-i-display-a-last-modified-time-on-ms-access-form-995507.html

    Dim obj As AccessObject
    Dim dbs As Object

    Set dbs = Application.CurrentData
    Set obj = dbs.AllTables("tblThisTableHasSomeReallyLongNameButItCouldBeMuchLonger")
    Debug.Print ">>>" & obj.Name
    Debug.Print "DateCreated: " & obj.DateCreated
    Debug.Print "DateModified: " & obj.DateModified

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

    Dim objXMLResDoc As Object

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

Public Function CodeLinesInProjectCount() As Long
' Ref: http://www.cpearson.com/excel/vbe.aspx
' Adjusted for Microsoft Access and Late Binding. No reference needed.
' Access.Application is used. Returns -1 if the VBProject is locked.

    Dim VBP As Object               'VBIDE.VBProject
    Dim VBComp As Object            'VBIDE.VBComponent
    Dim LineCount As Long

    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=245480
    Const vbext_pp_locked = 1

    Set VBP = Access.Application.VBE.ActiveVBProject

    If VBP.Protection = vbext_pp_locked Then
        CodeLinesInProjectCount = -1
        Exit Function
    End If

    For Each VBComp In VBP.VBComponents
        If Left(VBComp.Name, 3) <> "zzz" Then
            Debug.Print VBComp.Name, VBComp.CodeModule.CountOfLines
        End If
        LineCount = LineCount + VBComp.CodeModule.CountOfLines
    Next VBComp

    CodeLinesInProjectCount = LineCount

    Set VBP = Nothing

End Function

Public Sub ListFilesRecursively()
' Ref: http://blogs.msdn.com/b/gstemp/archive/2004/08/10/212113.aspx
'====================================================================
' Purpose:  List Files Recursively
' Author:   Peter Ennis
' Date:     February 10, 2011
' Comment:  Fix to work in VBA. Based on MSDN sample for WScript
' Requires: Reference to Microsoft Scripting Runtime
'====================================================================

   Dim strFolder As String
   Dim objFSO As Object
   Dim objFolder As Object
   Dim objFile As Object
   Dim colFiles As Object

   strFolder = TEST_FILE_PATH

   ' Create needed objects
   Dim wsh As Object  ' As Object if late-bound
   Set wsh = CreateObject("WScript.Shell")
    
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Set objFolder = objFSO.GetFolder(strFolder)

   Debug.Print "objFolder.Path = " & objFolder.Path

   Set colFiles = objFolder.Files

   For Each objFile In colFiles
       Debug.Print "objFile.Path = " & objFile.Path
   Next

   ShowSubFolders objFolder
   Debug.Print "DONE !!!"

   Set wsh = Nothing
   Set objFSO = Nothing
   Set objFolder = Nothing
   Set colFiles = Nothing

End Sub
 
Private Sub ShowSubFolders(objFolder)
' Ref: http://blogs.msdn.com/b/gstemp/archive/2004/08/10/212113.aspx

   Dim objFile As Object
   Dim objSubFolder As Object
   Dim colFolders As Object
   Dim colFiles As Object
   Dim wsh As Object  ' As Object if late-bound

   Set wsh = CreateObject("WScript.Shell")
   Set colFolders = objFolder.SubFolders
    
   For Each objSubFolder In colFolders
  
       Debug.Print "objSubFolder.Path = " & objSubFolder.Path
       Set colFiles = objSubFolder.Files
  
       For Each objFile In colFiles
           Debug.Print "objFile.Path = " & objFile.Path
       Next

       ShowSubFolders objSubFolder
   Next

   Set wsh = Nothing
   Set colFolders = Nothing

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

    Dim myWS As Object

    On Error GoTo ErrorHandler

    ' Use Windows scripting and try to read the registry key
    Set myWS = CreateObject("WScript.Shell")

    myWS.RegRead strRegKey
    ' Key was found
    RegKeyExists = True
    Exit Function
  
ErrorHandler:
    ' Key was not found
    RegKeyExists = False

End Function

Public Sub PrinterInfo()
' Ref: http://msdn.microsoft.com/en-us/library/office/aa139946(v=office.10).aspx
' Ref: http://answers.microsoft.com/en-us/office/forum/office_2010-access/how-do-i-change-default-printers-in-vba/d046a937-6548-4d2b-9517-7f622e2cfed2

    Dim prt As Printer
    Dim prtCount As Integer
    Dim i As Integer

    prtCount = Application.Printers.Count
    Debug.Print "Number of Printers=" & prtCount
    For Each prt In Printers
        Debug.Print , prt.DeviceName
    Next prt

    For i = 0 To prtCount - 1
        Debug.Print "DeviceName=" & Application.Printers(i).DeviceName
        Debug.Print , "BottomMargin=" & Application.Printers(i).BottomMargin
        Debug.Print , "ColorMode=" & Application.Printers(i).ColorMode
        Debug.Print , "ColumnSpacing=" & Application.Printers(i).ColumnSpacing
        Debug.Print , "Copies=" & Application.Printers(i).Copies
        Debug.Print , "DataOnly=" & Application.Printers(i).DataOnly
        Debug.Print , "DefaultSize=" & Application.Printers(i).DefaultSize
        Debug.Print , "DriverName=" & Application.Printers(i).DriverName
        Debug.Print , "Duplex=" & Application.Printers(i).Duplex
        Debug.Print , "ItemLayout=" & Application.Printers(i).ItemLayout
        Debug.Print , "ItemsAcross=" & Application.Printers(i).ItemsAcross
        Debug.Print , "ItemSizeHeight=" & Application.Printers(i).ItemSizeHeight
        Debug.Print , "ItemSizeWidth=" & Application.Printers(i).ItemSizeWidth
        Debug.Print , "LeftMargin=" & Application.Printers(i).LeftMargin
        Debug.Print , "Orientation=" & Application.Printers(i).Orientation
        Debug.Print , "PaperBin=" & Application.Printers(i).PaperBin
        Debug.Print , "PaperSize=" & Application.Printers(i).PaperSize
        Debug.Print , "Port=" & Application.Printers(i).Port
        Debug.Print , "PrintQuality=" & Application.Printers(i).PrintQuality
        Debug.Print , "RightMargin=" & Application.Printers(i).RightMargin
        Debug.Print , "RowSpacing=" & Application.Printers(i).RowSpacing
        Debug.Print , "TopMargin=" & Application.Printers(i).TopMargin
    Next

End Sub