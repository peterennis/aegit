Option Compare Database
Option Explicit

Public Sub FormUseDefaultPrinter()
' Ref: http://msdn.microsoft.com/en-us/library/office/ff845464(v=office.15).aspx

    Dim obj As Object
    For Each obj In CurrentProject.AllForms
        DoCmd.OpenForm FormName:=obj.Name, View:=acViewDesign
        If Not Forms(obj.Name).UseDefaultPrinter Then
            Forms(obj.Name).UseDefaultPrinter = True
            DoCmd.Save ObjectType:=acForm, ObjectName:=obj.Name
        End If
        DoCmd.Close
    Next obj

End Sub

Public Sub ReportUseDefaultPrinter()
' Ref: http://msdn.microsoft.com/en-us/library/office/ff845464(v=office.15).aspx

    Dim obj As Object
    For Each obj In CurrentProject.AllReports
        DoCmd.OpenReport ReportName:=obj.Name, View:=acViewDesign
        If Not Reports(obj.Name).UseDefaultPrinter Then
            Reports(obj.Name).UseDefaultPrinter = True
            DoCmd.Save ObjectType:=acReport, ObjectName:=obj.Name
        End If
        DoCmd.Close
    Next obj

End Sub

Public Sub SaveTableMacros()

    ' Export Table Data to XML
    ' Ref: http://technet.microsoft.com/en-us/library/ee692914.aspx
'    Application.ExportXML acExportTable, "aeItems", "C:\Temp\aeItemsData.xml"

    ' Save table macros as XML
    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=99179
    Application.SaveAsText acTableDataMacro, "aeItems", "C:\Temp\aeItems.xml"
    Debug.Print , "Items table macros saved to C:\Temp\aeItems.xml"

    PrettyXML "C:\Temp\aeItems.xml"

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