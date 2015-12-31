Option Compare Database
Option Explicit

Public Sub Version_Test()
    On Error GoTo 0
    Debug.Print GetEdition(Application.Version, Application.ProductCode)
End Sub

Public Function GetEdition( _
                    ByRef strAppVersion As String, _
                    ByRef strGuid As String _
                            ) As String
' Ref: http://www.makeuseof.com/tag/monitor-vba-apps-running-slick-script/
' Ref: http://p2p.wrox.com/excel-vba/82653-what-best-way-get-excel-version.html
' Ref: http://colinlegg.wordpress.com/2013/02/02/office-edition-in-vba/
'
' Ref: https://community.spiceworks.com/topic/150065-how-to-remove-ms-office-2010-standard-registry-keys
' This one is for Office Home and Student 2010): {90140000-003D-0000-0000-0000000FF1CE}

    On Error GoTo PROC_ERR

    Const strERR_MSG As String = "Unable to determine edition"

    Dim strSku As String

    Debug.Print "strAppVersion=" & strAppVersion
    Debug.Print "Val(strAppVersion)=" & Val(strAppVersion)
    Debug.Print "strGuid = " & strGuid
    Select Case Val(strAppVersion)
        Case Is < 9
            GetEdition = "Pre Office 2000: " & strERR_MSG

        Case Is < 10                            ' Office 2000
            strSku = Mid$(strGuid, 4, 2)
            GetEdition = GetEdition2000(strSku)

        Case Is < 11                            ' Office 2002
            strSku = Mid$(strGuid, 4, 2)
            GetEdition = GetEdition2002(strSku)

        Case Is < 12                            ' Office 2003
            strSku = Mid$(strGuid, 4, 2)
            GetEdition = GetEdition2003(strSku)

        Case Is < 13                            ' Office 2007
            strSku = Mid$(strGuid, 11, 4)
            GetEdition = GetEdition2007(strSku)

        Case Is < 15                            ' Office 2010
            strSku = Mid$(strGuid, 11, 4)
            GetEdition = GetEdition2010(strSku)

        Case Is < 16                            ' Office 2013
            strSku = Mid$(strGuid, 11, 4)
            Debug.Print "strSku=" & strSku
            GetEdition = GetEdition2013(strSku)

        Case Is < 17                            ' Office 2016
            strSku = Mid$(strGuid, 11, 4)
            Debug.Print "strSku=" & strSku
            GetEdition = GetEdition2016(strSku)

        Case Else
            GetEdition = "Post Office 2016: " & strERR_MSG

    End Select

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetEdition of Class aegit_expClass"
    GetEdition = strERR_MSG & vbNewLine & _
                    "Error Number: " & CStr(Err.Number) & _
                    vbNewLine & "Error Desc: " & Err.Description
 
End Function
 
Private Function GetEdition2000(ByRef strSku As String) As String
' Ref: http://support.microsoft.com/kb/230848/
 
    On Error GoTo 0
    Select Case strSku
        Case "00"
            GetEdition2000 = "Microsoft Office 2000 Premium Edition CD1"
        Case "01"
            GetEdition2000 = "Microsoft Office 2000 Professional Edition"
        Case "02"
            GetEdition2000 = "Microsoft Office 2000 Standard Edition"
        Case "03"
            GetEdition2000 = "Microsoft Office 2000 Small Business Edition"
        Case "04"
            GetEdition2000 = "Microsoft Office 2000 Premium CD2"
        Case "05"
            GetEdition2000 = "Office CD2 SMALL"
        Case "06" To "09", "0A" To "0F"
            GetEdition2000 = "(reserved)"
        Case "10"
            GetEdition2000 = "Microsoft Access 2000 (standalone)"
        Case "11"
            GetEdition2000 = "Microsoft Excel 2000 (standalone)"
        Case "12"
            GetEdition2000 = "Microsoft FrontPage 2000 (standalone)"
        Case "13"
            GetEdition2000 = "Microsoft PowerPoint 2000 (standalone)"
        Case "14"
            GetEdition2000 = "Microsoft Publisher 2000 (standalone)"
        Case "15"
            GetEdition2000 = "Office Server Extensions"
        Case "16"
            GetEdition2000 = "Microsoft Outlook 2000 (standalone)"
        Case "17"
            GetEdition2000 = "Microsoft Word 2000 (standalone)"
        Case "18"
            GetEdition2000 = "Microsoft Access 2000 runtime version"
        Case "19"
            GetEdition2000 = "FrontPage Server Extensions"
        Case "1A"
            GetEdition2000 = "Publisher Standalone OEM"
        Case "1B"
            GetEdition2000 = "DMMWeb"
        Case "1C"
            GetEdition2000 = "FP WECCOM"
        Case "1D" To "1F"
            GetEdition2000 = "(reserved standalone SKUs)"
        Case "20" To "29", "2A" To "2F"
            GetEdition2000 = "Office Language Packs"
        Case "30" To "39", "3A" To "3F"
            GetEdition2000 = "Proofing Tools Kit(s)"
        Case "40"
            GetEdition2000 = "Publisher Trial CD"
        Case "41"
            GetEdition2000 = "Publisher Trial Web"
        Case "42"
            GetEdition2000 = "SBB"
        Case "43"
            GetEdition2000 = "SBT"
        Case "44"
            GetEdition2000 = "SBT CD2"
        Case "45"
            GetEdition2000 = "SBTART"
        Case "46"
            GetEdition2000 = "Web Components"
        Case "47"
            GetEdition2000 = "VP Office CD2 with LVP"
        Case "48"
            GetEdition2000 = "VP PUB with LVP"
        Case "49"
            GetEdition2000 = "VP PUB with LVP OEM"
        Case "4F"
            GetEdition2000 = "Access 2000 SR-1 Run-Time Minimum"
        Case Else
            MsgBox "Error: GetEdition2000", vbCritical, "ERROR"
    End Select
End Function
 
Private Function GetEdition2002(ByRef strSku As String) As String
' Ref: http://support.microsoft.com/kb/302663/
 
    On Error GoTo 0
    Select Case strSku
        Case "11"
            GetEdition2002 = "Microsoft Office XP Professional"
        Case "12"
            GetEdition2002 = "Microsoft Office XP Standard"
        Case "13"
            GetEdition2002 = "Microsoft Office XP Small Business"
        Case "14"
            GetEdition2002 = "Microsoft Office XP Web Server"
        Case "15"
            GetEdition2002 = "Microsoft Access 2002"
        Case "16"
            GetEdition2002 = "Microsoft Excel 2002"
        Case "17"
            GetEdition2002 = "Microsoft FrontPage 2002"
        Case "18"
            GetEdition2002 = "Microsoft PowerPoint 2002"
        Case "19"
            GetEdition2002 = "Microsoft Publisher 2002"
        Case "1A"
            GetEdition2002 = "Microsoft Outlook 2002"
        Case "1B"
            GetEdition2002 = "Microsoft Word 2002"
        Case "1C"
            GetEdition2002 = "Microsoft Access 2002 Runtime"
        Case "1D"
            GetEdition2002 = "Microsoft FrontPage Server Extensions 2002"
        Case "1E"
            GetEdition2002 = "Microsoft Office Multilingual User Interface Pack"
        Case "1F"
            GetEdition2002 = "Microsoft Office Proofing Tools Kit"
        Case "20"
            GetEdition2002 = "System Files Update"
        Case "22"
            GetEdition2002 = "unused"
        Case "23"
            GetEdition2002 = "Microsoft Office Multilingual User Interface Pack Wizard"
        Case "24"
            GetEdition2002 = "Microsoft Office XP Resource Kit"
        Case "25"
            GetEdition2002 = "Microsoft Office XP Resource Kit Tools (download from Web)"
        Case "26"
            GetEdition2002 = "Microsoft Office Web Components"
        Case "27"
            GetEdition2002 = "Microsoft Project 2002"
        Case "28"
            GetEdition2002 = "Microsoft Office XP Professional with FrontPage"
        Case "29"
            GetEdition2002 = "Microsoft Office XP Professional Subscription"
        Case "2A"
            GetEdition2002 = "Microsoft Office XP Small Business Edition Subscription"
        Case "2B"
            GetEdition2002 = "Microsoft Publisher 2002 Deluxe Edition"
        Case "2F"
            GetEdition2002 = "Standalone IME (JPN Only)"
        Case "30"
            GetEdition2002 = "Microsoft Office XP Media Content"
        Case "31"
            GetEdition2002 = "Microsoft Project 2002 Web Client"
        Case "32"
            GetEdition2002 = "Microsoft Project 2002 Web Server"
        Case "33"
            GetEdition2002 = "Microsoft Office XP PIPC1 (Pre Installed PC) (JPN Only)"
        Case "34"
            GetEdition2002 = "Microsoft Office XP PIPC2 (Pre Installed PC) (JPN Only)"
        Case "35"
            GetEdition2002 = "Microsoft Office XP Media Content Deluxe"
        Case "3A"
            GetEdition2002 = "Project 2002 Standard"
        Case "3B"
            GetEdition2002 = "Project 2002 Professional"
        Case "51"
            GetEdition2002 = "Microsoft Office Visio Professional 2003"
        Case "54"
            GetEdition2002 = "Microsoft Office Visio Standard 2003"
        Case Else
            MsgBox "Error: GetEdition2002", vbCritical, "ERROR"
    End Select
End Function
 
Private Function GetEdition2003(ByRef strSku As String) As String
' Ref: http://support.microsoft.com/kb/832672/
 
    On Error GoTo 0
    Select Case strSku
        Case "11"
            GetEdition2003 = "Microsoft Office Professional Enterprise Edition 2003"
        Case "12"
            GetEdition2003 = "Microsoft Office Standard Edition 2003"
        Case "13"
            GetEdition2003 = "Microsoft Office Basic Edition 2003"
        Case "14"
            GetEdition2003 = "Microsoft Windows SharePoint Services 2.0"
        Case "15"
            GetEdition2003 = "Microsoft Office Access 2003"
        Case "16"
            GetEdition2003 = "Microsoft Office Excel 2003"
        Case "17"
            GetEdition2003 = "Microsoft Office FrontPage 2003"
        Case "18"
            GetEdition2003 = "Microsoft Office PowerPoint 2003"
        Case "19"
            GetEdition2003 = "Microsoft Office Publisher 2003"
        Case "1A"
            GetEdition2003 = "Microsoft Office Outlook Professional 2003"
        Case "1B"
            GetEdition2003 = "Microsoft Office Word 2003"
        Case "1C"
            GetEdition2003 = "Microsoft Office Access 2003 Runtime"
        Case "1E"
            GetEdition2003 = "Microsoft Office 2003 User Interface Pack"
        Case "1F"
            GetEdition2003 = "Microsoft Office 2003 Proofing Tools"
        Case "23"
            GetEdition2003 = "Microsoft Office 2003 Multilingual User Interface Pack"
        Case "24"
            GetEdition2003 = "Microsoft Office 2003 Resource Kit"
        Case "26"
            GetEdition2003 = "Microsoft Office XP Web Components"
        Case "2E"
            GetEdition2003 = "Microsoft Office 2003 Research Service SDK"
        Case "44"
            GetEdition2003 = "Microsoft Office InfoPath 2003"
        Case "83"
            GetEdition2003 = "Microsoft Office 2003 HTML Viewer"
        Case "92"
            GetEdition2003 = "Windows SharePoint Services 2.0 English Template Pack"
        Case "93"
            GetEdition2003 = "Microsoft Office 2003 English Web Parts and Components"
        Case "A1"
            GetEdition2003 = "Microsoft Office OneNote 2003"
        Case "A4"
            GetEdition2003 = "Microsoft Office 2003 Web Components"
        Case "A5"
            GetEdition2003 = "Microsoft SharePoint Migration Tool 2003"
        Case "AA"
            GetEdition2003 = "Microsoft Office PowerPoint 2003 Presentation Broadcast"
        Case "AB"
            GetEdition2003 = "Microsoft Office PowerPoint 2003 Template Pack 1"
        Case "AC"
            GetEdition2003 = "Microsoft Office PowerPoint 2003 Template Pack 2"
        Case "AD"
            GetEdition2003 = "Microsoft Office PowerPoint 2003 Template Pack 3"
        Case "AE"
            GetEdition2003 = "Microsoft Organization Chart 2.0"
        Case "CA"
            GetEdition2003 = "Microsoft Office Small Business Edition 2003"
        Case "D0"
            GetEdition2003 = "Microsoft Office Access 2003 Developer Extensions"
        Case "DC"
            GetEdition2003 = "Microsoft Office 2003 Smart Document SDK"
        Case "E0"
            GetEdition2003 = "Microsoft Office Outlook Standard 2003"
        Case "E3"
            GetEdition2003 = "Microsoft Office Professional Edition 2003 (with InfoPath 2003)"
        Case "FD"
            GetEdition2003 = "Microsoft Office Outlook 2003 (distributed by MSN)"
        Case "FF"
            GetEdition2003 = "Microsoft Office 2003 Edition Language Interface Pack"
        Case "F8"
            GetEdition2003 = "Remove Hidden Data Tool"
        Case "3A"
            GetEdition2003 = "Microsoft Office Project Standard 2003"
        Case "3B"
            GetEdition2003 = "Microsoft Office Project Professional 2003"
        Case "32"
            GetEdition2003 = "Microsoft Office Project Server 2003"
        Case "51"
            GetEdition2003 = "Microsoft Office Visio Professional 2003"
        Case "52"
            GetEdition2003 = "Microsoft Office Visio Viewer 2003"
        Case "53"
            GetEdition2003 = "Microsoft Office Visio Standard 2003"
        Case "55"
            GetEdition2003 = "Microsoft Office Visio for Enterprise Architects 2003"
        Case "5E"
            GetEdition2003 = "Microsoft Office Visio 2003 Multilingual User Interface Pack"
        Case Else
            MsgBox "Error: GetEdition2003", vbCritical, "ERROR"
    End Select
End Function
 
Private Function GetEdition2007(ByRef strSku As String) As String
' Ref: http://support.microsoft.com/kb/928516/
 
    On Error GoTo 0
    Select Case strSku
        Case "0011"
            GetEdition2007 = "Microsoft Office Professional Plus 2007"
        Case "0012"
            GetEdition2007 = "Microsoft Office Standard 2007"
        Case "0013"
            GetEdition2007 = "Microsoft Office Basic 2007"
        Case "0014"
            GetEdition2007 = "Microsoft Office Professional 2007"
        Case "0015"
            GetEdition2007 = "Microsoft Office Access 2007"
        Case "0016"
            GetEdition2007 = "Microsoft Office Excel 2007"
        Case "0017"
            GetEdition2007 = "Microsoft Office SharePoint Designer 2007"
        Case "0018"
            GetEdition2007 = "Microsoft Office PowerPoint 2007"
        Case "0019"
            GetEdition2007 = "Microsoft Office Publisher 2007"
        Case "001A"
            GetEdition2007 = "Microsoft Office Outlook 2007"
        Case "001B"
            GetEdition2007 = "Microsoft Office Word 2007"
        Case "001C"
            GetEdition2007 = "Microsoft Office Access Runtime 2007"
        Case "0020"
            GetEdition2007 = "Microsoft Office Compatibility Pack for Word, Excel, and PowerPoint 2007 File Formats"
        Case "0026"
            GetEdition2007 = "Microsoft Expression Web"
        Case "0029"
            GetEdition2007 = "Microsoft Office Excel 2007"
        Case "002B"
            GetEdition2007 = "Microsoft Office Word 2007"
        Case "002E"
            GetEdition2007 = "Microsoft Office Ultimate 2007"
        Case "002F"
            GetEdition2007 = "Microsoft Office Home and Student 2007"
        Case "0030"
            GetEdition2007 = "Microsoft Office Enterprise 2007"
        Case "0031"
            GetEdition2007 = "Microsoft Office Professional Hybrid 2007"
        Case "0033"
            GetEdition2007 = "Microsoft Office Personal 2007"
        Case "0035"
            GetEdition2007 = "Microsoft Office Professional Hybrid 2007"
        Case "0037"
            GetEdition2007 = "Microsoft Office PowerPoint 2007"
        Case "003A"
            GetEdition2007 = "Microsoft Office Project Standard 2007"
        Case "003B"
            GetEdition2007 = "Microsoft Office Project Professional 2007"
        Case "0044"
            GetEdition2007 = "Microsoft Office InfoPath 2007"
        Case "0051"
            GetEdition2007 = "Microsoft Office Visio Professional 2007"
        Case "0052"
            GetEdition2007 = "Microsoft Office Visio Viewer 2007"
        Case "0053"
            GetEdition2007 = "Microsoft Office Visio Standard 2007"
        Case "00A1"
            GetEdition2007 = "Microsoft Office OneNote 2007"
        Case "00A3"
            GetEdition2007 = "Microsoft Office OneNote Home Student 2007"
        Case "00A7"
            GetEdition2007 = "Calendar Printing Assistant for Microsoft Office Outlook 2007"
        Case "00A9"
            GetEdition2007 = "Microsoft Office InterConnect 2007"
        Case "00AF"
            GetEdition2007 = "Microsoft Office PowerPoint Viewer 2007 (English)"
        Case "00B0"
            GetEdition2007 = "The Microsoft Save as PDF add-in"
        Case "00B1"
            GetEdition2007 = "The Microsoft Save as XPS add-in"
        Case "00B2"
            GetEdition2007 = "The Microsoft Save as PDF or XPS add-in"
        Case "00BA"
            GetEdition2007 = "Microsoft Office Groove 2007"
        Case "00CA"
            GetEdition2007 = "Microsoft Office Small Business 2007"
        Case "00E0"
            GetEdition2007 = "Microsoft Office Outlook 2007"
        Case "10D7"
            GetEdition2007 = "Microsoft Office InfoPath Forms Services"
        Case "110D"
            GetEdition2007 = "Microsoft Office SharePoint Server 2007"
        Case "1122"
            GetEdition2007 = "Windows SharePoint Services Developer Resources 1.2"
        Case "0010"
            GetEdition2007 = "SKU - Microsoft Software Update for Web Folders (English) 12"
        Case Else
            MsgBox "Error: GetEdition2007", vbCritical, "ERROR"
    End Select
End Function
 
Private Function GetEdition2010(ByRef strSku As String) As String
' Ref: http://support.microsoft.com/kb/2186281
 
    On Error GoTo 0
    Select Case strSku
        Case "0011"
            GetEdition2010 = "Microsoft Office Professional Plus 2010"
        Case "0012"
            GetEdition2010 = "Microsoft Office Standard 2010"
        Case "0013"
            GetEdition2010 = "Microsoft Office Home and Business 2010"
        Case "0014"
            GetEdition2010 = "Microsoft Office Professional 2010"
        Case "0015"
            GetEdition2010 = "Microsoft Access 2010"
        Case "0016"
            GetEdition2010 = "Microsoft Excel 2010"
        Case "0017"
            GetEdition2010 = "Microsoft SharePoint Designer 2010"
        Case "0018"
            GetEdition2010 = "Microsoft PowerPoint 2010"
        Case "0019"
            GetEdition2010 = "Microsoft Publisher 2010"
        Case "001A"
            GetEdition2010 = "Microsoft Outlook 2010"
        Case "001B"
            GetEdition2010 = "Microsoft Word 2010"
        Case "001C"
            GetEdition2010 = "Microsoft Access Runtime 2010"
        Case "001F"
            GetEdition2010 = "Microsoft Office Proofing Tools Kit Compilation 2010"
        Case "002F"
            GetEdition2010 = "Microsoft Office Home and Student 2010"
        Case "003A"
            GetEdition2010 = "Microsoft Project Standard 2010"
        Case "003B"
            GetEdition2010 = "Microsoft Project Professional 2010"
        Case "0044"
            GetEdition2010 = "Microsoft InfoPath 2010"
        Case "0052"
            GetEdition2010 = "Microsoft Visio Viewer 2010"
        Case "0057"
            GetEdition2010 = "Microsoft Visio 2010"
        Case "007A"
            GetEdition2010 = "Microsoft Outlook Connector"
        Case "008B"
            GetEdition2010 = "Microsoft Office Small Business Basics 2010"
        Case "00A1"
            GetEdition2010 = "Microsoft OneNote 2010"
        Case "00AF"
            GetEdition2010 = "Microsoft PowerPoint Viewer 2010"
        Case "00BA"
            GetEdition2010 = "Microsoft Office SharePoint Workspace 2010"
        Case "110D"
            GetEdition2010 = "Microsoft Office SharePoint Server 2010"
        Case "110F"
            GetEdition2010 = "Microsoft Project Server 2010"
        Case Else
            MsgBox "Error: GetEdition2010", vbCritical, "ERROR"
            Debug.Print "strSku = " & strSku
    End Select
End Function
 
Private Function GetEdition2013(ByRef strSku As String) As String
' Ref: http://support.microsoft.com/kb/2786054
 
    On Error GoTo 0
    Debug.Print "GetEdition2013 strSku=" & strSku
    Select Case strSku
        Case "0011"
            GetEdition2013 = "Microsoft Office Professional Plus 2013"
        Case "0012"
            GetEdition2013 = "Microsoft Office Standard 2013"
        Case "0013"
            GetEdition2013 = "Microsoft Office Home and Business 2013"
        Case "0014"
            GetEdition2013 = "Microsoft Office Professional 2013"
        Case "000F"
            GetEdition2013 = "Microsoft Access 2013"
        Case "0016"
            GetEdition2013 = "Microsoft Excel 2013"
        Case "0017"
            GetEdition2013 = "Microsoft SharePoint Designer 2013"
        Case "0018"
            GetEdition2013 = "Microsoft PowerPoint 2013"
        Case "0019"
            GetEdition2013 = "Microsoft Publisher 2013"
        Case "001A"
            GetEdition2013 = "Microsoft Outlook 2013"
        Case "001B"
            GetEdition2013 = "Microsoft Word 2013"
        Case "001C"
            GetEdition2013 = "Microsoft Access Runtime 2013"
        Case "001F"
            GetEdition2013 = "Microsoft Office Proofing Tools Kit Compilation 2013"
        Case "002F"
            GetEdition2013 = "Microsoft Office Home and Student 2013"
        Case "003A"
            GetEdition2013 = "Microsoft Project Standard 2013"
        Case "003B"
            GetEdition2013 = "Microsoft Project Professional 2013"
        Case "0044"
            GetEdition2013 = "Microsoft InfoPath 2013"
        Case "0051"
            GetEdition2013 = "Microsoft Visio Professional 2013"
        Case "0053"
            GetEdition2013 = "Microsoft Visio Standard 2013"
        Case "00A1"
            GetEdition2013 = "Microsoft OneNote 2013"
        Case "00BA"
            GetEdition2013 = "Microsoft Office SharePoint Workspace 2013"
        Case "110D"
            GetEdition2013 = "Microsoft Office SharePoint Server 2013"
        Case "110F"
            GetEdition2013 = "Microsoft Project Server 2013"
        Case "012B"
            GetEdition2013 = "Microsoft Lync 2013"
        Case Else
            MsgBox "Error: GetEdition2013", vbCritical, "ERROR"
    End Select
End Function

Private Function GetEdition2016(ByRef strSku As String) As String
' Ref: http://support.microsoft.com/kb/2786054
 
    On Error GoTo 0
    Debug.Print "GetEdition2016 strSku=" & strSku
    Select Case strSku
        Case "0011"
            GetEdition2016 = "Microsoft Office Professional Plus FIXME"
        Case "0012"
            GetEdition2016 = "Microsoft Office Standard FIXME"
        Case "0013"
            GetEdition2016 = "Microsoft Office Home and Business FIXME"
        Case "0014"
            GetEdition2016 = "Microsoft Office Professional FIXME"
        Case "000F"
            GetEdition2016 = "Microsoft Access FIXME"
        Case "0016"
            GetEdition2016 = "Microsoft Excel FIXME"
        Case "0017"
            GetEdition2016 = "Microsoft SharePoint Designer FIXME"
        Case "0018"
            GetEdition2016 = "Microsoft PowerPoint FIXME"
        Case "0019"
            GetEdition2016 = "Microsoft Publisher FIXME"
        Case "001A"
            GetEdition2016 = "Microsoft Outlook FIXME"
        Case "001B"
            GetEdition2016 = "Microsoft Word FIXME"
        Case "001C"
            GetEdition2016 = "Microsoft Access Runtime FIXME"
        Case "001F"
            GetEdition2016 = "Microsoft Office Proofing Tools Kit Compilation FIXME"
        Case "002F"
            GetEdition2016 = "Microsoft Office Home and Student FIXME"
        Case "003A"
            GetEdition2016 = "Microsoft Project Standard FIXME"
        Case "003B"
            GetEdition2016 = "Microsoft Project Professional FIXME"
        Case "0044"
            GetEdition2016 = "Microsoft InfoPath FIXME"
        Case "0051"
            GetEdition2016 = "Microsoft Visio Professional FIXME"
        Case "0053"
            GetEdition2016 = "Microsoft Visio Standard FIXME"
        Case "00A1"
            GetEdition2016 = "Microsoft OneNote FIXME"
        Case "00BA"
            GetEdition2016 = "Microsoft Office SharePoint Workspace FIXME"
        Case "110D"
            GetEdition2016 = "Microsoft Office SharePoint Server FIXME"
        Case "110F"
            GetEdition2016 = "Microsoft Project Server FIXME"
        Case "012B"
            GetEdition2016 = "Microsoft Lync FIXME"
        Case Else
            MsgBox "Error: GetEdition2016", vbCritical, "ERROR"
    End Select
End Function