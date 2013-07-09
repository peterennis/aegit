Option Compare Database
Option Explicit

' Ref: http://www.makeuseof.com/tag/monitor-vba-apps-running-slick-script/

Public Sub Version_Test()
    Dim Edition As String
    Debug.Print "Application.Version = " & Val(Application.Version)
    Debug.Print "Application.Name = " & Val(Application.Name)
    Edition = Get_Edition
    Debug.Print Edition
End Sub

Public Function Get_Edition() As String
' Ref: http://p2p.wrox.com/excel-vba/82653-what-best-way-get-excel-version.html
' Updated for Office 2013

    Const ERRORMESSAGE As String = " : Unable to determine edition"
    Dim Sku As String

    Select Case VBA.Val(Application.Version)
        Case Is < 9
            Get_Edition = "Pre Office 2000" & ERRORMESSAGE

        Case Is < 10                            'Office 2000
            Sku = VBA.Mid$(Application.ProductCode, 4, 2)
            Get_Edition = Get_Edition_2000(Sku)

        Case Is < 11                            'Office 2002
            Sku = VBA.Mid$(Application.ProductCode, 4, 2)
            Get_Edition = Get_Edition_2002(Sku)

        Case Is < 12                            'Office 2003
            Sku = VBA.Mid$(Application.ProductCode, 4, 2)
            Get_Edition = Get_Edition_2003(Sku)

        Case Is < 13                            'Office 2007
            Sku = VBA.Mid$(Application.ProductCode, 11, 4)
            Get_Edition = Get_Edition_2007(Sku)

        Case Is < 15                            'Office 2010
            Sku = VBA.Mid$(Application.ProductCode, 11, 4)
            Get_Edition = Get_Edition_2010(Sku)

        Case Is < 16                            'Office 2013
            Sku = VBA.Mid$(Application.ProductCode, 11, 4)
            Get_Edition = Get_Edition_2013(Sku)

        Case Else
            Get_Edition = "Post Office 2010" & ERRORMESSAGE

    End Select

End Function
 
Private Function Get_Edition_2000(ByRef Sku As String) As String
    'reference: http://support.microsoft.com/kb/230848/
    Select Case Sku
        Case "00": Get_Edition_2000 = "Microsoft Office 2000 Premium Edition CD1"
        Case "01": Get_Edition_2000 = "Microsoft Office 2000 Professional Edition"
        Case "02": Get_Edition_2000 = "Microsoft Office 2000 Standard Edition"
        Case "03": Get_Edition_2000 = "Microsoft Office 2000 Small Business Edition"
        Case "04": Get_Edition_2000 = "Microsoft Office 2000 Premium CD2"
        Case "05": Get_Edition_2000 = "Office CD2 SMALL"
        Case "06": Get_Edition_2000 = "0F (reserved)"
        Case "10": Get_Edition_2000 = "Microsoft Access 2000 (standalone)"
        Case "11": Get_Edition_2000 = "Microsoft Excel 2000 (standalone)"
        Case "12": Get_Edition_2000 = "Microsoft FrontPage 2000 (standalone)"
        Case "13": Get_Edition_2000 = "Microsoft PowerPoint 2000 (standalone)"
        Case "14": Get_Edition_2000 = "Microsoft Publisher 2000 (standalone)"
        Case "15": Get_Edition_2000 = "Office Server Extensions"
        Case "16": Get_Edition_2000 = "Microsoft Outlook 2000 (standalone)"
        Case "17": Get_Edition_2000 = "Microsoft Word 2000 (standalone)"
        Case "18": Get_Edition_2000 = "Microsoft Access 2000 runtime version"
        Case "19": Get_Edition_2000 = "FrontPage Server Extensions"
        Case "1A": Get_Edition_2000 = "Publisher Standalone OEM"
        Case "1B": Get_Edition_2000 = "DMMWeb"
        Case "1C": Get_Edition_2000 = "FP WECCOM"
        Case "1D": Get_Edition_2000 = "1F (reserved standalone SKUs)"
        Case "20": Get_Edition_2000 = "2F Office Language Packs"
        Case "30": Get_Edition_2000 = "3F Proofing Tools Kit(s)"
        Case "40": Get_Edition_2000 = "Publisher Trial CD"
        Case "41": Get_Edition_2000 = "Publisher Trial Web"
        Case "42": Get_Edition_2000 = "SBB"
        Case "43": Get_Edition_2000 = "SBT"
        Case "44": Get_Edition_2000 = "SBT CD2"
        Case "45": Get_Edition_2000 = "SBTART"
        Case "46": Get_Edition_2000 = "Web Components"
        Case "47": Get_Edition_2000 = "VP Office CD2 with LVP"
        Case "48": Get_Edition_2000 = "VP PUB with LVP"
        Case "49": Get_Edition_2000 = "VP PUB with LVP OEM"
        Case "4F": Get_Edition_2000 = "Access 2000 SR-1 Run-Time Minimum"
    End Select
End Function
 
Private Function Get_Edition_2002(ByRef Sku As String) As String
    'reference: http://support.microsoft.com/kb/302663/
    Select Case Sku
        Case "11": Get_Edition_2002 = "Microsoft Office XP Professional"
        Case "12": Get_Edition_2002 = "Microsoft Office XP Standard"
        Case "13": Get_Edition_2002 = "Microsoft Office XP Small Business"
        Case "14": Get_Edition_2002 = "Microsoft Office XP Web Server"
        Case "15": Get_Edition_2002 = "Microsoft Access 2002"
        Case "16": Get_Edition_2002 = "Microsoft Excel 2002"
        Case "17": Get_Edition_2002 = "Microsoft FrontPage 2002"
        Case "18": Get_Edition_2002 = "Microsoft PowerPoint 2002"
        Case "19": Get_Edition_2002 = "Microsoft Publisher 2002"
        Case "1A": Get_Edition_2002 = "Microsoft Outlook 2002"
        Case "1B": Get_Edition_2002 = "Microsoft Word 2002"
        Case "1C": Get_Edition_2002 = "Microsoft Access 2002 Runtime"
        Case "1D": Get_Edition_2002 = "Microsoft FrontPage Server Extensions 2002"
        Case "1E": Get_Edition_2002 = "Microsoft Office Multilingual User Interface Pack"
        Case "1F": Get_Edition_2002 = "Microsoft Office Proofing Tools Kit"
        Case "20": Get_Edition_2002 = "System Files Update"
        Case "22": Get_Edition_2002 = "unused"
        Case "23": Get_Edition_2002 = "Microsoft Office Multilingual User Interface Pack Wizard"
        Case "24": Get_Edition_2002 = "Microsoft Office XP Resource Kit"
        Case "25": Get_Edition_2002 = "Microsoft Office XP Resource Kit Tools (download from Web)"
        Case "26": Get_Edition_2002 = "Microsoft Office Web Components"
        Case "27": Get_Edition_2002 = "Microsoft Project 2002"
        Case "28": Get_Edition_2002 = "Microsoft Office XP Professional with FrontPage"
        Case "29": Get_Edition_2002 = "Microsoft Office XP Professional Subscription"
        Case "2A": Get_Edition_2002 = "Microsoft Office XP Small Business Edition Subscription"
        Case "2B": Get_Edition_2002 = "Microsoft Publisher 2002 Deluxe Edition"
        Case "2F": Get_Edition_2002 = "Standalone IME (JPN Only)"
        Case "30": Get_Edition_2002 = "Microsoft Office XP Media Content"
        Case "31": Get_Edition_2002 = "Microsoft Project 2002 Web Client"
        Case "32": Get_Edition_2002 = "Microsoft Project 2002 Web Server"
        Case "33": Get_Edition_2002 = "Microsoft Office XP PIPC1 (Pre Installed PC) (JPN Only)"
        Case "34": Get_Edition_2002 = "Microsoft Office XP PIPC2 (Pre Installed PC) (JPN Only)"
        Case "35": Get_Edition_2002 = "Microsoft Office XP Media Content Deluxe"
        Case "3A": Get_Edition_2002 = "Project 2002 Standard"
        Case "3B": Get_Edition_2002 = "Project 2002 Professional"
        Case "51": Get_Edition_2002 = "Microsoft Office Visio Professional 2003"
        Case "54": Get_Edition_2002 = "Microsoft Office Visio Standard 2003"
    End Select
End Function
 
Private Function Get_Edition_2003(ByRef Sku As String) As String
    'reference: http://support.microsoft.com/kb/832672/
    Select Case Sku
        Case "11": Get_Edition_2003 = "Microsoft Office Professional Enterprise Edition 2003"
        Case "12": Get_Edition_2003 = "Microsoft Office Standard Edition 2003"
        Case "13": Get_Edition_2003 = "Microsoft Office Basic Edition 2003"
        Case "14": Get_Edition_2003 = "Microsoft Windows SharePoint Services 2.0"
        Case "15": Get_Edition_2003 = "Microsoft Office Access 2003"
        Case "16": Get_Edition_2003 = "Microsoft Office Excel 2003"
        Case "17": Get_Edition_2003 = "Microsoft Office FrontPage 2003"
        Case "18": Get_Edition_2003 = "Microsoft Office PowerPoint 2003"
        Case "19": Get_Edition_2003 = "Microsoft Office Publisher 2003"
        Case "1A": Get_Edition_2003 = "Microsoft Office Outlook Professional 2003"
        Case "1B": Get_Edition_2003 = "Microsoft Office Word 2003"
        Case "1C": Get_Edition_2003 = "Microsoft Office Access 2003 Runtime"
        Case "1E": Get_Edition_2003 = "Microsoft Office 2003 User Interface Pack"
        Case "1F": Get_Edition_2003 = "Microsoft Office 2003 Proofing Tools"
        Case "23": Get_Edition_2003 = "Microsoft Office 2003 Multilingual User Interface Pack"
        Case "24": Get_Edition_2003 = "Microsoft Office 2003 Resource Kit"
        Case "26": Get_Edition_2003 = "Microsoft Office XP Web Components"
        Case "2E": Get_Edition_2003 = "Microsoft Office 2003 Research Service SDK"
        Case "44": Get_Edition_2003 = "Microsoft Office InfoPath 2003"
        Case "83": Get_Edition_2003 = "Microsoft Office 2003 HTML Viewer"
        Case "92": Get_Edition_2003 = "Windows SharePoint Services 2.0 English Template Pack"
        Case "93": Get_Edition_2003 = "Microsoft Office 2003 English Web Parts and Components"
        Case "A1": Get_Edition_2003 = "Microsoft Office OneNote 2003"
        Case "A4": Get_Edition_2003 = "Microsoft Office 2003 Web Components"
        Case "A5": Get_Edition_2003 = "Microsoft SharePoint Migration Tool 2003"
        Case "AA": Get_Edition_2003 = "Microsoft Office PowerPoint 2003 Presentation Broadcast"
        Case "AB": Get_Edition_2003 = "Microsoft Office PowerPoint 2003 Template Pack 1"
        Case "AC": Get_Edition_2003 = "Microsoft Office PowerPoint 2003 Template Pack 2"
        Case "AD": Get_Edition_2003 = "Microsoft Office PowerPoint 2003 Template Pack 3"
        Case "AE": Get_Edition_2003 = "Microsoft Organization Chart 2.0"
        Case "CA": Get_Edition_2003 = "Microsoft Office Small Business Edition 2003"
        Case "D0": Get_Edition_2003 = "Microsoft Office Access 2003 Developer Extensions"
        Case "DC": Get_Edition_2003 = "Microsoft Office 2003 Smart Document SDK"
        Case "E0": Get_Edition_2003 = "Microsoft Office Outlook Standard 2003"
        Case "E3": Get_Edition_2003 = "Microsoft Office Professional Edition 2003 (with InfoPath 2003)"
        Case "FD": Get_Edition_2003 = "Microsoft Office Outlook 2003 (distributed by MSN)"
        Case "FF": Get_Edition_2003 = "Microsoft Office 2003 Edition Language Interface Pack"
        Case "F8": Get_Edition_2003 = "Remove Hidden Data Tool"
        Case "3A": Get_Edition_2003 = "Microsoft Office Project Standard 2003"
        Case "3B": Get_Edition_2003 = "Microsoft Office Project Professional 2003"
        Case "32": Get_Edition_2003 = "Microsoft Office Project Server 2003"
        Case "51": Get_Edition_2003 = "Microsoft Office Visio Professional 2003"
        Case "52": Get_Edition_2003 = "Microsoft Office Visio Viewer 2003"
        Case "53": Get_Edition_2003 = "Microsoft Office Visio Standard 2003"
        Case "55": Get_Edition_2003 = "Microsoft Office Visio for Enterprise Architects 2003"
        Case "5E": Get_Edition_2003 = "Microsoft Office Visio 2003 Multilingual User Interface Pack"
    End Select
End Function
 
Private Function Get_Edition_2007(ByRef Sku As String) As String
    'reference: http://support.microsoft.com/kb/928516/
    Select Case Sku
        Case "0011":  Get_Edition_2007 = "Microsoft Office Professional Plus 2007"
        Case "0012":  Get_Edition_2007 = "Microsoft Office Standard 2007"
        Case "0013":  Get_Edition_2007 = "Microsoft Office Basic 2007"
        Case "0014":  Get_Edition_2007 = "Microsoft Office Professional 2007"
        Case "0015":  Get_Edition_2007 = "Microsoft Office Access 2007"
        Case "0016":  Get_Edition_2007 = "Microsoft Office Excel 2007"
        Case "0017":  Get_Edition_2007 = "Microsoft Office SharePoint Designer 2007"
        Case "0018":  Get_Edition_2007 = "Microsoft Office PowerPoint 2007"
        Case "0019":  Get_Edition_2007 = "Microsoft Office Publisher 2007"
        Case "001A":  Get_Edition_2007 = "Microsoft Office Outlook 2007"
        Case "001B":  Get_Edition_2007 = "Microsoft Office Word 2007"
        Case "001C":  Get_Edition_2007 = "Microsoft Office Access Runtime 2007"
        Case "0020":  Get_Edition_2007 = "Microsoft Office Compatibility Pack for Word, Excel, and PowerPoint 2007 File Formats"
        Case "0026":  Get_Edition_2007 = "Microsoft Expression Web"
        Case "0029":  Get_Edition_2007 = "Microsoft Office Excel 2007"
        Case "002B":  Get_Edition_2007 = "Microsoft Office Word 2007"
        Case "002E":  Get_Edition_2007 = "Microsoft Office Ultimate 2007"
        Case "002F":  Get_Edition_2007 = "Microsoft Office Home and Student 2007"
        Case "0030":  Get_Edition_2007 = "Microsoft Office Enterprise 2007"
        Case "0031":  Get_Edition_2007 = "Microsoft Office Professional Hybrid 2007"
        Case "0033":  Get_Edition_2007 = "Microsoft Office Personal 2007"
        Case "0035":  Get_Edition_2007 = "Microsoft Office Professional Hybrid 2007"
        Case "0037":  Get_Edition_2007 = "Microsoft Office PowerPoint 2007"
        Case "003A":  Get_Edition_2007 = "Microsoft Office Project Standard 2007"
        Case "003B":  Get_Edition_2007 = "Microsoft Office Project Professional 2007"
        Case "0044":  Get_Edition_2007 = "Microsoft Office InfoPath 2007"
        Case "0051":  Get_Edition_2007 = "Microsoft Office Visio Professional 2007"
        Case "0052":  Get_Edition_2007 = "Microsoft Office Visio Viewer 2007"
        Case "0053":  Get_Edition_2007 = "Microsoft Office Visio Standard 2007"
        Case "00A1":  Get_Edition_2007 = "Microsoft Office OneNote 2007"
        Case "00A3":  Get_Edition_2007 = "Microsoft Office OneNote Home Student 2007"
        Case "00A7":  Get_Edition_2007 = "Calendar Printing Assistant for Microsoft Office Outlook 2007"
        Case "00A9":  Get_Edition_2007 = "Microsoft Office InterConnect 2007"
        Case "00AF":  Get_Edition_2007 = "Microsoft Office PowerPoint Viewer 2007 (English)"
        Case "00B0":  Get_Edition_2007 = "The Microsoft Save as PDF add-in"
        Case "00B1":  Get_Edition_2007 = "The Microsoft Save as XPS add-in"
        Case "00B2":  Get_Edition_2007 = "The Microsoft Save as PDF or XPS add-in"
        Case "00BA":  Get_Edition_2007 = "Microsoft Office Groove 2007"
        Case "00CA":  Get_Edition_2007 = "Microsoft Office Small Business 2007"
        Case "00E0":  Get_Edition_2007 = "Microsoft Office Outlook 2007"
        Case "10D7":  Get_Edition_2007 = "Microsoft Office InfoPath Forms Services"
        Case "110D":  Get_Edition_2007 = "Microsoft Office SharePoint Server 2007"
        Case "1122":  Get_Edition_2007 = "Windows SharePoint Services Developer Resources 1.2"
        Case "0010":  Get_Edition_2007 = "SKU - Microsoft Software Update for Web Folders (English) 12"
    End Select
End Function
 
Private Function Get_Edition_2010(ByRef Sku As String) As String
    'reference: http://support.microsoft.com/kb/2186281
    Select Case Sku
        Case "0011":  Get_Edition_2010 = "Microsoft Office Professional Plus 2010"
        Case "0012":  Get_Edition_2010 = "Microsoft Office Standard 2010"
        Case "0013":  Get_Edition_2010 = "Microsoft Office Home and Business 2010"
        Case "0014":  Get_Edition_2010 = "Microsoft Office Professional 2010"
        Case "0015":  Get_Edition_2010 = "Microsoft Access 2010"
        Case "0016":  Get_Edition_2010 = "Microsoft Excel 2010"
        Case "0017":  Get_Edition_2010 = "Microsoft SharePoint Designer 2010"
        Case "0018":  Get_Edition_2010 = "Microsoft PowerPoint 2010"
        Case "0019":  Get_Edition_2010 = "Microsoft Publisher 2010"
        Case "001A":  Get_Edition_2010 = "Microsoft Outlook 2010"
        Case "001B":  Get_Edition_2010 = "Microsoft Word 2010"
        Case "001C":  Get_Edition_2010 = "Microsoft Access Runtime 2010"
        Case "001F":  Get_Edition_2010 = "Microsoft Office Proofing Tools Kit Compilation 2010"
        Case "002F":  Get_Edition_2010 = "Microsoft Office Home and Student 2010"
        Case "003A":  Get_Edition_2010 = "Microsoft Project Standard 2010"
        Case "003B":  Get_Edition_2010 = "Microsoft Project Professional 2010"
        Case "0044":  Get_Edition_2010 = "Microsoft InfoPath 2010"
        Case "0052":  Get_Edition_2010 = "Microsoft Visio Viewer 2010"
        Case "0057":  Get_Edition_2010 = "Microsoft Visio 2010"
        Case "007A":  Get_Edition_2010 = "Microsoft Outlook Connector"
        Case "008B":  Get_Edition_2010 = "Microsoft Office Small Business Basics 2010"
        Case "00A1":  Get_Edition_2010 = "Microsoft OneNote 2010"
        Case "00AF":  Get_Edition_2010 = "Microsoft PowerPoint Viewer 2010"
        Case "00BA":  Get_Edition_2010 = "Microsoft Office SharePoint Workspace 2010"
        Case "110D":  Get_Edition_2010 = "Microsoft Office SharePoint Server 2010"
        Case "110F":  Get_Edition_2010 = "Microsoft Project Server 2010"
    End Select
End Function

Private Function Get_Edition_2013(ByRef Sku As String) As String
    'reference: http://support.microsoft.com/kb/2786054
    Select Case Sku
        Case "0011":  Get_Edition_2013 = "Microsoft Office Professional Plus 2013"
        Case "0012":  Get_Edition_2013 = "Microsoft Office Standard 2013"
        Case "0013":  Get_Edition_2013 = "Microsoft Office Home and Business 2013"
        Case "0014":  Get_Edition_2013 = "Microsoft Office Professional 2013"
        Case "0015":  Get_Edition_2013 = "Microsoft Access 2013"
        Case "0016":  Get_Edition_2013 = "Microsoft Excel 2013"
        Case "0017":  Get_Edition_2013 = "Microsoft SharePoint Designer 2013"
        Case "0018":  Get_Edition_2013 = "Microsoft PowerPoint 2013"
        Case "0019":  Get_Edition_2013 = "Microsoft Publisher 2013"
        Case "001A":  Get_Edition_2013 = "Microsoft Outlook 2013"
        Case "001B":  Get_Edition_2013 = "Microsoft Word 2013"
        Case "001C":  Get_Edition_2013 = "Microsoft Access Runtime 2013"
        Case "001F":  Get_Edition_2013 = "Microsoft Office Proofing Tools Kit Compilation 2013"
        Case "002F":  Get_Edition_2013 = "Microsoft Office Home and Student 2013"
        Case "003A":  Get_Edition_2013 = "Microsoft Project Standard 2013"
        Case "003B":  Get_Edition_2013 = "Microsoft Project Professional 2013"
        Case "0044":  Get_Edition_2013 = "Microsoft InfoPath 2013"
        Case "0051":  Get_Edition_2013 = "Microsoft Visio Professional 2013"
        Case "0053":  Get_Edition_2013 = "Microsoft Visio Standard 2013"
        Case "00A1":  Get_Edition_2013 = "Microsoft OneNote 2013"
        Case "00BA":  Get_Edition_2013 = "Microsoft Office SharePoint Workspace 2013"
        Case "110D":  Get_Edition_2013 = "Microsoft Office SharePoint Server 2013"
        Case "110F":  Get_Edition_2013 = "Microsoft Project Server 2013"
        Case "012B":  Get_Edition_2013 = "Microsoft Lync 2013"
    End Select
End Function

Public Sub IsAppOpen(strAppName As String)
' Ref: http://www.ehow.com/how_12111794_determine-excel-already-running-vba.html
' Ref: http://msdn.microsoft.com/en-us/library/office/aa164798(v=office.10).aspx

    Const ERR_APP_NOTRUNNING As Long = 429

    On Error GoTo Err_IsAppOpen:

    Dim objApp As Object

    Select Case strAppName
        Case "Access"
            Set objApp = GetObject(, "Access.Application")
            'Debug.Print objApp.Name
            If (objApp.Name = "Microsoft Access") Then
                Debug.Print "Access is running!"
            End If
        Case "Excel"
            'Debug.Print objApp.Name
            Set objApp = GetObject(, "Excel.Application")
            If (objApp = "Microsoft Excel") Then
                Debug.Print "Excel is running!"
            End If
        Case Else
            Debug.Print "Invalid App Name"
    End Select

    Set objApp = Nothing

Exit_IsAppOpen:
    Exit Sub

Err_IsAppOpen:
    If Err.Number = ERR_APP_NOTRUNNING Then
        Debug.Print strAppName & " is not running!"
    End If
    Set objApp = Nothing

End Sub

 