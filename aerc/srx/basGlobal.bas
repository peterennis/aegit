Attribute VB_Name = "basGlobal"
Option Explicit

' Configuration for flexible sheet names
Public Const gstrSetupInformation = "Project Information"
Public Const gstrMenuTitle = "adaept Process Management Menu"
Public Const gstrMsgTitle = "adaept Process Management"
'
Public Const gstrPassword As String = ""
Public Const gblnAutoLock As Boolean = True
Public Const gblnDisplayTheProjectMenu = False
'
' Setting gblnUnsecured = True will turn off
' protection on the sheets and workbook.
Public gblnUnsecured As Boolean

Public gblnInformationIsVisible As Boolean

' Determine if workbook is being closed
''Public gblnWorkbookClosing As Boolean

' Determine if workbook is being opened
''Public gblnWorkbookOpening As Boolean

' Determine if workbook password was entered
Public gblnWorkbookPasswordEntered As Boolean
' Determine if Import Completed
Public gblnImportCompleted As Boolean
' Determine if Import is in operation
Public gblnImporting As Boolean



Public Type ApplicationSettingsType
    blnDisplayFormulas As Boolean
    blnDisplayGridlines As Boolean
    blnDisplayHeadings As Boolean
    blnDisplayOutline As Boolean
    blnDisplayZeros As Boolean
    blnDisplayHorizontalScrollBar As Boolean
    blnDisplayVerticalScrollBar As Boolean
    blnDisplayWorkbookTabs As Boolean
    blnDisplayFormulaBar As Boolean
    blnDisplayStatusBar As Boolean
    blnShowWindowsInTaskbar As Boolean
'    .DisplayCommentIndicator = xlCommentIndicatorOnly
    intDisplayCommentIndicator As Integer
'    ActiveSheet.DisplayAutomaticPageBreaks = True
    blnDisplayAutomaticPageBreaks As Boolean
'    ActiveWorkbook.DisplayDrawingObjects = xlHide
    intDisplayDrawingObjects As Integer
End Type

Public gtypAppSettings As ApplicationSettingsType

Public Function FileExists(strFullPath As String) As Boolean
' Ref:      http://www.contextures.com/xlfaqMac.html#FileExist
    FileExists = Len(Dir$(strFullPath))
End Function

Public Sub RememberApplicationSettings(Optional blnDebug As Boolean)

    On Error GoTo Err_RememberApplicationSettings

    With ActiveWindow
        gtypAppSettings.blnDisplayFormulas = .DisplayFormulas
        gtypAppSettings.blnDisplayGridlines = .DisplayGridlines
        gtypAppSettings.blnDisplayHeadings = .DisplayHeadings
        gtypAppSettings.blnDisplayOutline = .DisplayOutline
        gtypAppSettings.blnDisplayZeros = .DisplayZeros
        gtypAppSettings.blnDisplayHorizontalScrollBar = .DisplayHorizontalScrollBar
        gtypAppSettings.blnDisplayVerticalScrollBar = .DisplayVerticalScrollBar
        gtypAppSettings.blnDisplayWorkbookTabs = .DisplayWorkbookTabs
    End With
    '
    With Application
        gtypAppSettings.blnDisplayFormulaBar = .DisplayFormulaBar
        gtypAppSettings.blnDisplayStatusBar = .DisplayStatusBar
        gtypAppSettings.blnShowWindowsInTaskbar = .ShowWindowsInTaskbar
' .DisplayCommentIndicator = xlCommentIndicatorOnly
        gtypAppSettings.intDisplayCommentIndicator = .DisplayCommentIndicator
    End With
    '
' ActiveSheet.DisplayAutomaticPageBreaks = True
    gtypAppSettings.blnDisplayAutomaticPageBreaks = ActiveSheet.DisplayAutomaticPageBreaks
' ActiveWorkbook.DisplayDrawingObjects = xlHide
    gtypAppSettings.intDisplayDrawingObjects = ActiveWorkbook.DisplayDrawingObjects

    If blnDebug Then
        With gtypAppSettings
        Debug.Print .blnDisplayFormulas
        Debug.Print .blnDisplayGridlines
        Debug.Print .blnDisplayHeadings
        Debug.Print .blnDisplayOutline
        Debug.Print .blnDisplayZeros
        Debug.Print .blnDisplayHorizontalScrollBar
        Debug.Print .blnDisplayVerticalScrollBar
        Debug.Print .blnDisplayWorkbookTabs
        Debug.Print .blnDisplayFormulaBar
        Debug.Print .blnDisplayStatusBar
        Debug.Print .blnShowWindowsInTaskbar
        Debug.Print .intDisplayCommentIndicator
        Debug.Print .blnDisplayAutomaticPageBreaks
        Debug.Print .intDisplayDrawingObjects
        End With
    End If
    
Exit_RememberApplicationSettings:
    Exit Sub
    
Err_RememberApplicationSettings:
    MsgBox Err.Description, vbInformation, "RememberApplicationSettings Error " & Err
    Resume Next
    
End Sub

Public Sub ResetApplicationSettings()

On Error GoTo Err_ResetApplicationSettings

    With ActiveWindow
        .DisplayFormulas = gtypAppSettings.blnDisplayFormulas
        .DisplayGridlines = gtypAppSettings.blnDisplayGridlines
        .DisplayHeadings = gtypAppSettings.blnDisplayHeadings
        .DisplayOutline = gtypAppSettings.blnDisplayOutline
        .DisplayZeros = gtypAppSettings.blnDisplayZeros
        .DisplayHorizontalScrollBar = gtypAppSettings.blnDisplayHorizontalScrollBar
        .DisplayVerticalScrollBar = gtypAppSettings.blnDisplayVerticalScrollBar
        .DisplayWorkbookTabs = gtypAppSettings.blnDisplayWorkbookTabs
    End With
    '
    With Application
        .DisplayFormulaBar = gtypAppSettings.blnDisplayFormulaBar
        .DisplayStatusBar = gtypAppSettings.blnDisplayStatusBar
        .ShowWindowsInTaskbar = gtypAppSettings.blnShowWindowsInTaskbar
' .DisplayCommentIndicator = xlCommentIndicatorOnly
        '.DisplayCommentIndicator = gtypAppSettings.intDisplayCommentIndicator
    End With
    '
' ActiveSheet.DisplayAutomaticPageBreaks = True
    'ActiveSheet.DisplayAutomaticPageBreaks = gtypAppSettings.blnDisplayAutomaticPageBreaks
' ActiveWorkbook.DisplayDrawingObjects = xlHide
    'ActiveWorkbook.DisplayDrawingObjects = gtypAppSettings.intDisplayDrawingObjects

Exit_ResetApplicationSettings:
    Exit Sub
    
Err_ResetApplicationSettings:
    MsgBox Err.Description, vbInformation, "ResetApplicationSettings Error " & Err
    Resume Next

End Sub

Public Sub PrintTheApplicationSettings()

On Error GoTo Err_PrintTheApplicationSettings

    With ActiveWindow
        'Debug.Print .DisplayFormulas
        'Debug.Print .DisplayGridlines
        'Debug.Print .DisplayHeadings
        'Debug.Print .DisplayOutline
        'Debug.Print .DisplayZeros
        'Debug.Print .DisplayHorizontalScrollBar
        'Debug.Print .DisplayVerticalScrollBar
        'Debug.Print .DisplayWorkbookTabs
    End With
    '
    With Application
        'Debug.Print .DisplayFormulaBar
        'Debug.Print .DisplayStatusBar
        'Debug.Print .ShowWindowsInTaskbar
        'Debug.Print .DisplayCommentIndicator
    End With
    '
    'Debug.Print ActiveSheet.DisplayAutomaticPageBreaks
    'Debug.Print ActiveWorkbook.DisplayDrawingObjects

Exit_PrintTheApplicationSettings:
    Exit Sub
    
Err_PrintTheApplicationSettings:
    MsgBox Err.Description, vbInformation, "PrintTheApplicationSettings Error " & Err
    Resume Next

End Sub

Public Sub TotalReset()

    UnprotectWorkbook
    UnprotectAllSheets
    gblnUnsecured = True
    ' Reset Commandbars
    Application.CommandBars("Worksheet Menu Bar").Reset
    Application.CommandBars("Standard").Visible = True
    Application.CommandBars("Formatting").Visible = True
    Application.CommandBars("Ply").Enabled = True              ' Enable View Code Popup
    Application.CommandBars("Toolbar List").Enabled = True

End Sub


