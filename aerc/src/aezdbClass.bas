Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Ref: http://stackoverflow.com/questions/17835107/can-a-mobile-service-server-script-schedule-access-other-sql-azure-databases
' Ref: http://blogs.office.com/b/microsoft-access/archive/2012/07/20/introducing-access-2013-.aspx
' Ref: http://www.petri.co.il/build-windows-server-2012-r2-domain-controller-windows-azure-ip-address-virtual-network.htm
' Ref: http://blog.gluwer.com/2013/07/windows-azure-websites-and-nodejs-the-setup/#
' *** Ref: http://blogs.office.com/b/microsoft-access/archive/2011/04/08/power-tip-improve-the-security-of-database-connections.aspx ***

' Ref: http://www.di-mgt.com.au/cl_Simple.html
' =======================================================================
' Author:   Peter F. Ennis
' Date:     December 26, 2013
' Comment:  Create class for slq azure database management
'           Used aegitClass as template
' =======================================================================

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Private Const THE_CLASS_NAME As String = "aezdbClass"
Private Const THE_VERSION As String = "0.0.3"
Private Const THE_VERSION_DATE As String = "July 4, 2014"
'
' SQL Azure Database
Private miSN        ' "my server name"
Private miUN        ' "my user name"
Private miPW        ' "my password"
Private miDB        ' "my database"
Private miNC        ' "my ssnc"
'
Private aezdbMsg As String
'

Private Sub Class_Initialize()
    ' Ref: http://www.cadalyst.com/cad/autocad/programming-with-class-part-1-5050
    ' Ref: http://www.bigresource.com/Tracker/Track-vb-cyJ1aJEyKj/
    ' Ref: http://stackoverflow.com/questions/1731052/is-there-a-way-to-overload-the-constructor-initialize-procedure-for-a-class-in

    aezdbMsg = "Azure SQL Database Management"

    miSN = "theServer"
    miUN = "theUser"
    miPW = "thePassword"
    miDB = "theDatabase"
    miNC = "theSSNC"

    Debug.Print "Class_Initialize"
    Debug.Print , aezdbMsg

End Sub

Private Sub Class_Terminate()
    Debug.Print
    Debug.Print "Class_Terminate"
    Debug.Print , THE_CLASS_NAME & " VERSION: " & THE_VERSION
    Debug.Print , THE_CLASS_NAME & " VERSION_DATE: " & THE_VERSION_DATE
End Sub

Property Get TheGetMsg(Optional DebugTheCode As Variant) As Boolean
    If IsMissing(DebugTheCode) Then
        Debug.Print "Get TheGetMsg"
        Debug.Print , "DebugTheCode IS missing so no parameter is passed to faeMsg"
        Debug.Print , "DEBUGGING IS OFF"
        TheGetMsg = faeMsg
    Else
        Debug.Print "Get TheGetMsg"
        Debug.Print , "DebugTheCode IS NOT missing so a variant parameter is passed to faeMsg"
        Debug.Print , "DEBUGGING TURNED ON"
        TheGetMsg = faeMsg(DebugTheCode)
    End If
End Property

Property Let TheLetMsg(ByVal strTheLetMsg As String)
    ' Ref: http://www.techrepublic.com/article/build-your-skills-using-class-modules-in-an-access-database-solution/5031814
    ' Ref: http://www.utteraccess.com/wiki/index.php/Classes
    aezdbMsg = strTheLetMsg
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

Property Let SN(SN As String)
    miSN = SN
End Property

Property Let UID(UID As String)
    miUN = UID
End Property

Property Let PWD(PWD As String)
    miPW = PWD
End Property

Property Let db(strDb As String)
    miDB = strDb
End Property

Property Let NC(NC As String)
    miNC = NC
End Property

Property Get InitializeConnection(strDbName As String) As Boolean
    ' Ref: http://blogs.office.com/b/microsoft-access/archive/2011/04/08/power-tip-improve-the-security-of-database-connections.aspx
    ' Description:  Should be called in the application's startup to ensure that Access has a cached connection for all other ODBC objects' use
    Debug.Print "InitializeConnection"

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim qdf As DAO.QueryDef
    'Dim rst As DAO.Recordset
    Dim cnn As String

    cnn = miConSafe(strDbName)

    Set dbs = DBEngine(0)(0)
    Set qdf = dbs.CreateQueryDef("")
    
    With qdf
        .Connect = cnn & _
            "UID=" & miUN & "@" & miSN & ";" & _
            "PWD=" & miPW
        .sql = "SELECT CURRENT_USER();"
        'Set rst = .OpenRecordset(dbOpenSnapshot, dbSQLPassThrough)
        Debug.Print , ".Connect=" & .Connect
    End With
    InitializeConnection = True

PROC_EXIT:
    On Error Resume Next
    'Set rst = Nothing
    Set qdf = Nothing
    Set dbs = Nothing
    Debug.Print "InitializeConnection Exit"
    Exit Property

PROC_ERR:
    MsgBox "Erl=" & Erl & " " & Err.Description & " (" & Err.Number & ") encountered", _
        vbOKOnly + vbCritical, "InitializeConnection"
    TestODBCErr "InitializeConnection"
    InitializeConnection = False
    Resume PROC_EXIT

End Property

Property Get CreateLogin(strLogin As String, strPass As String) As Boolean
    Dim sql As String
    Debug.Print "A: Connecting to the master db"
    sql = "CREATE LOGIN " & strLogin & " WITH password = '" & strPass & "';"
    Debug.Print "CreateLogin sql=" & sql
    CreateLogin = ExecSql(sql, "master")
End Property

Private Function Pause(NumberOfSeconds As Variant)
    ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?p=952355

    On Error GoTo PROC_ERR

    Dim PauseTime As Variant, Start As Variant

    PauseTime = NumberOfSeconds
    Start = Timer
    Do While Timer < Start + PauseTime
        Sleep 100
        DoEvents
    Loop

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Pause of Class " & THE_CLASS_NAME
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
    MsgBox "erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure WaitSeconds of Class " & THE_CLASS_NAME
    Resume PROC_EXIT

End Sub

Private Function FileLocked(strFileName As String) As Boolean
    ' Ref: http://support.microsoft.com/kb/209189
    On Error Resume Next
    ' If the file is already opened by another process,
    ' and the specified Type of access is not allowed,
    ' the Open operation fails and an error occurs.
    Open strFileName For Binary Access Read Write Lock Read Write As #1
    Close 1
    ' If an error occurs, the document is currently open.
    If Err.Number <> 0 Then
        ' Display the error number and description.
        MsgBox "erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FileLocked of Class " & THE_CLASS_NAME
        FileLocked = True
        Err.Clear
    End If
End Function

Private Sub KillProperly(Killfile As String)
    ' Ref: http://word.mvps.org/faqs/macrosvba/DeleteFiles.htm

    On Error GoTo PROC_ERR

    If Len(Dir(Killfile)) > 0 Then
        SetAttr Killfile, vbNormal
        Kill Killfile
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure KillProperly of Class " & THE_CLASS_NAME
    Resume PROC_EXIT

End Sub

Private Function FolderExists(strPath As String) As Boolean
    ' Ref: http://allenbrowne.com/func-11.html
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Private Function aeExists(strAccObjType As String, _
    strAccObjName As String, Optional varDebug As Variant) As Boolean
    ' Ref: http://vbabuff.blogspot.com/2010/03/does-access-object-exists.html
    '
    ' ====================================================================
    ' Author:     Peter F. Ennis
    ' Date:       February 18, 2011
    ' Comment:    Return True if the object exists
    ' Parameters: strAccObjType: "Tables", "Queries", "Forms",
    '                            "Reports", "Macros", "Modules"
    '             strAccObjName: The name of the object
    ' ====================================================================

    Dim objType As Object
    Dim obj As Variant
    Dim blnDebug As Boolean

    On Error GoTo PROC_ERR

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
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class " & THE_CLASS_NAME
    If blnDebug Then Debug.Print ">>>Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeExists of Class " & THE_CLASS_NAME
    aeExists = False
    Resume PROC_EXIT

End Function

' ==================================================
' aezdb Routines
' ==================================================

Private Function faeMsg(Optional varDebug As Variant) As Boolean
    Dim blnDebug As Boolean
    If IsMissing(varDebug) Then
        blnDebug = False
        Debug.Print , "varDebug IS missing so blnDebug of faeMsg is set to False"
        Debug.Print , "DEBUGGING IS OFF"
        faeMsg = True
    Else
        blnDebug = True
        Debug.Print , "varDebug IS NOT missing so blnDebug of faeMsg is set to True"
        Debug.Print , "NOW DEBUGGING..."
        Debug.Print , aezdbMsg
        faeMsg = True
    End If
End Function

Private Function TestODBCErr(strProc As String) As Boolean
    ' Ref: http://www.utteraccess.com/forum/Runtime-Error-3146-t1993453.html&p=2285193

    Dim errX As Variant
    If Errors.Count > 1 Then
        For Each errX In DAO.Errors
            Debug.Print "TestODBCErr " & strProc & " - ODBC Error: errX=" & errX.Number & vbCrLf & vbTab & errX.Description
        Next errX
    Else
        Debug.Print "TestODBCErr " & strProc & " - VBA Error: Err=" & Err.Number & vbCrLf & vbTab & Err.Description
    End If

End Function

Private Function ExecSql(strSQL As String, strDbName As String) As Boolean
    ' Ref: http://msdn.microsoft.com/en-us/library/office/ff195966.aspx

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim qdf As DAO.QueryDef

    ExecSql = False

    Set dbs = CurrentDb

    Set qdf = dbs.CreateQueryDef("")

    qdf.Connect = miConSafe(strDbName)

    qdf.sql = strSQL
    Debug.Print "strSQL=" & strSQL
    ' ReturnsRecords must be set to False if the SQL does not return records
    qdf.ReturnsRecords = False
    qdf.Execute dbFailOnError

    ' If no errors were raised the query was successfully executed
    ExecSql = True

PROC_EXIT:
    On Error Resume Next
    Set qdf = Nothing
    Set dbs = Nothing
    Exit Function

PROC_ERR:
    If Err.Number = 3151 Then
        MsgBox "Err.Number = 3151 - Connection failed!" & vbCrLf & Err.Description
    ElseIf Err.Number = 3146 Then
        MsgBox "Err.Number = 3146 - User exists?" & vbCrLf & Err.Description
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & vbCrLf & Err.Description _
            & vbCrLf & "In procedure ExecSql"
        Debug.Print "Erl=" & Erl & " Error " & Err.Number & vbCrLf & Err.Description _
            & vbCrLf & "In procedure ExecSql"
    End If
    ExecSql = False
    Resume PROC_EXIT

End Function

Private Function miConSafe(strDbName As String) As String

    Debug.Print "miConSafe"
    miConSafe = "ODBC;" _
        & "DRIVER=" & miNC _
        & "SERVER=tcp:" & miSN & ".database.windows.net,1433;" _
        & "DATABASE=" & strDbName & ";" _
        & "Encrypt=Yes;" _
        & "Connection Timeout=30;"
    Debug.Print , miConSafe
    Debug.Print "miConSafe Exit"

End Function