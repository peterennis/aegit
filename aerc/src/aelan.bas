Option Compare Database
Option Explicit
' Ref: http://www.vbforums.com/showthread.php?279162-Using-Winsock-to-Connect-with-another-computer

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare PtrSafe Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Private Declare PtrSafe Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare PtrSafe Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Const MAX_WSADescription        As Long = 256
Private Const MAX_WSASYSStatus          As Long = 128
Private Const WS_VERSION_REQD           As Long = &H101
Private Const MAX_COMPUTERNAME_LENGTH   As Long = 31

Private Type HOSTENT
    hName            As Long
    hAliases         As Long
    hAddrType        As Integer
    hLen             As Integer
    hAddrList        As Long
End Type

Private Type WSADATA
    wVersion                                 As Integer
    wHighVersion                             As Integer
    szDescription(0 To MAX_WSADescription)   As Byte
    szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
    wMaxSockets                              As Integer
    wMaxUDPDG                                As Integer
    dwVendorInfo                             As Long
End Type

Public Sub Startup()

    Dim strMsg      As String
    Dim strname     As String

    On Error GoTo ErrHandler:

    strname = ComputerName
    strMsg = "Computer Name:" & vbTab & strname & vbCrLf
    strMsg = strMsg & "IP Address:" & vbTab & GetIPAddress(strname)
    MsgBox strMsg
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbCritical, "Error"

End Sub

Private Property Get ComputerName() As String

    On Error GoTo 0
    Dim dwLen       As Long
    Dim strString   As String

    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String$(dwLen, "X")
    GetComputerName strString, dwLen
    strString = Left$(strString, dwLen)
    ComputerName = strString

End Property

Private Function GetIPAddress(ByVal pstrHost As String) As String

    On Error GoTo 0
    Dim lngIndex        As Integer
    Dim lngHost         As Long
    Dim lngIP           As Long
    Dim bytIP()         As Byte
    Dim strIPAddress    As String
    Dim udtHost         As HOSTENT
    Dim udtSocket       As WSADATA

    If Not WSAStartup(WS_VERSION_REQD, udtSocket) Then
        lngHost = gethostbyname(pstrHost)

        If lngHost = 0 Then
            Err.Raise vbObjectError, , "Unable to locate Server."
        Else
            CopyMemory udtHost, lngHost, Len(udtHost)
            CopyMemory lngIP, udtHost.hAddrList, Len(udtHost.hAddrList)
            ReDim bytIP(udtHost.hLen - 1) As Byte
            CopyMemory bytIP(0), lngIP, udtHost.hLen
            
            For lngIndex = 0 To udtHost.hLen - 1

                If Not (lngIndex = 0) Then
                    strIPAddress = strIPAddress & "."
                End If

                strIPAddress = strIPAddress & bytIP(lngIndex)

            Next
            GetIPAddress = strIPAddress
        End If

        Call WSACleanup
    Else
        Err.Raise vbObjectError, , "Could not start winsock service"
    End If

End Function