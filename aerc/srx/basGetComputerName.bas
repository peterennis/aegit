Attribute VB_Name = "basGetComputerName"
Option Explicit
'
' Reference:
' http://www.citilink.com/~jgarrick/vbasic/
'
Private Declare Function w32_GetComputerName Lib "kernel32" Alias "GetComputerNameA" ( _
  ByVal lpBuffer As String, _
  nSize As Long) As Long

Private Declare Function w32_SetComputerName Lib "kernel32" Alias "SetComputerNameA" ( _
  ByVal lpComputerName As String) As Long

Public Function GetComputerName(rpsComputerName) As Boolean
'    Created By: Joe Garrick 01/13/97 5:03 PM
'*******************************************************************************
'       Purpose: Get the local computer name
'    Parameters: (Output Only)
'                   rpsComputerName - return buffer for the computer name
'       Returns: True if successful
'         Notes: See below
'*******************************************************************************

  Dim sComputerName As String
  Dim lComputerNameLen As Long
  
  Dim lResult As Long
  Dim fRV As Boolean
  Dim i As Integer

  lComputerNameLen = 256
  sComputerName = Space(lComputerNameLen)
  
  lResult = w32_GetComputerName(sComputerName, lComputerNameLen)
  If lResult <> 0 Then
    rpsComputerName = Left$(sComputerName, lComputerNameLen)
    fRV = True
  Else
    fRV = False
  End If

  GetComputerName = fRV

End Function

' The result is converted from the long returned by the API into a Boolean to make using the
' function easier. However, this is a reliable enough function that in most cases the return
' value could be ignored or discarded.
' The SetComputerName function is even simpler. You simply call the API function, passing the
' desired value as a string.

Public Function SetComputerName(psComputerName As String) As Boolean
'    Created By: Joe Garrick 01/13/97 5:03 PM
'*******************************************************************************
'       Purpose: Set the local computer name
'    Parameters: (Input Only)
'                   psComputerName - the name to assign
'       Returns: True if successful
'         Notes: See below
'*******************************************************************************

  Dim lResult As Long
  Dim fRV As Boolean
  
  lResult = w32_SetComputerName(psComputerName)
  If lResult <> 0 Then
    fRV = True
  Else
    fRV = False
  End If
  
  SetComputerName = fRV
  
End Function

' Again, the return value is converted to a Boolean for simplicity.
' Like the GetComputerName function, in most cases the return value could probably be ignored.
' Notes
' Check the Win32 SDK documentation for additional information on these functions.
' As always, save your work often when dealing with the API - it's exceptionally unforgiving of errors.
' In GetComputerName, the name parameter is preallocated to 256 bytes, which is (as far as I know)
' more than long enough to hold any valid computer name on a Windows platform. However, if the
' function should fail because a particular OS allows computer names longer than this,
' just increase the original buffer size.
' Both of these functions are exceptionally reliable.
' If you're confident enough in the return values or can live with the possibility of failure
' in the function, the GetComputerName function could be converted to simply return the name
' (rather than the Boolean result code) as the return value.
' Remember that after assigning the name you'll need to reboot the machine to have the new name recognized.
' Check out the ExitWindowsEx function to restart the system.
' Windows 95
' If the name for SetComputerName contains one or more characters that are outside the standard character set,
' those characters are coerced into standard characters.
' Windows NT
' If the name for SetComputerName contains one or more characters that are outside the standard character set,
' SetComputerName returns ERROR_INVALID_PARAMETER. The VB wrapper will translate this into a
' return value of False. Unlike Win95, SetComputerName does not coerce the characters outside
' the standard set.
' You must have Administrator authority to assign a computer name under Windows NT.
' The standard character set for the computer name includes letters, numbers, and the following symbols:
' ! @ # $ % ^ & ' ) ( . - _ { } ~ .


