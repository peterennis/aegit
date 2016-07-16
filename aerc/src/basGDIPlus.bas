Option Compare Database
Option Explicit


'-------------------------------------------------
'    Picture functions using GDIPlus-API (GDIP)   |
'-------------------------------------------------

'-------------------------------------------------
'   (c) mossSOFT / Sascha Trowitzsch rev. 04/2009 |
'-------------------------------------------------

'- Reference to library "OLE Automation" (stdole) needed!
'- Code work under Office 2007 and Office 2010 x86 and Office 2010 x64 (see *Remark below)

'  rev. 07/2010 (Support for Office 2010 x64)
'  rev. 10/2011 better Timer Support
'  rev. 08/2013 InitGDIP() updated

Public Const GUID_IPicture As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"    'IPicture

'User-defined types: ----------------------------------------------------------------------

Public Enum PicFileType
    pictypeBMP = 1
    pictypeGIF = 2
    pictypePNG = 3
    pictypeJPG = 4
End Enum

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type TSize
    X As Double
    Y As Double
End Type


#If Win64 Then
    Private Type PICTDESC
        cbSizeOfStruct As Long
        PicType As Long
        hImage As LongPtr
        xExt As Long
        yExt As Long
    End Type
    
    Private Type GDIPStartupInput
        GdiplusVersion As Long
        DebugEventCallback As LongPtr
        SuppressBackgroundThread As LongPtr
        SuppressExternalCodecs As LongPtr
    End Type
    
    Private Type EncoderParameter
        UUID As GUID
        NumberOfValues As LongPtr
        Type As LongPtr
        Value As LongPtr
    End Type
    
#Else

    Private Type PICTDESC
        cbSizeOfStruct As Long
        PicType As Long
        hImage As Long
        xExt As Long
        yExt As Long
    End Type
    
    Private Type GDIPStartupInput
        GdiplusVersion As Long
        DebugEventCallback As Long
        SuppressBackgroundThread As Long
        SuppressExternalCodecs As Long
    End Type
    
    Private Type EncoderParameter
        UUID As GUID
        NumberOfValues As Long
        Type As Long
        Value As Long
    End Type
#End If

Private Type EncoderParameters
    Count As Long
    Parameter As EncoderParameter
End Type

#If Win64 Then

    'API-Declarations: ----------------------------------------------------------------------------

    ' G.A.: olepro32 in oleaut32 geändert. Olepro32 ist in x64 nicht verfügbar.
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef lpPictDesc As PICTDESC, ByRef riid As GUID, ByVal fPictureOwnsHandle As LongPtr, ByRef IPic As Object) As Long

    'Retrieve GUID-Type from string :
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, ByRef pclsid As GUID) As Long

    'Memory functions:
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As LongPtr, ByVal dwBytes As LongPtr) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByRef Source As Byte, ByVal Length As LongPtr)

    'Modules API:
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As LongPtr) As Long
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    
    'Timer API:
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) As Long
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

    'OLE-Stream functions :
    Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As LongPtr, ByRef ppstm As Any) As Long
    Private Declare PtrSafe Function GetHGlobalFromStream Lib "ole32.dll" (ByVal pstm As Any, ByRef phglobal As LongPtr) As Long
    
    
    'GDIPlus Flat-API declarations:
    
    'Initialization GDIP:
    Private Declare PtrSafe Function GdiplusStartup Lib "GDIPlus" (ByRef token As LongPtr, ByRef inputbuf As GDIPStartupInput, Optional ByVal outputbuf As Long = 0) As Long
    'Tear down GDIP:
    Private Declare PtrSafe Function GdiplusShutdown Lib "GDIPlus" (ByVal token As LongPtr) As Long
    'Load GDIP-Image from file :
    Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal FileName As LongPtr, BITMAP As LongPtr) As Long
    'Create GDIP- graphical area from Windows-DeviceContext:
    Private Declare PtrSafe Function GdipCreateFromHDC Lib "GDIPlus" (ByVal hdc As LongPtr, ByRef GpGraphics As LongPtr) As Long
    'Delete GDIP graphical area :
    Private Declare PtrSafe Function GdipDeleteGraphics Lib "GDIPlus" (ByVal graphics As LongPtr) As Long
    'Copy GDIP-Image to graphical area:
    Private Declare PtrSafe Function GdipDrawImageRect Lib "GDIPlus" (ByVal graphics As LongPtr, ByVal Image As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
    'Clear allocated bitmap memory from GDIP :
    Private Declare PtrSafe Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As LongPtr) As Long
    'Retrieve windows bitmap handle from GDIP-Image:
    Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal BITMAP As LongPtr, ByRef hbmReturn As LongPtr, ByVal background As LongPtr) As Long
    'Retrieve Windows-Icon-Handle from GDIP-Image:
    Public Declare PtrSafe Function GdipCreateHICONFromBitmap Lib "GDIPlus" (ByVal BITMAP As LongPtr, ByRef hbmReturn As LongPtr) As Long
    'Scaling GDIP-Image size:
    Private Declare PtrSafe Function GdipGetImageThumbnail Lib "GDIPlus" (ByVal Image As LongPtr, ByVal thumbWidth As LongPtr, ByVal thumbHeight As LongPtr, ByRef thumbImage As LongPtr, Optional ByVal callback As LongPtr = 0, Optional ByVal callbackData As LongPtr = 0) As Long
    'Retrieve GDIP-Image from Windows-Bitmap-Handle:
    Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As LongPtr, ByVal hPal As LongPtr, ByRef BITMAP As LongPtr) As Long
    'Retrieve GDIP-Image from Windows-Icon-Handle:
    Private Declare PtrSafe Function GdipCreateBitmapFromHICON Lib "GDIPlus" (ByVal hicon As LongPtr, ByRef BITMAP As LongPtr) As Long
    'Retrieve width of a GDIP-Image (Pixel):
    Private Declare PtrSafe Function GdipGetImageWidth Lib "GDIPlus" (ByVal Image As LongPtr, ByRef Width As LongPtr) As Long
    'Retrieve height of a GDIP-Image (Pixel):
    Private Declare PtrSafe Function GdipGetImageHeight Lib "GDIPlus" (ByVal Image As LongPtr, ByRef Height As LongPtr) As Long
    'Save GDIP-Image to file in seletable format:
    Private Declare PtrSafe Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As LongPtr, ByVal FileName As LongPtr, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
    'Save GDIP-Image in OLE-Stream with seletable format:
    Private Declare PtrSafe Function GdipSaveImageToStream Lib "GDIPlus" (ByVal Image As LongPtr, ByVal stream As IUnknown, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
    'Retrieve GDIP-Image from OLE-Stream-Object:
    Private Declare PtrSafe Function GdipLoadImageFromStream Lib "GDIPlus" (ByVal stream As IUnknown, ByRef Image As LongPtr) As Long
    'Create a gdip image from scratch
    Private Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "GDIPlus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByRef scan0 As Any, ByRef BITMAP As Long) As Long
    'Get the DC of an gdip image
    Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "GDIPlus" (ByVal Image As LongPtr, ByRef graphics As LongPtr) As Long
    'Blit the contents of an gdip image to another image DC using positioning
    Private Declare PtrSafe Function GdipDrawImageRectRectI Lib "GDIPlus" (ByVal graphics As LongPtr, ByVal Image As LongPtr, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long

    
    '-----------------------------------------------------------------------------------------
    'Global module variable:
    Private lGDIP As LongPtr
    '-----------------------------------------------------------------------------------------
    
    Private TempVarGDIPlus As LongPtr
#Else

    'API-Declarations: ----------------------------------------------------------------------------
    
    'Convert a windows bitmap to OLE-Picture :
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PICTDESC, riid As GUID, ByVal fPictureOwnsHandle As Long, IPic As Object) As Long
    'Retrieve GUID-Type from string :
    Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As GUID) As Long
    
    'Memory functions:
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByRef Source As Byte, ByVal Length As Long)
    
    'Modules API:
    Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
    Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    
    'Timer API:
    Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
    
    
    'OLE-Stream functions :
    Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
    Private Declare Function GetHGlobalFromStream Lib "ole32.dll" (ByVal pstm As Any, ByRef phglobal As Long) As Long
    
    'GDIPlus Flat-API declarations:
    
    'Initialization GDIP:
    Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GDIPStartupInput, Optional ByVal outputbuf As Long = 0) As Long
    'Tear down GDIP:
    Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
    'Load GDIP-Image from file :
    Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal FileName As Long, BITMAP As Long) As Long
    'Create GDIP- graphical area from Windows-DeviceContext:
    Private Declare Function GdipCreateFromHDC Lib "GDIPlus" (ByVal hdc As Long, GpGraphics As Long) As Long
    'Delete GDIP graphical area :
    Private Declare Function GdipDeleteGraphics Lib "GDIPlus" (ByVal graphics As Long) As Long
    'Copy GDIP-Image to graphical area:
    Private Declare Function GdipDrawImageRect Lib "GDIPlus" (ByVal graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
    'Clear allocated bitmap memory from GDIP :
    Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
    'Retrieve windows bitmap handle from GDIP-Image:
    Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal BITMAP As Long, hbmReturn As Long, ByVal background As Long) As Long
    'Retrieve Windows-Icon-Handle from GDIP-Image:
    Public Declare Function GdipCreateHICONFromBitmap Lib "GDIPlus" (ByVal BITMAP As Long, hbmReturn As Long) As Long
    'Scaling GDIP-Image size:
    Private Declare Function GdipGetImageThumbnail Lib "GDIPlus" (ByVal Image As Long, ByVal thumbWidth As Long, ByVal thumbHeight As Long, thumbImage As Long, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
    'Retrieve GDIP-Image from Windows-Bitmap-Handle:
    Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
    'Retrieve GDIP-Image from Windows-Icon-Handle:
    Private Declare Function GdipCreateBitmapFromHICON Lib "GDIPlus" (ByVal hicon As Long, BITMAP As Long) As Long
    'Retrieve width of a GDIP-Image (Pixel):
    Private Declare Function GdipGetImageWidth Lib "GDIPlus" (ByVal Image As Long, Width As Long) As Long
    'Retrieve height of a GDIP-Image (Pixel):
    Private Declare Function GdipGetImageHeight Lib "GDIPlus" (ByVal Image As Long, Height As Long) As Long
    'Save GDIP-Image to file in seletable format:
    Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
    'Save GDIP-Image in OLE-Stream with seletable format:
    Private Declare Function GdipSaveImageToStream Lib "GDIPlus" (ByVal Image As Long, ByVal stream As IUnknown, clsidEncoder As GUID, encoderParams As Any) As Long
    'Retrieve GDIP-Image from OLE-Stream-Object:
    Private Declare Function GdipLoadImageFromStream Lib "GDIPlus" (ByVal stream As IUnknown, Image As Long) As Long
    'Create a gdip image from scratch
    Private Declare Function GdipCreateBitmapFromScan0 Lib "GDIPlus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, BITMAP As Long) As Long
    'Get the DC of an gdip image
    Private Declare Function GdipGetImageGraphicsContext Lib "GDIPlus" (ByVal Image As Long, graphics As Long) As Long
    'Blit the contents of an gdip image to another image DC using positioning
    Private Declare Function GdipDrawImageRectRectI Lib "GDIPlus" (ByVal graphics As Long, ByVal Image As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long

    '-----------------------------------------------------------------------------------------
    'Global module variable:
    Private lGDIP As Long

#End If

Private tVarTimer() As Long
Private lCounter    As Long
Private bSharedLoad As Boolean


'Initialize GDI+
Private Function InitGDIP() As Boolean

    On Error GoTo 0

    Dim TGDP As GDIPStartupInput
    Dim hMod As Long

    'Debug.Print Now(), "InitGDIP"

    If lGDIP = 0 Then
        #If Win64 Then
            If TempVarGDIPlus = 0 Then 'If lGDIP is broken due to unhandled errors restore it from the Tempvars collection
                TGDP.GdiplusVersion = 1
                hMod = GetModuleHandle("GDIPlus.dll")    'ogl.dll not yet loaded?
                If hMod = 0 Then
                    hMod = LoadLibrary("GDIPlus.dll")
                    bSharedLoad = False
                Else
                    bSharedLoad = True
                End If
                GdiplusStartup lGDIP, TGDP 'Get a personal instance of GDIPlus
                TempVarGDIPlus = lGDIP
                
            Else
                lGDIP = TempVarGDIPlus
            End If
            
        #Else
            If IsNull(TempVars("GDIPlusHandle")) Then
                'Debug.Print Now(), "InitGDIP, start INIT"
                TGDP.GdiplusVersion = 1
                hMod = GetModuleHandle("GDIPlus.dll")
                If hMod = 0 Then
                    hMod = LoadLibrary("GDIPlus.dll")
                    bSharedLoad = False
                Else
                    bSharedLoad = True
                End If
                GdiplusStartup lGDIP, TGDP
                TempVars("GDIPlusHandle") = lGDIP
            Else
                lGDIP = TempVars("GDIPlusHandle")
            End If
            
        #End If
        
    End If
    
    InitGDIP = (lGDIP <> 0)
    'Debug.Print Now(), "InitGDIP End", lGDIP
    
    AutoShutDown
    
End Function

'Clear GDI+
Private Sub ShutDownGDIP()
    'Debug.Print Now(), "ShutDownGDIP"

    On Error GoTo 0

    If lGDIP <> 0 Then
    
        Dim lngDummy As Long
        Dim lngDummyTimer As Long
        For lngDummy = 0 To lCounter - 1
            lngDummyTimer = tVarTimer(lngDummy)
                        
            If lngDummyTimer <> 0 Then
                If KillTimer(0&, CLng(lngDummyTimer)) Then
                    'Debug.Print Now(), "ShutDownGDIP, Timer " & CLng(lngDummyTimer) & " KILLED"
                    tVarTimer(lngDummy) = 0
                End If
            
            End If
        Next
            
        GdiplusShutdown lGDIP
        lGDIP = 0
        
        #If Win64 Then
            TempVarGDIPlus = 0
        #Else
            TempVars("GDIPlusHandle") = Null
        #End If
        
        If Not bSharedLoad Then FreeLibrary GetModuleHandle("GDIPlus.dll")
        
    End If
       
End Sub

'Scheduled ShutDown of GDI+ handle to avoid memory leaks
Private Sub AutoShutDown()
    'Set to 5 seconds for next shutdown
    'That's IMO appropriate for looped routines  - but configure for your own purposes
 
    On Error GoTo 0
 
    If lGDIP <> 0 Then
        ReDim Preserve tVarTimer(lCounter)
        tVarTimer(lCounter) = SetTimer(0&, 0&, 5000, AddressOf TimerProc)
        'Debug.Print Now(), "AutoShutDown SET", tVarTimer(lCounter), lCounter
    End If
    'Debug.Print Now(), "AutoShutDown", tVarTimer(lCounter), lCounter
    lCounter = lCounter + 1
    
End Sub

'Callback for AutoShutDown
Private Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    'Debug.Print Now(), "TimerProc Start"
    On Error GoTo 0
    ShutDownGDIP
    'Debug.Print Now(), "TimerProc End"
End Sub

'Load image file with GDIP
'It's equivalent to the method LoadPicture() in OLE-Automation library (stdole2.tlb)
'Allowed format: bmp, gif, jp(e)g, tif, png, wmf, emf, ico
Private Function LoadPictureGDIP(ByRef sFileName As String) As StdPicture
    #If Win64 Then
        Dim hBmp As LongPtr
        Dim hPic As LongPtr
    #Else
        Dim hBmp As Long
        Dim hPic As Long
    #End If
    
    On Error GoTo 0
    
    If Not InitGDIP Then Exit Function
    
        If GdipCreateBitmapFromFile(StrPtr(sFileName), hPic) = 0 Then
   
        GdipCreateHBITMAPFromBitmap hPic, hBmp, 0&
    
        If hBmp <> 0 Then
            Set LoadPictureGDIP = BitmapToPicture(hBmp)
            GdipDisposeImage hPic
        End If
    End If

End Function

'Create an OLE-Picture from Byte-Array PicBin()
Public Function ArrayToPicture(ByRef PicBin() As Byte) As Picture

    On Error GoTo 0

    Dim IStm As IUnknown
    #If Win64 Then
        Dim lBitmap As LongPtr
        Dim hBmp As LongPtr
    #Else
        Dim lBitmap As Long
        Dim hBmp As Long
    #End If

    Dim ret As Long

    'Debug.Print Now(), "ArrayToPicture"
    
    If Not InitGDIP Then
        Debug.Print "Exit"
        Exit Function
    End If

    ret = CreateStreamOnHGlobal(VarPtr(PicBin(0)), 0, IStm)  'Create stream from memory stack

    If ret = 0 Then    'OK, start GDIP :
        'Convert stream to GDIP-Image :
        ret = GdipLoadImageFromStream(IStm, lBitmap)
        If ret = 0 Then
            'Get Windows-Bitmap from GDIP-Image:
            GdipCreateHBITMAPFromBitmap lBitmap, hBmp, 0&
            If hBmp <> 0 Then
                'Convert bitmap to picture object :
                Set ArrayToPicture = BitmapToPicture(hBmp)
            End If
        End If
        'Clear memory ...
        GdipDisposeImage lBitmap
    End If

End Function

#If Win64 Then
    'Help function to get a OLE-Picture from Windows-Bitmap-Handle
    'If bIsIcon = TRUE, an Icon-Handle is committed
    Private Function BitmapToPicture(ByVal hBmp As LongPtr, Optional ByRef bIsIcon As Boolean = False) As StdPicture

        On Error GoTo 0

        Dim TPicConv As PICTDESC
        Dim UID As GUID

        With TPicConv
            If bIsIcon Then
                .cbSizeOfStruct = 16
                .PicType = 3    'PicType Icon
            Else
                .cbSizeOfStruct = Len(TPicConv)
                .PicType = 1    'PicType Bitmap
            End If
            .hImage = hBmp
        End With

        CLSIDFromString StrPtr(GUID_IPicture), UID
        OleCreatePictureIndirect TPicConv, UID, True, BitmapToPicture

    End Function
#Else
    Private Function BitmapToPicture(ByVal hBmp As Long, Optional bIsIcon As Boolean = False) As StdPicture
        
        On Error GoTo 0
        
        Dim TPicConv As PICTDESC, UID As GUID
    
        With TPicConv
            If bIsIcon Then
                .cbSizeOfStruct = 16
                .PicType = 3    'PicType Icon
            Else
                .cbSizeOfStruct = Len(TPicConv)
                .PicType = 1    'PicType Bitmap
            End If
            .hImage = hBmp
        End With
    
        CLSIDFromString StrPtr(GUID_IPicture), UID
        OleCreatePictureIndirect TPicConv, UID, True, BitmapToPicture
    
    End Function
#End If