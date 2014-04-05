Option Compare Database
Option Explicit

' Ref: http://stackoverflow.com/questions/5062021/use-png-as-custom-ribbon-icon-in-access-2007

'================================================================================
'  Declarations required to load .png's in Ribbon
Private Type GUID
    Data1                   As Long
    Data2                   As Integer
    Data3                   As Integer
    Data4(0 To 7)           As Byte
End Type

Private Type PICTDESC
    Size                        As Long
    Type                        As Long
    hPic                        As Long
    hPal                        As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, _
    inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, bitmap As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal bitmap As Long, _
    hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal image As Long) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, _
    RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
'================================================================================

Public Sub GetRibbonImage(ByVal ctl As Object, ByVal image As Variant)        'IRibbonControl)
    On Error GoTo 0
    Dim Path As String
    Path = Application.CurrentProject.Path & "\Icons\" & ctl.Tag
    Set image = LoadImage(Path)
End Sub

Private Function LoadImage(ByVal strFName As String) As IPicture

    On Error GoTo 0
    Dim uGdiInput As GdiplusStartupInput
    Dim hGdiPlus As Long
    Dim hGdiImage As Long
    Dim hBitmap As Long

    uGdiInput.GdiplusVersion = 1

    If GdiplusStartup(hGdiPlus, uGdiInput) = 0 Then
        If GdipCreateBitmapFromFile(StrPtr(strFName), hGdiImage) = 0 Then
            GdipCreateHBITMAPFromBitmap hGdiImage, hBitmap, 0
            Set LoadImage = ConvertToIPicture(hBitmap)
            GdipDisposeImage hGdiImage
        End If
        GdiplusShutdown hGdiPlus
    End If

End Function

Private Function ConvertToIPicture(ByVal hPic As Long) As IPicture

    On Error GoTo 0
    Dim uPicInfo As PICTDESC
    Dim IID_IDispatch As GUID
    Dim IPic As IPicture

    Const PICTYPE_BITMAP = 1

    With IID_IDispatch
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With

    With uPicInfo
        .Size = Len(uPicInfo)
        .Type = PICTYPE_BITMAP
        .hPic = hPic
        .hPal = 0
    End With

    OleCreatePictureIndirect uPicInfo, IID_IDispatch, True, IPic

    Set ConvertToIPicture = IPic

End Function