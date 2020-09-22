Attribute VB_Name = "modGDIPlus"
Option Explicit

'+++++++++++++++++++++++++++++++++++++++++ GDI RESIZE +++++++++++++++++++++++++++++++++++++++++++

Private Type PICTDESC
   Size     As Long
   Type     As Long
   hBmp     As Long
   hPal     As Long
   Reserved As Long
End Type

Private Type PWMFRect16
    Left   As Integer
    Top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Private Type wmfPlaceableFileHeader
    Key         As Long
    hMf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    Reserved    As Long
    CheckSum    As Integer
End Type

' GDI Functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

' GDI+ functions
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal image As Long, Height As Long) As Long
Private Declare Function GdipCreateMetafileFromWmf Lib "gdiplus.dll" (ByVal hWmf As Long, ByVal deleteWmf As Long, WmfHeader As wmfPlaceableFileHeader, Metafile As Long) As Long
Private Declare Function GdipCreateMetafileFromEmf Lib "gdiplus.dll" (ByVal hEmf As Long, ByVal deleteEmf As Long, Metafile As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus.dll" (ByVal hIcon As Long, GpBitmap As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long

' GDI and GDI+ constants
Private Const PLANES = 14            '  Number of planes
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PATCOPY = &HF00021     '  (DWORD) dest = pattern
Private Const PICTYPE_BITMAP = 1     '  Bitmap type
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2

'++++++++++++++++++++++++++++++++++++++++ GDI SAVE BITMAP +++++++++++++++++++++++++++++++++++++++++++

Public Const GdiplusVersion     As Long = 1
Private Const CP_ACP            As Long = 0

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type ImageCodecInfo
   ClassID As GUID
   FormatID As GUID
   CodecName As Long
   DllName As Long
   FormatDescription As Long
   FilenameExtension As Long
   MimeType As Long
   flags As Long
   Version As Long
   SigCount As Long
   SigSize As Long
   SigPattern As Long
   SigMask As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GdiplusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As Long, ByVal FileName As Long, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hbm As Long, ByVal hPal As Long, nBitmap As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (ByRef numEncoders As Long, ByRef Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, ByRef encoders As Any) As Long

'++++++++++++++++++++++++++++++++++++++++ GDI SAVE BITMAP +++++++++++++++++++++++++++++++++++++++++++

'init GDI Plus
Public Function InitGDIPlus() As Long
    
    On Error Resume Next
    
    Dim token    As Long
    Dim gdipInit As GdiplusStartupInput
    
    gdipInit.GdiplusVersion = 1
    GdiplusStartup token, gdipInit, ByVal 0&
    InitGDIPlus = token
    
End Function

'free GDI Plus
Public Sub FreeGDIPlus(token As Long)
    
    On Error Resume Next
    
    GdiplusShutdown token
    
End Sub

Public Function SavePictureFromHDC(ByVal hBitmap As Long, ByVal sFileName As String) As Boolean
    
    On Error Resume Next
    
    Dim lBitmap As Long
    Dim PicEncoder As GUID
    Dim sID As String
    
    ' Use file name extention to determine,
    ' what format we want to save the file in.
    Select Case LCase$(Right$(sFileName, 4))
        Case ".png"
            sID = "image/png"
        Case ".gif"
            sID = "image/gif"
        Case ".jpg"
            sID = "image/jpeg"
        Case "jpeg"
            sID = "image/jpeg"
        Case ".tif"
            sID = "image/tiff"
        Case "tiff"
            sID = "image/tiff"
        Case ".bmp"
            sID = "image/bmp"
        Case Else
            Exit Function
    End Select
    
    If GdipCreateBitmapFromHBITMAP(hBitmap, 0&, lBitmap) = 0 Then
        If GetEncoderClsID(sID, PicEncoder) = True Then
            SavePictureFromHDC = (GdipSaveImageToFile(lBitmap, StrPtr(sFileName), PicEncoder, ByVal 0) = 0)
        End If
        GdipDisposeImage lBitmap
    End If
    
End Function

Private Function GetEncoderClsID(strMimeType As String, ClassID As GUID) As Boolean
    
    On Error Resume Next
    
    Dim Num As Long
    Dim Size As Long
    Dim imgCodecInfo() As ImageCodecInfo
    Dim lval As Long
    Dim Buffer() As Byte

    GdipGetImageEncodersSize Num, Size
    If Size Then
        ReDim imgCodecInfo(Num) As ImageCodecInfo
        ReDim Buffer(Size) As Byte

        GdipGetImageEncoders Num, Size, Buffer(0)
        CopyMemory imgCodecInfo(0), Buffer(0), (Len(imgCodecInfo(0)) * Num)

        For lval = 0 To Num - 1
            'image/bmp,image/jpeg,image/gif,image/tiff,image/png
            If StrComp(GetStrFromPtrW(imgCodecInfo(lval).MimeType), strMimeType, vbTextCompare) = 0 Then
                ClassID = imgCodecInfo(lval).ClassID
                GetEncoderClsID = True
                Exit For
            End If
        Next
        Erase imgCodecInfo
        Erase Buffer
    End If
    
End Function

Private Function GetStrFromPtrW(lpszW As Long) As String
    
    On Error Resume Next
    
    Dim sRV As String

    sRV = String$(lstrlenW(ByVal lpszW) * 2, vbNullChar)
    WideCharToMultiByte CP_ACP, 0, ByVal lpszW, -1, ByVal sRV, Len(sRV), 0, 0
    GetStrFromPtrW = Left$(sRV, lstrlenW(StrPtr(sRV)))
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++ GDI RESIZE +++++++++++++++++++++++++++++++++++++++++

' Loads the picture (optionally resized)
Public Function LoadPictureGDIPlus(PicFile As String, Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    
    On Error Resume Next
    
    Dim hDC     As Long
    Dim hBitmap As Long
    Dim Img     As Long
        
    ' Load the image
    If GdipLoadImageFromFile(StrPtr(PicFile), Img) <> 0 Then
        Err.Raise 999, "GDI+ Module", "Error loading picture " & PicFile
        Exit Function
    End If
    
    ' Calculate picture's width and height if not specified
    If Width = -1 Or Height = -1 Then
        GdipGetImageWidth Img, Width
        GdipGetImageHeight Img, Height
    End If
    
    ' Initialise the hDC
    InitDC hDC, hBitmap, BackColor, Width, Height

    ' Resize the picture
    gdipResize Img, hDC, Width, Height, RetainRatio
    GdipDisposeImage Img
    
    ' Get the bitmap back
    GetBitmap hDC, hBitmap

    ' Create the picture
    Set LoadPictureGDIPlus = CreatePicture(hBitmap)
    
End Function

' Initialises the hDC to draw
Private Sub InitDC(hDC As Long, hBitmap As Long, BackColor As Long, Width As Long, Height As Long)
    
    On Error Resume Next
    
    Dim hBrush As Long
        
    ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hDC = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hDC, hBitmap)
    hBrush = CreateSolidBrush(BackColor)
    hBrush = SelectObject(hDC, hBrush)
    PatBlt hDC, 0, 0, Width, Height, PATCOPY
    DeleteObject SelectObject(hDC, hBrush)
    
End Sub

' Resize the picture using GDI plus
Private Sub gdipResize(Img As Long, hDC As Long, Width As Long, Height As Long, Optional RetainRatio As Boolean = False)
    
    On Error Resume Next
    
    Dim Graphics   As Long      ' Graphics Object Pointer
    Dim OrWidth    As Long      ' Original Image Width
    Dim OrHeight   As Long      ' Original Image Height
    Dim OrRatio    As Double    ' Original Image Ratio
    Dim DesRatio   As Double    ' Destination rect Ratio
    Dim DestX      As Long      ' Destination image X
    Dim DestY      As Long      ' Destination image Y
    Dim DestWidth  As Long      ' Destination image Width
    Dim DestHeight As Long      ' Destination image Height
    
    GdipCreateFromHDC hDC, Graphics
    GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic
    
    If RetainRatio Then
        GdipGetImageWidth Img, OrWidth
        GdipGetImageHeight Img, OrHeight
        
        OrRatio = OrWidth / OrHeight
        DesRatio = Width / Height
        
        ' Calculate destination coordinates
        DestWidth = IIf(DesRatio < OrRatio, Width, Height * OrRatio)
        DestHeight = IIf(DesRatio < OrRatio, Width / OrRatio, Height)
        DestX = (Width - DestWidth) / 2
        DestY = (Height - DestHeight) / 2

        GdipDrawImageRectRectI Graphics, Img, DestX, DestY, DestWidth, DestHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
    Else
        GdipDrawImageRectI Graphics, Img, 0, 0, Width, Height
    End If
    
    GdipDeleteGraphics Graphics
    
End Sub

' Replaces the old bitmap of the hDC, Returns the bitmap and Deletes the hDC
Private Sub GetBitmap(hDC As Long, hBitmap As Long)
    
    On Error Resume Next
    
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
    
End Sub

' Creates a Picture Object from a handle to a bitmap
Private Function CreatePicture(hBitmap As Long) As IPicture
    
    On Error Resume Next
    
    Dim IID_IDispatch As GUID
    Dim Pic           As PICTDESC
    Dim IPic          As IPicture
    
    ' Fill in OLE IDispatch Interface ID
    IID_IDispatch.Data1 = &H20400
    IID_IDispatch.Data4(0) = &HC0
    IID_IDispatch.Data4(7) = &H46
        
    ' Fill Pic with necessary parts
    Pic.Size = Len(Pic)        ' Length of structure
    Pic.Type = PICTYPE_BITMAP  ' Type of Picture (bitmap)
    Pic.hBmp = hBitmap         ' Handle to bitmap

    ' Create the picture
    OleCreatePictureIndirect Pic, IID_IDispatch, True, IPic
    Set CreatePicture = IPic
    
End Function

' Returns a resized version of the picture
Public Function Resize(Handle As Long, PicType As PictureTypeConstants, Width As Long, Height As Long, Optional BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    
    On Error Resume Next
    
    Dim Img       As Long
    Dim hDC       As Long
    Dim hBitmap   As Long
    Dim WmfHeader As wmfPlaceableFileHeader
    
    ' Determine pictyre type
    Select Case PicType
    Case vbPicTypeBitmap
         GdipCreateBitmapFromHBITMAP Handle, ByVal 0&, Img
    Case vbPicTypeMetafile
         FillInWmfHeader WmfHeader, Width, Height
         GdipCreateMetafileFromWmf Handle, False, WmfHeader, Img
    Case vbPicTypeEMetafile
         GdipCreateMetafileFromEmf Handle, False, Img
    Case vbPicTypeIcon
         ' Does not return a valid Image object
         GdipCreateBitmapFromHICON Handle, Img
    End Select
    
    ' Continue with resizing only if we have a valid image object
    If Img Then
        InitDC hDC, hBitmap, BackColor, Width, Height
        gdipResize Img, hDC, Width, Height, RetainRatio
        GdipDisposeImage Img
        GetBitmap hDC, hBitmap
        Set Resize = CreatePicture(hBitmap)
    End If
    
End Function

' Fills in the wmfPlacable header
Private Sub FillInWmfHeader(WmfHeader As wmfPlaceableFileHeader, Width As Long, Height As Long)
    
    On Error Resume Next
    
    WmfHeader.BoundingBox.Right = Width
    WmfHeader.BoundingBox.Bottom = Height
    WmfHeader.Inch = 1440
    WmfHeader.Key = GDIP_WMF_PLACEABLEKEY
    
End Sub

