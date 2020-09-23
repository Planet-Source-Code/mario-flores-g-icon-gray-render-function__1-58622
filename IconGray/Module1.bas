Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32.dll" (ByRef piconinfo As ICONINFO) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


' Type - GetObjectAPI.lpObject
Private Type BITMAP
  bmType       As Long    'LONG   // Specifies the bitmap type. This member must be zero.
  bmWidth      As Long    'LONG   // Specifies the width, in pixels, of the bitmap. The width must be greater than zero.
  bmHeight     As Long    'LONG   // Specifies the height, in pixels, of the bitmap. The height must be greater than zero.
  bmWidthBytes As Long    'LONG   // Specifies the number of bytes in each scan line. This value must be divisible by 2, because Windows assumes that the bit values of a bitmap form an array that is word aligned.
  bmPlanes     As Integer 'WORD   // Specifies the count of color planes.
  bmBitsPixel  As Integer 'WORD   // Specifies the number of bits required to indicate the color of a pixel.
  bmBits       As Long    'LPVOID // Points to the location of the bit values for the bitmap. The bmBits member must be a long pointer to an array of character (1-byte) values.
End Type

' Type - CreateIconIndirect / GetIconInfo
Private Type ICONINFO
  fIcon    As Long 'BOOL    // Specifies whether this structure defines an icon or a cursor. A value of TRUE specifies an icon; FALSE specifies a cursor.
  xHotspot As Long 'DWORD   // Specifies the x-coordinate of a cursor’s hot spot. If this structure defines an icon, the hot spot is always in the center of the icon, and this member is ignored.
  yHotspot As Long 'DWORD   // Specifies the y-coordinate of the cursor’s hot spot. If this structure defines an icon, the hot spot is always in the center of the icon, and this member is ignored.
  hbmMask  As Long 'HBITMAP // Specifies the icon bitmask bitmap. If this structure defines a black and white icon, this bitmask is formatted so that the upper half is the icon AND bitmask and the lower half is the icon XOR bitmask. Under this condition, the height should be an even multiple of two. If this structure defines a color icon, this mask only defines the AND bitmask of the icon.
  hbmColor As Long 'HBITMAP // Identifies the icon color bitmap. This member can be optional if this structure defines a black and white icon. The AND bitmask of hbmMask is applied with the SRCAND flag to the destination; subsequently, the color bitmap is applied (using XOR) to the destination by using the SRCINVERT flag.
End Type



Public Function RenderIconGrayscale(ByVal Dest_hDC As Long, _
                                    ByVal hIcon As Long, _
                                    Optional ByVal Dest_X As Long, _
                                    Optional ByVal Dest_Y As Long, _
                                    Optional ByVal Dest_Height As Long, _
                                    Optional ByVal Dest_Width As Long) As Boolean
  
  Dim hBMP_Mask  As Long
  Dim hBMP_Image As Long
  Dim hBMP_Prev  As Long
  Dim hIcon_Temp As Long
  Dim hDC_Temp   As Long
  
  ' Make sure parameters passed are valid
  If Dest_hDC = 0 Or hIcon = 0 Then Exit Function
  
  ' Extract the bitmaps from the icon
  If GetIconBitmaps(hIcon, hBMP_Mask, hBMP_Image) = False Then Exit Function
  
  ' Create a memory DC to work with
  hDC_Temp = CreateCompatibleDC(0)
  

  If hDC_Temp = 0 Then GoTo CleanUp
  
  ' Make the image bitmap gradient
  If RenderBitmapGrayscale(hDC_Temp, hBMP_Image, 0, 0) = False Then GoTo CleanUp
  
  ' Extract the gradient bitmap out of the DC
  SelectObject hDC_Temp, hBMP_Prev

  
  ' Take the newly gradient bitmap and make a gradient icon from it
  hIcon_Temp = CreateIconFromBMP(hBMP_Mask, hBMP_Image)
  If hIcon_Temp = 0 Then GoTo CleanUp
  
  ' Draw the newly created gradient icon onto the specified DC
  If DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hIcon_Temp, Dest_Width, Dest_Height, 0, 0, &H3) <> 0 Then
    RenderIconGrayscale = True
  End If
  
CleanUp:
  
  DestroyIcon hIcon_Temp: hIcon_Temp = 0
  DeleteDC hDC_Temp: hDC_Temp = 0
  DeleteObject hBMP_Mask: hBMP_Mask = 0
  DeleteObject hBMP_Image: hBMP_Image = 0
  
End Function


Public Function GetIconBitmaps(ByVal hIcon As Long, _
                               ByRef Return_hBmpMask As Long, _
                               ByRef Return_hBmpImage As Long) As Boolean
  
  Dim TempICONINFO As ICONINFO
  
  If GetIconInfo(hIcon, TempICONINFO) = 0 Then Exit Function
  Return_hBmpMask = TempICONINFO.hbmMask
  Return_hBmpImage = TempICONINFO.hbmColor
  GetIconBitmaps = True
  
End Function

'=============================================================================================================
Public Function RenderBitmapGrayscale(ByVal Dest_hDC As Long, _
                                      ByVal hBitmap As Long, _
                                      Optional ByVal Dest_X As Long, _
                                      Optional ByVal Dest_Y As Long, _
                                      Optional ByVal Srce_X As Long, _
                                      Optional ByVal Srce_Y As Long _
                                      ) As Boolean
  
  Dim TempBITMAP  As BITMAP
  Dim hScreen     As Long
  Dim hDC_Temp    As Long
  Dim hBMP_Prev   As Long
  Dim MyCounterX  As Long
  Dim MyCounterY  As Long
  Dim NewColor    As Long
  Dim hNewPicture As Long
  Dim DeletePic   As Boolean
  
  ' Make sure parameters passed are valid
  If Dest_hDC = 0 Or hBitmap = 0 Then Exit Function
  
  ' Get the handle to the screen DC
  hScreen = GetDC(0)
  If hScreen = 0 Then Exit Function
  
  ' Create a memory DC to work with the picture
  hDC_Temp = CreateCompatibleDC(hScreen)
  If hDC_Temp = 0 Then GoTo CleanUp
  
  ' If the user specifies NOT to alter the original, then make a copy of it to use
    DeletePic = False
    hNewPicture = hBitmap
    
  ' Select the bitmap into the DC
  hBMP_Prev = SelectObject(hDC_Temp, hNewPicture)
  
  ' Get the height / width of the bitmap in pixels
  If GetObjectAPI(hNewPicture, Len(TempBITMAP), TempBITMAP) = 0 Then GoTo CleanUp
  If TempBITMAP.bmHeight <= 0 Or TempBITMAP.bmWidth <= 0 Then GoTo CleanUp
  
  ' Loop through each pixel and conver it to it's grayscale equivelant
  For MyCounterX = 0 To TempBITMAP.bmWidth - 1
    For MyCounterY = 0 To TempBITMAP.bmHeight - 1
      NewColor = GetPixel(hDC_Temp, MyCounterX, MyCounterY)
      If NewColor <> -1 Then
        Select Case NewColor
          ' If the color is already a grey shade, no need to convert it
          Case vbBlack, vbWhite, &H101010, &H202020, &H303030, &H404040, &H505050, &H606060, &H707070, &H808080, &HA0A0A0, &HB0B0B0, &HC0C0C0, &HD0D0D0, &HE0E0E0, &HF0F0F0
            NewColor = NewColor

          Case Else
            NewColor = 0.33 * (NewColor Mod 256) + _
                     0.59 * ((NewColor \ 256) Mod 256) + _
                     0.11 * ((NewColor \ 65536) Mod 256)
            NewColor = RGB(NewColor, NewColor, NewColor)

        End Select
        SetPixel hDC_Temp, MyCounterX, MyCounterY, NewColor
      End If
    Next MyCounterY
  Next MyCounterX
  
  ' Display the picture on the specified hDC
  BitBlt Dest_hDC, Dest_X, Dest_Y, TempBITMAP.bmWidth, TempBITMAP.bmHeight, hDC_Temp, Srce_X, Srce_Y, vbSrcCopy
  
  RenderBitmapGrayscale = True
  
CleanUp:
  
  ReleaseDC 0, hScreen: hScreen = 0
  SelectObject hDC_Temp, hBMP_Prev
  DeleteDC hDC_Temp: hDC_Temp = 0
  If DeletePic = True Then
    DeleteObject hNewPicture
    hNewPicture = 0
  End If
  
End Function

Public Function CreateIconFromBMP(ByVal hBMP_Mask As Long, _
                                  ByVal hBMP_Image As Long) As Long
  
  Dim TempICONINFO As ICONINFO
  
  If hBMP_Mask = 0 Or hBMP_Image = 0 Then Exit Function
  
  TempICONINFO.fIcon = 1
  TempICONINFO.hbmMask = hBMP_Mask
  TempICONINFO.hbmColor = hBMP_Image
  
  CreateIconFromBMP = CreateIconIndirect(TempICONINFO)
  
End Function

