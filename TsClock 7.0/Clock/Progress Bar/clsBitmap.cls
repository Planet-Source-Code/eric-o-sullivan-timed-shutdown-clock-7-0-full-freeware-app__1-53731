VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Creates and uses a bitmap effiently and also provides access to some graphical api functions and procedures and some higher-level graphical sub-routines"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=================================================
'AUTHOR : Eric O'Sullivan
' -----------------------------------------------
'DATE : 11/Nov/2001
' -----------------------------------------------
'CONTACT: DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE : Bitmap Class
' -----------------------------------------------
'COMMENTS :
'This module was made for using api graphics
'functions in your programs. With the following
'api calls and function and procedures written
'by me, you have to tools to do almost anything.
'The only api function listed here that is not
'directly used by any piece of code in this module
'is BitBlt. You have the tools to create and
'manipulate graphics, but it is still necessary
'to display them manually. The functions
'themselves mostly need hDc or a handle to work.
'You can find this hDc in both a forms and
'pictureboxs' properties. I have also set up a data
'type called BitmapStruc. For my programs, I have
'used this structure almost exclusivly for the
'graphics. The structure holds all the information
'needed to reference a bitmap created using this
'module (CreateNewBitmap, DeleteBitmap). Please
'keep in mind that any object (bitmap, brush, pen
'etc) needs to be deleted after use or else it
'will stay in memory until the program is finished.
'Not doing so will eventually cause your program
'to take up ALL your computers recources. Also for
'anyone using optional paramters, it is probably
'better to use a default parameter values to
'determine whether or not a parameter has been
'passed than the function "IsMissing()".
'
'---------------------
' 19/July/2002
' The comments above were for the module
' APIGraphics. This class is a direct conversion
' from that module and is not completely tested as
' yet. For the purposes of this program (whatever
' this class happens to currently be in), the
' class will work correctly and is tested for the
' program purpose only. Currently known, there is
' are two dependancies, msimg32.dll and stdole2.tlb
'---------------------
'
'Thank you,
'Eric
'=================================================

Option Explicit

'--------------------------------------------------------------------------
'API calls
'--------------------------------------------------------------------------
'These functions are sorted alphabetically.

'this will alphablend two bitmaps by a specified
'blend amount.
Private Declare Function AlphaBlend _
       Lib "msimg32" _
            (ByVal hDcDest As Long, _
             ByVal intLeftDest As Integer, _
             ByVal intTopDest As Integer, _
             ByVal intWidthDest As Integer, _
             ByVal intHeightDest As Integer, _
             ByVal hDcSource As Long, _
             ByVal intLeftSource As Integer, _
             ByVal intTopSource As Integer, _
             ByVal intWidthSource As Integer, _
             ByVal intHeightSource As Integer, _
             ByVal lngBlendFunctionStruc As Long) _
             As Long

'This is used to copy bitmaps
Private Declare Function BitBlt _
       Lib "gdi32" _
            (ByVal hDestDC As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal nWidth As Long, _
             ByVal nHeight As Long, _
             ByVal hSrcDC As Long, _
             ByVal xSrc As Long, _
             ByVal ySrc As Long, _
             ByVal dwRop As Long) _
             As Long

'used to change various windows settings
Private Declare Function ChangeDisplaySettings _
       Lib "user32" _
       Alias "ChangeDisplaySettingsA" _
            (ByRef wef As Any, _
             ByVal i As Long) _
             As Long

'creates a brush object which can be applied to a bitmap
Private Declare Function CreateBrushIndirect _
       Lib "gdi32" _
            (lpLogBrush As LogBrush) _
             As Long

'creates a colourspace object which can be applied to a bitmap
Private Declare Function CreateColorSpace _
       Lib "gdi32" _
       Alias "CreateColorSpaceA" _
            (lplogcolorspace As LogColorSpace) _
             As Long

'the will create a bitmap compatable with the passed hDc
Private Declare Function CreateCompatibleBitmap _
       Lib "gdi32" _
            (ByVal hDc As Long, _
            ByVal nWidth As Long, _
            ByVal nHeight As Long) _
            As Long

'this create a compatable device context with the specified
'windows handle
Private Declare Function CreateCompatibleDC _
       Lib "gdi32" _
            (ByVal hDc As Long) _
             As Long

'creates an elliptical region in a hDc
Private Declare Function CreateEllipticRgn _
       Lib "gdi32" _
            (ByVal X1 As Long, _
             ByVal Y1 As Long, _
             ByVal X2 As Long, _
             ByVal Y2 As Long) _
             As Long

'creates an elliptical region in a hDc
Private Declare Function CreateEllipticRgnIndirect _
       Lib "gdi32" _
            (EllRect As Rect) _
             As Long

'creates a font compatable with the specified device context
Private Declare Function CreateFontIndirect _
       Lib "gdi32" _
       Alias "CreateFontIndirectA" _
            (lpLogFont As LogFont) _
             As Long

'creates a pen that can be applied to a hDc
Private Declare Function CreatePen _
       Lib "gdi32" _
            (ByVal nPenStyle As Long, _
             ByVal nWidth As Long, _
             ByVal crColor As Long) _
             As Long

'creates a pen that can be applied to a hDc
Private Declare Function CreatePenIndirect _
       Lib "gdi32" _
            (lpLogPen As LogPen) _
             As Long

'creates a rectangular region on a hDc
Private Declare Function CreateRectRgn _
       Lib "gdi32" _
            (Left As Integer, _
             Top As Integer, _
             Right As Integer, _
             Bottom As Integer) _
             As Long

'creates a solid colour brush to be applied to
'a bitmap
Private Declare Function CreateSolidBrush _
       Lib "gdi32" _
            (ColorRef As Long) _
             As Long

'removes a device context from memory
Private Declare Function DeleteDC _
       Lib "gdi32" _
            (ByVal hDc As Long) _
             As Long

'removes an object such as a brush or bitmap from memory
Private Declare Function DeleteObject _
       Lib "gdi32" _
            (ByVal hObject As Long) _
             As Long

'this draws the animated minimize/maximize rectangeles
Private Declare Function DrawAnimatedRects _
       Lib "user32" _
            (ByVal hwnd As Long, _
             ByVal idAni As Long, _
             lprcFrom As Rect, _
             lprcTo As Rect) _
             As Long

'this draws an icon onto a surphase (eg, a bitmap)
Private Declare Function DrawIconEx _
       Lib "user32" _
            (ByVal hDc As Long, _
             ByVal xLeft As Long, _
             ByVal yTop As Long, _
             ByVal hIcon As Long, _
             ByVal cxWidth As Long, _
             ByVal cyWidth As Long, _
             ByVal istepIfAniCur As Long, _
             ByVal hbrFlickerFreeDraw As Long, _
             ByVal diFlags As Long) _
             As Long

'this draws text onto the bitmap
Private Declare Function DrawText _
       Lib "user32" _
       Alias "DrawTextA" _
            (ByVal hDc As Long, _
             ByVal lpStr As String, _
             ByVal nCount As Long, _
             lpRect As Rect, _
             ByVal wFormat As Long) _
             As Long

'this draws an ellipse onto the bitmap
Private Declare Function Ellipse _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             X1 As Integer, _
             Y1 As Integer, _
             X2 As Integer, _
             Y2 As Integer) _
             As Boolean

'this will set the display settings
Private Declare Function EnumDisplaySettings _
       Lib "user32" _
       Alias "EnumDisplaySettingsA" _
            (ByVal A As Long, _
             ByVal B As Long, _
             wef As DEVMODE) _
             As Boolean

'this provides more control than the CreatePen function
Private Declare Function ExtCreatePen _
       Lib "gdi32" _
            (ByVal dwPenStyle As Long, _
             ByVal dwWidth As Long, _
             lplb As LogBrush, _
             ByVal dwStyleCount As Long, _
             lpStyle As Long) _
             As Long

'this will fill the rectangle with the brush applied
'to the hDc
Private Declare Function FillRect _
       Lib "user32" _
            (ByVal hwnd As Long, _
             Fill As Rect, _
             hBrush As Long) _
             As Integer

'this will fill a region with the brush specified
Private Declare Function FillRgn _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal HRgn As Long, _
             hBrush As Long) _
             As Boolean

'this will find a window based on its class and
'window name
Private Declare Function FindWindow _
       Lib "user32" _
       Alias "FindWindowA" _
            (ByVal lpClassName As String, _
             ByVal lpWindowName As String) _
             As Long

'this will return the handle of the top-most window
Private Declare Function GetActiveWindow _
       Lib "user32" _
            () _
             As Long

'this will get the state of any specified key - even
'the mouse buttons
Private Declare Function GetAsyncKeyState _
       Lib "user32" _
            (ByVal vKey As Long) _
             As Integer

'this will get the dimensions of the specified bitmap
Private Declare Function GetBitmapDimensionEx _
       Lib "gdi32" _
            (ByVal hBitmap As Long, _
             lpDimension As SizeType) _
             As Long

'this will get the class name from the handle of the
'window specified
Private Declare Function GetClassName _
       Lib "user32" _
       Alias "GetClassNameA" _
            (ByVal hwnd As Long, _
             ByVal lpClassName As String, _
             ByVal nMaxCount As Long) _
             As Long

'this will capture a screen shot of the specified
'area of the client
Private Declare Function GetClientRect _
       Lib "user32" _
            (ByVal hwnd As Long, _
             lpRect As Rect) _
             As Long

'this gets the cursors icon picture
Private Declare Function GetCursor _
       Lib "user32" _
            () _
             As Long

'this gets the position of the cursor on the screen
'(given in pixels)
Private Declare Function GetCursorPos _
       Lib "user32" _
            (lpPoint As PointAPI) _
             As Long

'gets a hDc of the specified window
Private Declare Function GetDC _
       Lib "user32" _
            (ByVal hwnd As Long) _
            As Long

'gets the entire screen area
Private Declare Function GetDesktopWindow _
       Lib "user32" _
            () _
             As Long

'this will get the current devices' capabilities
Private Declare Function GetDeviceCaps _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal nIndex As Long) _
             As Long

'get the last error to occur from within the api
Private Declare Function GetLastError _
       Lib "kernel32" _
            () _
             As Long

'get the handle of the menu bar on a window
Private Declare Function GetMenu _
        Lib "user32" _
            (ByVal hwnd As Long) _
             As Long

'Get the sub menu ID number. This is used to
'reference sub menus along with their handle
Private Declare Function GetMenuItemID _
        Lib "user32" _
            (ByVal hMenu As Long, _
             ByVal nPos As Long) _
             As Long

'get information about the specified object such as
'its dimensions etc.
Private Declare Function GetObjectAPI _
       Lib "gdi32" _
       Alias "GetObjectA" _
            (ByVal hObject As Long, _
             ByVal nCount As Long, _
             lpObject As Any) _
             As Long

'get the colour of the specified pixel
Private Declare Function GetPixel _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal X As Long, _
             ByVal Y As Long) _
             As Long

'this will get the handle of a specified
'sub menu using the handle of the menu
'and the item ID
Private Declare Function GetSubMenu _
        Lib "user32" _
            (ByVal hMenu As Long, _
             ByVal nPos As Long) _
             As Long

'get the specified system colour
Private Declare Function GetSysColor _
        Lib "user32" _
            (ByVal nIndex As Long) _
             As Long

'get the dimensions of the applied text metrics for
'the device context
Private Declare Function GetTextMetrics _
       Lib "gdi32" _
       Alias "GetTextMetricsA" _
            (ByVal hDc As Long, _
             lpMetrics As TEXTMETRIC) _
             As Long

'returns the amount of time windows has been active for
'in milliseconds (sec/1000)
Private Declare Function GetTickCount _
       Lib "kernel32" _
            () _
             As Long  'very usefull timing function ;)

'returns extended information about the version of the
'operating system currently running
Private Declare Function GetVersionEx _
        Lib "kernel32" _
        Alias "GetVersionExA" _
            (ByRef lpVersionInformation As OSVERSIONINFO) _
             As Long

'retrieves the handle of a window that has the
'specified relationship to the specified window.
Private Declare Function GetWindow _
       Lib "user32" _
            (ByVal hwnd As Long, _
             ByVal wCmd As Long) _
             As Long

'gets the area the specified window takes up
Private Declare Function GetWindowRect _
       Lib "user32" _
            (ByVal hwnd As Long, _
             lpRect As Rect) _
             As Long

'increases the size of a rect structure
Private Declare Function InflateRect _
       Lib "user32" _
            (lpRect As Rect, _
             ByVal X As Long, _
             ByVal Y As Long) _
             As Long

'gets any intersection of two rectangles
Private Declare Function IntersectRect _
       Lib "user32" _
            (lpDestRect As Rect, _
             lpSrc1Rect As Rect, _
             lpSrc2Rect As Rect) _
             As Long

'draws a line from the current position to the
'specified co-ordinates
Private Declare Function LineTo _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             XEnd As Integer, _
             YEnd As Integer) _
             As Boolean

'this will load a bitmap directly from a file (NT compatable)
Private Declare Function LoadNTBitmap _
       Lib "user32" _
       Alias "LoadBitmapA" _
            (ByVal hInstance As Long, _
             ByVal lpBitmapName As Any) _
             As Long

'this will load a cursor into a device context
Private Declare Function LoadCursor _
       Lib "user32" _
       Alias "LoadCursorA" _
            (ByVal hInstance As Long, _
             ByVal lpCursorName As Any) _
             As Long

'this will load an image into a device context (NOT NT compatable)
Private Declare Function LoadImage _
       Lib "user32" _
       Alias "LoadImageA" _
            (ByVal hInst As Long, _
             ByVal lpsz As String, _
             ByVal un1 As Long, _
             ByVal n1 As Long, _
             ByVal n2 As Long, _
             ByVal un2 As Long) _
             As Long

'This stops the specified window from updating
'its display. This is mainly used to help cut out
'flicker but does not work on controls.
Private Declare Function LockWindowUpdate _
       Lib "user32" _
            (ByVal hwndLock As Long) _
             As Long
'changes some of a menu's properties
Private Declare Function ModifyMenu _
        Lib "user32" _
        Alias "ModifyMenuA" _
            (ByVal hMenu As Long, _
             ByVal nPosition As Long, _
             ByVal wFlags As Long, _
             ByVal wIDNewItem As Long, _
             ByVal lpString As Any) _
             As Long

'moves the current position to the specified
'point
Private Declare Function MoveToEx _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             X As Integer, _
             Y As Integer, _
             lpPoint As PointAPI) _
             As Boolean

'multiplies two numbers and divides them by a third
'numbers
Private Declare Function MulDiv _
       Lib "kernel32" _
            (ByVal nNumber As Long, _
             ByVal nNumerator As Long, _
             ByVal nDenominator As Long) _
             As Long

'This will increase or decrease a rectangles
'co-ordinates by the specified amount. Usefull
'for moving graphic blocks as rect structures.
Private Declare Function OffsetRect _
       Lib "user32" _
            (lpRect As Rect, _
             ByVal X As Long, _
             ByVal Y As Long) _
             As Long

'Pattern Blitter. Used to draw a pattern onto
'a device context
Private Declare Function PatBlt _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal nWidth As Long, _
             ByVal nHeight As Long, _
             ByVal dwRop As Long) _
             As Long

'This draws a polygon onto a device context
'useing an array.
Private Declare Function Polygon _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             lpPoint As PointAPI, _
             ByVal nCount As Long) _
             As Long

'This will draw a set of lines to the specifed
'points
Private Declare Function Polyline _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             lpPoint As PointAPI, _
             ByVal nCount As Long) _
             As Long

'This will draw a set of lines starting from
'the current position on the device context.
Private Declare Function PolylineTo _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             lppt As PointAPI, _
             ByVal cCount As Long) _
             As Long

'This draws a rectangle onto the device
'context
Private Declare Function Rectangle _
       Lib "gdi32" _
            (ByVal hwnd As Long, _
             X1 As Integer, _
             Y1 As Integer, _
             X2 As Integer, _
             Y2 As Integer) _
             As Long

'This will release a device context from
'memory. Not to be confused with DeleteDC
Private Declare Function ReleaseDC _
       Lib "user32" _
            (ByVal hwnd As Long, _
             ByVal hDc As Long) _
             As Long

'this will draw a round-cornered rectangle
'onto the specified device context
Private Declare Function RoundRect _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal Left As Long, _
             ByVal Top As Long, _
             ByVal Right As Long, _
             ByVal Bottom As Long, _
             ByVal X3 As Long, _
             ByVal Y3 As Long) _
             As Long

'this can convert entire type structures
'to other types like a Long
Private Declare Sub RtlMoveMemory _
        Lib "kernel32.dll" _
            (Destination As Any, _
             Source As Any, _
             ByVal Length As Long)

'this will select the specified object to
'a window or device context
Private Declare Function SelectObject _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal hObject As Long) _
             As Long

'This sets the background colour on a bitmap
Private Declare Function SetBkColor _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal crColor As Long) _
             As Long

'This sets the background mode on a bitmap
'(eg, transparent, solid etc)
Private Declare Function SetBkMode _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal nBkMode As Long) _
             As Long

'The color adjustment values are used to adjust
'the input color of the source bitmap for calls
'to the StretchBlt and StretchDIBits functions
'when HALFTONE mode is set.
Private Declare Function SetColorAdjustment _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             lpca As COLORADJUSTMENT) _
             As Long

'sets the colourspace to a device context
Private Declare Function SetColorSpace _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal hcolorspace As Long) _
             As Long

'sets the position of the cursor on the screen
Private Declare Function SetCursorPos _
        Lib "user32" _
            (ByVal X As Long, _
             ByVal Y As Long) _
             As Long

'sets the bitmap of a menu
Private Declare Function SetMenuItemBitmaps _
        Lib "user32" _
            (ByVal hMenu As Long, _
             ByVal nPosition As Long, _
             ByVal wFlags As Long, _
             ByVal hBitmapUnchecked As Long, _
             ByVal hBitmapChecked As Long) _
             As Long

'sets the current information about the selected menu
Private Declare Function SetMenuItemInfo _
        Lib "user32" _
        Alias "SetMenuItemInfoA" _
            (ByVal hMenu As Long, _
             ByVal uItem As Long, _
             ByVal fByPosition As Long, _
             lpmii As MENUITEMINFO) _
             As Long

'sets the colour of the specified pixel
Private Declare Function SetPixel _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal crColor As Long) _
             As Long

'sets the rectangles size and position
Private Declare Function SetRect _
       Lib "user32" _
            (lpRect As Rect, _
             ByVal X1 As Long, _
             ByVal Y1 As Long, _
             ByVal X2 As Long, _
             ByVal Y2 As Long) _
             As Long

'set a system colour
Private Declare Function SetSysColors _
        Lib "user32" _
            (ByVal nChanges As Long, _
             lpSysColor As Long, _
             lpColorValues As Long) _
             As Long

'sets the current text colour
Private Declare Function SetTextColor _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal crColor As Long) _
             As Long

'pauses the execution of the programs thread
'for a specified amount of milliseconds
Private Declare Sub Sleep _
       Lib "kernel32" _
            (ByVal dwMilliseconds As Long)

'used to stretch or shrink a bitmap
Private Declare Function StretchBlt _
       Lib "gdi32" _
            (ByVal hDc As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal nWidth As Long, _
             ByVal nHeight As Long, _
             ByVal hSrcDC As Long, _
             ByVal xSrc As Long, _
             ByVal ySrc As Long, _
             ByVal nSrcWidth As Long, _
             ByVal nSrcHeight As Long, _
             ByVal dwRop As Long) _
             As Long

'--------------------------------------------------------------------------
'enumerator section
'--------------------------------------------------------------------------

'the direction of the gradient
Public Enum GradientTo
    GradHorizontal = 0
    GradVertical = 1
End Enum

'in twips or pixels
Public Enum Scaling
    InTwips = 0
    InPixels = 1
End Enum

'The key values of the mouse buttons
Public Enum MouseKeys
    MouseLeft = 1
    MouseRight = 2
    MouseMiddle = 4
End Enum

'text alignment constants
Public Enum AlignText
    vbLeftAlign = 1
    vbCentreAlign = 2
    vbRightAlign = 3
End Enum

'bitmap flip constants
Public Enum FlipType
    FlipHorizontally = 0
    FlipVertically = 1
End Enum

'image load constants
Public Enum LoadType
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
    'IMAGE_ENHMETAFILE  = 3
End Enum
    
'rotate bitmap constants
Public Enum RotateType
    RotateRight = 0
    RotateLeft = 1
    Rotate180 = 2
End Enum

Public Enum PaintMode
    'BitBlt constants
    P_AND = &H8800C6  ' (DWORD) dest = source AND dest
    P_COPY = &HCC0020 ' (DWORD) dest = source
    P_ERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
    P_INVERT = &H660046       ' (DWORD) dest = source XOR dest
    P_PAINT = &HEE0086        ' (DWORD) dest = source OR dest
    P_MERGE_COPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
    P_MERGE_PAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
    P_NOT_COPY = &H330008      ' (DWORD) dest = (NOT source)
    P_NOT_ERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
End Enum

'system colour constants
Public Enum SystemColours
     col_Scrollbar = 0              'The Scrollbar colour
     col_Background = 1             'Colour of the background with no wallpaper
     col_ActiveCaption = 2          'Caption of Active Window
     col_InactiveCaption = 3        'Caption of Inactive window
     col_MenuBar = 4                'Menu
     col_Window = 5                 'Windows background
     col_WindowFrame = 6            'Window frame
     col_MenuText = 7               'Window Text
     col_WindowText = 8             '3D dark shadow (Win95)
     col_CaptionText = 9            'Text in window caption
     col_ActiveBorder = 10          'Border of active window
     col_InactiveBorder = 11        'Border of inactive window
     col_AppWorkSpace = 12          'Background of MDI desktop
     col_Highlight = 13             'Selected item background
     col_HighlightText = 14         'Selected menu item
     col_BtnFace = 15               'Button
     col_BtnSadow = 16              '3D shading of button
     col_GrayText = 17              'Grey text, of zero if dithering is used.
     col_BtnText = 18               'Button text
     col_InactiveCaptionText = 19   'Text of inactive window
     col_BtnHighlight = 20          '3D highlight of button
     col_ActiveCaption2 = 27        'Win98 only: 2nd active window color
     col_InactiveCaption2 = 28      'Win98 only: 2nd inactive window color
End Enum

'--------------------------------------------------------------------------
'Programmer defined data types
'--------------------------------------------------------------------------

'AlphaBlend information for bitmaps
Private Type BLENDFUNCTION
    bytBlendOp              As Byte 'currently the only blend op supported by windows 98+ is AC_SRC_OVER
    bytBlendFlags           As Byte 'must be left blank
    bytSourceConstantAlpha  As Byte 'the amount to blend by. Must be between 0 and 255
    bytAlphaFormat          As Byte 'don't set this. If you wish more infor, go to "http://msdn.microsoft.com/library/default.asp?url=/library/en-us/gdi/bitmaps_3b3m.asp"
End Type

'Bitmap structue for menu information
Private Type MENUITEMINFO
    cbSize          As Long
    fMask           As Long
    fType           As Long
    fState          As Long
    wID             As Long
    hSubMenu        As Long
    hbmpChecked     As Long
    hbmpUnchecked   As Long
    dwItemData      As Long
    dwTypeData      As String
    cch             As Long
End Type

'size structure
Private Type SizeType
    cx As Long
    cy As Long
End Type

'Text metrics
Private Type TEXTMETRIC
    tmHeight            As Long
    tmAscent            As Long
    tmDescent           As Long
    tmInternalLeading   As Long
    tmExternalLeading   As Long
    tmAveCharWidth      As Long
    tmMaxCharWidth      As Long
    tmWeight            As Long
    tmOverhang          As Long
    tmDigitizedAspectX  As Long
    tmDigitizedAspectY  As Long
    tmFirstChar         As Byte
    tmLastChar          As Byte
    tmDefaultChar       As Byte
    tmBreakChar         As Byte
    tmItalic            As Byte
    tmUnderlined        As Byte
    tmStruckOut         As Byte
    tmPitchAndFamily    As Byte
    tmCharSet           As Byte
End Type


Private Type COLORADJUSTMENT
    caSize              As Integer
    caFlags             As Integer
    caIlluminantIndex   As Integer
    caRedGamma          As Integer
    caGreenGamma        As Integer
    caBlueGamma         As Integer
    caReferenceBlack    As Integer
    caReferenceWhite    As Integer
    caContrast          As Integer
    caBrightness        As Integer
    caColorfulness      As Integer
    caRedGreenTint      As Integer
End Type

Private Type CIEXYZ
    ciexyzX As Long
    ciexyzY As Long
    ciexyzZ As Long
End Type

Private Type CIEXYZTRIPLE
    ciexyzRed   As CIEXYZ
    ciexyzGreen As CIEXYZ
    ciexyBlue   As CIEXYZ
End Type

Private Type LogColorSpace
    lcsSignature    As Long
    lcsVersion      As Long
    lcsSize         As Long
    lcsCSType       As Long
    lcsIntent       As Long
    lcsEndPoints    As CIEXYZTRIPLE
    lcsGammaRed     As Long
    lcsGammaGreen   As Long
    lcsGammaBlue    As Long
    lcsFileName     As String * 26 'MAX_PATH
End Type

'display settings (800x600 etc)
Private Type DEVMODE
    dmDeviceName        As String * 32
    dmSpecVersion       As Integer
    dmDriverVersion     As Integer
    dmSize              As Integer
    dmDriverExtra       As Integer
    dmFields            As Long
    dmOrientation       As Integer
    dmPaperSize         As Integer
    dmPaperLength       As Integer
    dmPaperWidth        As Integer
    dmScale             As Integer
    dmCopies            As Integer
    dmDefaultSource     As Integer
    dmPrintQuality      As Integer
    dmColor             As Integer
    dmDuplex            As Integer
    dmYResolution       As Integer
    dmTTOption          As Integer
    dmCollate           As Integer
    dmFormName          As String * 32
    dmUnusedPadding     As Integer
    dmBitsPerPel        As Long
    dmPelsWidth         As Long
    dmPelsHeight        As Long
    dmDisplayFlags      As Long
    dmDisplayFrequency  As Long
End Type

Private Type Rect
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type BitmapStruc
    lngDC       As Long
    lngBitmap   As Long
    lngPointer  As Long
    Area        As Rect
End Type

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type LogPen
    lopnStyle As Long
    lopnWidth As PointAPI
    lopnColor As Long
End Type

Private Type LogBrush
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Type FontStruc
    Name        As String
    Alignment   As AlignText
    Bold        As Boolean
    Italic      As Boolean
    Underline   As Boolean
    StrikeThru  As Boolean
    PointSize   As Byte
    Colour      As Long
End Type

Private Type LogFont
    'for the DrawText api call
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName(1 To 32) As Byte
End Type

Private Type Point
    'you'll need this to reference a point on the
    'screen'
    X As Integer
    Y As Integer
End Type

'To hold the RGB value
Private Type RGBVal
    Red     As Single
    Green   As Single
    Blue    As Single
End Type

'bitmap structure for the GetObject api call
Private Type BITMAP '24 bytes
        bmType          As Long
        bmWidth         As Long
        bmHeight        As Long
        bmWidthBytes    As Long
        bmPlanes        As Integer
        bmBitsPixel     As Integer
        bmBits          As Long
End Type

'holds version information about the operating system
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

'--------------------------------------------------------------------------
'Constants section
'--------------------------------------------------------------------------

'general constants
Private Const GW_CHILD              As Integer = 5
Private Const GW_HWNDFIRST          As Integer = 0
Private Const GW_HWNDLAST           As Integer = 1
Private Const GW_HWNDNEXT           As Integer = 2
Private Const GW_HWNDPREV           As Integer = 3
Private Const GWL_WNDPROC           As Integer = (-4)
Private Const IDANI_OPEN            As Long = &H1
Private Const IDANI_CLOSE           As Long = &H2
Private Const IDANI_CAPTION         As Long = &H3
Private Const WM_USER               As Long = &H400

'alphablend constants
Private Const AC_SRC_OVER           As Long = &H0
Private Const AC_SRC_ALPHA          As Long = &H0

'Image load constants
Private Const LR_LOADFROMFILE       As Long = &H10
Private Const LR_CREATEDIBSECTION   As Long = &H2000
Private Const LR_DEFAULTSIZE        As Long = &H40

'PatBlt constants
Private Const PATCOPY               As Long = &HF00021  ' (DWORD) dest = pattern
Private Const PATINVERT             As Long = &H5A0049  ' (DWORD) dest = pattern XOR dest
Private Const PATPAINT              As Long = &HFB0A09  ' (DWORD) dest = DPSnoo
Private Const DSTINVERT             As Long = &H550009  ' (DWORD) dest = (NOT dest)
Private Const BLACKNESS             As Long = &H42      ' (DWORD) dest = BLACK
Private Const WHITENESS             As Long = &HFF0062  ' (DWORD) dest = WHITE

'Display constants
Private Const CDS_FULLSCREEN        As Integer = 4
Private Const DM_BITSPERPEL         As Long = &H40000
Private Const DM_PELSWIDTH          As Long = &H80000
Private Const DM_PELSHEIGHT         As Long = &H100000
Private Const DM_DISPLAYFLAGS       As Long = &H200000
Private Const DM_DISPLAYFREQUENCY   As Long = &H400000

'DrawText constants
Private Const DT_CENTER             As Long = &H1
Private Const DT_BOTTOM             As Long = &H8
Private Const DT_CALCRECT           As Long = &H400
Private Const DT_EXPANDTABS         As Long = &H40
Private Const DT_EXTERNALLEADING    As Long = &H200
Private Const DT_LEFT               As Long = &H0
Private Const DT_NOCLIP             As Long = &H100
Private Const DT_NOPREFIX           As Long = &H800
Private Const DT_RIGHT              As Long = &H2
Private Const DT_SINGLELINE         As Long = &H20
Private Const DT_TABSTOP            As Long = &H80
Private Const DT_TOP                As Long = &H0
Private Const DT_VCENTER            As Long = &H4
Private Const DT_WORDBREAK          As Long = &H10
Private Const TRANSPARENT           As Integer = 1
Private Const OPAQUE                As Integer = 2

'CreateBrushIndirect constants
Private Const BS_DIBPATTERN         As Integer = 5
Private Const BS_DIBPATTERN8X8      As Integer = 8
Private Const BS_DIBPATTERNPT       As Integer = 6
Private Const BS_HATCHED            As Integer = 2
Private Const BS_HOLLOW             As Integer = 1
Private Const BS_NULL               As Integer = 1
Private Const BS_PATTERN            As Integer = 3
Private Const BS_PATTERN8X8         As Integer = 7
Private Const BS_SOLID              As Integer = 0
Private Const HS_BDIAGONAL          As Integer = 3  '  /////
Private Const HS_CROSS              As Integer = 4  '  +++++
Private Const HS_DIAGCROSS          As Integer = 5  '  xxxxx
Private Const HS_FDIAGONAL          As Integer = 2  '  \\\\\
Private Const HS_HORIZONTAL         As Integer = 0  '  -----
Private Const HS_NOSHADE            As Integer = 17
Private Const HS_SOLID              As Integer = 8
Private Const HS_SOLIDBKCLR         As Integer = 23
Private Const HS_SOLIDCLR           As Integer = 19
Private Const HS_VERTICAL           As Integer = 1  '  |||||

'LogFont constants
Private Const LF_FACESIZE           As Integer = 32
Private Const FW_BOLD               As Integer = 700
Private Const FW_DONTCARE           As Integer = 0
Private Const FW_EXTRABOLD          As Integer = 800
Private Const FW_EXTRALIGHT         As Integer = 200
Private Const FW_HEAVY              As Integer = 900
Private Const FW_LIGHT              As Integer = 300
Private Const FW_MEDIUM             As Integer = 500
Private Const FW_NORMAL             As Integer = 400
Private Const FW_SEMIBOLD           As Integer = 600
Private Const FW_THIN               As Integer = 100
Private Const DEFAULT_CHARSET       As Integer = 1
Private Const OUT_CHARACTER_PRECIS  As Integer = 2
Private Const OUT_DEFAULT_PRECIS    As Integer = 0
Private Const OUT_DEVICE_PRECIS     As Integer = 5
Private Const OUT_OUTLINE_PRECIS    As Integer = 8
Private Const OUT_RASTER_PRECIS     As Integer = 6
Private Const OUT_STRING_PRECIS     As Integer = 1
Private Const OUT_STROKE_PRECIS     As Integer = 3
Private Const OUT_TT_ONLY_PRECIS    As Integer = 7
Private Const OUT_TT_PRECIS         As Integer = 4
Private Const CLIP_CHARACTER_PRECIS As Integer = 1
Private Const CLIP_DEFAULT_PRECIS   As Integer = 0
Private Const CLIP_EMBEDDED         As Integer = 128
Private Const CLIP_LH_ANGLES        As Integer = 16
Private Const CLIP_MASK             As Long = &HF
Private Const CLIP_STROKE_PRECIS    As Integer = 2
Private Const CLIP_TT_ALWAYS        As Integer = 32
Private Const WM_SETFONT            As Long = &H30
Private Const LF_FULLFACESIZE       As Integer = 64
Private Const DEFAULT_PITCH         As Integer = 0
Private Const DEFAULT_QUALITY       As Integer = 0
Private Const PROOF_QUALITY         As Integer = 2

'GetDeviceCaps constants
Private Const LOGPIXELSY            As Integer = 90       '  Logical pixels/inch in Y

'colourspace constants
Private Const MAX_PATH              As Integer = 260

'pen constants
Private Const PS_COSMETIC           As Long = &H0
Private Const PS_DASH               As Integer = 1      '  -------
Private Const PS_DASHDOT            As Integer = 3      '  _._._._
Private Const PS_DASHDOTDOT         As Integer = 4      '  _.._.._
Private Const PS_DOT                As Integer = 2      '  .......
Private Const PS_ENDCAP_ROUND       As Long = &H0
Private Const PS_ENDCAP_SQUARE      As Long = &H100
Private Const PS_ENDCAP_FLAT        As Long = &H200
Private Const PS_GEOMETRIC          As Long = &H10000
Private Const PS_INSIDEFRAME        As Long = 6
Private Const PS_JOIN_BEVEL         As Long = &H1000
Private Const PS_JOIN_MITER         As Long = &H2000
Private Const PS_JOIN_ROUND         As Long = &H0
Private Const PS_SOLID              As Long = 0

'mouse cursor constants
Private Const IDC_APPSTARTING       As Long = 32650&
Private Const IDC_ARROW             As Long = 32512&
Private Const IDC_CROSS             As Long = 32515&
Private Const IDC_IBEAM             As Long = 32513&
Private Const IDC_ICON              As Long = 32641&
Private Const IDC_NO                As Long = 32648&
Private Const IDC_SIZE              As Long = 32640&
Private Const IDC_SIZEALL           As Long = 32646&
Private Const IDC_SIZENESW          As Long = 32643&
Private Const IDC_SIZENS            As Long = 32645&
Private Const IDC_SIZENWSE          As Long = 32642&
Private Const IDC_SIZEWE            As Long = 32644&
Private Const IDC_UPARROW           As Long = 32516&
Private Const IDC_WAIT              As Long = 32514&

'menu constants
Private Const MFT_RADIOCHECK        As Long = &H200&
Private Const MF_BITMAP             As Long = &H4&
Private Const MIIM_TYPE             As Long = &H10
Private Const MIIM_SUBMENU          As Long = &H4

'some key values for GetASyncKeyState
Private Const KLeft                 As Integer = 37
Private Const KUp                   As Integer = 38
Private Const KRight                As Integer = 39
Private Const KDown                 As Integer = 40
Private Const K_KEY_DOWN            As Integer = -32767
Private Const K_PRESSED             As Integer = -32768

'extra system colour constants to be extended from
'the vb collection SystemColorConstants (Win98+)
Private Const vbActiveTitleBar2     As Double = 2147483675#
Private Const vbInactiveTitleBar2   As Double = 2147483676#

'some mathimatical constants
Private Const PI                    As Single = 3.14159265358979

'some class constants
Private Const NO_COLOUR             As Long = -2147483648# '2 ^ 32 --> 4 bytes

'--------------------------------------------------------------------------
'Variable declarations section
'--------------------------------------------------------------------------

'This stores the various points of a polygon. Use DrawPoly to draw
'the polygon with the points specified in the array.
'Also see AddToPoly, ClearPoly and DelFromPoly
Private PolygonPoints()     As PointAPI

'some general variables used by some api calls
Private mudtRectFrom        As Rect
Private mudtRectTo          As Rect
Private mlnghTray           As Long
Private mlnghStartMenu      As Long
Private mlnghChild          As Long
Private mstrClass           As String * 255
Private mlngClassNameLen    As Long
Private mlngRetVal          As Long
Private mblnResChanged      As Boolean

'=================
' class variables
'=================
Private mudtBitmap          As BitmapStruc   'the bitmap that this class encapsulates
Private mlngBackColour      As Long      'the background colour of the bitmap

'--------------------------------------------------------------------------
'Procedures/functions section
'--------------------------------------------------------------------------

Public Sub WinBlend(ByVal lngPic1hDc As Long, _
                    ByVal lngPic2hDc As Long, _
                    Optional ByVal intDestX As Integer = 0, _
                    Optional ByVal intDestY As Integer = 0, _
                    Optional ByVal intWidth As Integer = -1, _
                    Optional ByVal intHeight As Integer = -1, _
                    Optional ByVal lngPic1X As Integer = 0, _
                    Optional ByVal lngPic1Y As Integer = 0, _
                    Optional ByVal lngPic2X As Integer = 0, _
                    Optional ByVal lngPic2Y As Integer = 0, _
                    Optional ByVal sngBlendAmount As Single = 0.5, _
                    Optional ByVal enmMeasurement As Scaling = InPixels)

    'This uses the windows GDI to blend two bitmaps
    'together. If you need to use a blend mask, then
    'please use the BFAlphaBlend procedure. This function
    'is only supported by w98+ and w2000+ using the
    'Msimg32.dll library
    
    Dim udtTempBmp    As BitmapStruc    'this temperorily holds the blended pictures before copying to the destination bitmap
    Dim udtBlendInfo  As BLENDFUNCTION  'this sets the blend information for the api call
    Dim lngBlendStruc As Long           'this will hold the converted BLENDFUNCTION structure
    Dim lngResult     As Long           'this holds any error value returned from the api call
    Dim intPxlHeight  As Integer        'the height in twips of a pixel
    Dim intPxlWidth   As Integer        'the width in twips of a pixel
    
    'set the default heights and widths if necessary
    If intWidth < 0 Then
        intWidth = mudtBitmap.Area.Left + _
                   mudtBitmap.Area.Right
    End If
    If intHeight < 0 Then
        intHeight = mudtBitmap.Area.Top + _
                    mudtBitmap.Area.Bottom
    End If
    
    'convert values to pixels if necessary
    If enmMeasurement = InTwips Then
        'get the pixel intHeight and intWidth
        'values per twips in the current
        'screen resolution.
        intPxlHeight = Screen.TwipsPerPixelY
        intPxlWidth = Screen.TwipsPerPixelX
        
        'start converting the twip values to pixels
        intDestX = intDestX / intPxlWidth
        intWidth = intWidth / intPxlWidth
        lngPic1X = lngPic1X / intPxlWidth
        lngPic2X = lngPic2X / intPxlWidth
        intDestY = intDestY / intPxlHeight
        intHeight = intHeight / intPxlHeight
        lngPic1Y = lngPic1Y / intPxlHeight
        lngPic2Y = lngPic2Y / intPxlHeight
    End If
    
    'set the blend information
    With udtBlendInfo
        .bytBlendOp = AC_SRC_OVER
        .bytSourceConstantAlpha = CByte(255 * sngBlendAmount)
    End With
    
    'convert the type to a long
    Call RtlMoveMemory(lngBlendStruc, _
                       udtBlendInfo, _
                       4)
    
    With udtTempBmp
        'set the bitmap dimensions
        With .Area
            .Bottom = intHeight
            .Right = intWidth
        End With
        
        'create the new bitmap
        Call CreateNewBitmap(.lngDC, _
                             .lngBitmap, _
                             .lngPointer, _
                             .Area, _
                             mudtBitmap.lngDC)
        
        'copy the first picture
        lngResult = BitBlt(.lngDC, _
                           0, _
                           0, _
                           intWidth, _
                           intHeight, _
                           lngPic1hDc, _
                           lngPic1X, _
                           lngPic1Y, _
                           P_COPY)
        
        'blend the second picture with
        'the first picture
        lngResult = AlphaBlend(.lngDC, _
                               0, _
                               0, _
                               intWidth, _
                               intHeight, _
                               lngPic2hDc, _
                               lngPic2X, _
                               lngPic2Y, _
                               intWidth, _
                               intHeight, _
                               lngBlendStruc)
        
        'copy the picture to the destination
        lngResult = BitBlt(mudtBitmap.lngDC, _
                           intDestX, _
                           intDestY, _
                           intWidth, _
                           intHeight, _
                           .lngDC, _
                           0, _
                           0, _
                           P_COPY)
        
        'remove the temperory bitmap from memory
        Call DeleteBitmap(.lngDC, _
                          .lngBitmap, _
                          .lngPointer)
    End With
End Sub

'currently not in use for this class. See WinBlend
Private Sub BFAlphaBlend(ByVal lngDesthDc As Long, _
                        ByVal lngPic1hDc As Long, _
                        ByVal lngPic2hDc As Long, _
                        ByVal intDestX As Integer, _
                        ByVal intDestY As Integer, _
                        ByVal intWidth As Integer, _
                        ByVal intHeight As Integer, _
                        ByVal lngPic1X As Integer, _
                        ByVal lngPic1Y As Integer, _
                        ByVal lngPic2X As Integer, _
                        ByVal lngPic2Y As Integer, _
                        Optional ByVal sngBlendAmount As Single = 0.5, _
                        Optional ByVal lngMaskhDc As Long = 0, _
                        Optional ByVal intMeasurement As Scaling = InPixels)

    'This is a "brute force" alpha blend function. Because it's written in
    'vb, this function is not as fast at it might otherwise be in another
    'language like C++ or Fox.
    'The purpose of the function is to mix the colours of two bitmaps to
    'produce a result that looks like both pictures. Think of it as fading
    'one picture into another. I get the pixel colour of a point in picture1,
    'and the colour of the corresponding pixel in pixture2, find the 'middle'
    'colour and put it into the destination bitmap. There are no calls to
    'other procedures or functions other than api calls. This is to improve
    'speed as all calculations are made internally.
    'The parameter sngBlendAmount MUST be between 1 and 0. If not then
    'the value is rounded to 1 or zero.
    'However, sngBlendAmount is ignored if a Mask bitmap has been specified.
    'Please note that if the mask used only contains black or white pixels,
    'then it is recommended that you use the MergeBitmaps procedure as
    'it will process the bitmaps much faster (by about 15 to 30 times).
    'Keep in mind that using a mask bitmap is about 25% slower than
    'specifying the blend amount.

    Dim TempBmp         As BitmapStruc  'a temperory bitmap
    Dim lngResult       As Long         'any result returned from an api call
    Dim Col1            As RGBVal       'used to store a pixel colour in RGB format
    Dim Col2            As RGBVal       'used to store a pixel colour in RGB format
    Dim BlendCol        As RGBVal       'used to store a pixel colour in RGB format
    Dim MaskCol         As RGBVal       'used to store a pixel colour in RGB format
    Dim lngCounterX     As Long         'scan the rows of the bitmap
    Dim lngCounterY     As Long         'scan the columns of the bitmap
    Dim intPxlHeight    As Integer      'the pixel height in twips
    Dim intPxlWidth     As Integer      'the pixel width in twips
    Dim lngBlendCol     As Long         'the blended colour calculated from the two pixel colours of the bitmaps
    Dim lngCol1         As Long         'used to store a pixel colour in Long format
    Dim lngCol2         As Long         'used to store a pixel colour in Long format
    Dim lngMaskCol      As Long         'used to store a pixel colour in Long format
    
    'first convert the passed values if they
    'need to be converted.
    If intMeasurement = InTwips Then
        'get the pixel intHeight and intWidth
        'values per twips in the current
        'screen resolution.
        intPxlHeight = Screen.TwipsPerPixelY
        intPxlWidth = Screen.TwipsPerPixelX
        
        'start converting the twip values to pixels
        intDestX = intDestX / intPxlWidth
        intWidth = intWidth / intPxlWidth
        lngPic1X = lngPic1X / intPxlWidth
        lngPic2X = lngPic2X / intPxlWidth
        intDestY = intDestY / intPxlHeight
        intHeight = intHeight / intPxlHeight
        lngPic1Y = lngPic1Y / intPxlHeight
        lngPic2Y = lngPic2Y / intPxlHeight
    End If
    
    'validate the sngBlendAmount parameter.
    'It must be a values between 0 and
    '1. If the parameter is outside these
    'bounds, then round to nearist
    'bounding value (0 or 1)
    Select Case sngBlendAmount
    Case Is >= 1
        sngBlendAmount = 1
        
        'just copy the picture instead
        'of trying to blend it
        lngResult = BitBlt(lngDesthDc, _
                           intDestX, _
                           intDestY, _
                           intWidth, _
                           intHeight, _
                           lngPic2hDc, _
                           lngPic2X, _
                           lngPic2Y, _
                           P_COPY)
        Exit Sub
    Case Is <= 0
        sngBlendAmount = 0
        
        'just copy the picture instead
        'of trying to blend it
        lngResult = BitBlt(lngDesthDc, _
                           intDestX, _
                           intDestY, _
                           intWidth, _
                           intHeight, _
                           lngPic1hDc, _
                           lngPic1X, _
                           lngPic1Y, _
                           P_COPY)
        Exit Sub
    End Select
    
    'create a temperory destination
    'bitmap
    TempBmp.Area.Right = intWidth
    TempBmp.Area.Bottom = intHeight
    Call CreateNewBitmap(TempBmp.lngDC, _
                         TempBmp.lngBitmap, _
                         TempBmp.lngPointer, _
                         TempBmp.Area, lngDesthDc)
    
    'start going through the 2 source
    'bitmaps and blending the colours.
    For lngCounterX = 0 To intWidth
        For lngCounterY = 0 To intHeight
            'get the pixel colours
            lngCol1 = GetPixel(lngPic1hDc, _
                               lngPic1X + lngCounterX, _
                               lngPic1Y + lngCounterY)
            lngCol2 = GetPixel(lngPic2hDc, _
                               lngPic2X + lngCounterX, _
                               lngPic2Y + lngCounterY)
            
            'if a blend mask has been specified,
            'then get the blend amount
            'from the bitmap.
            If lngMaskhDc <> 0 Then
                lngMaskCol = GetPixel(lngMaskhDc, _
                                      lngCounterX, _
                                      lngCounterY)
                
                'convert the long value into
                'the blend amount
                MaskCol.Blue = lngMaskCol \ 65536
                MaskCol.Green = ((lngMaskCol - (MaskCol.Blue * 65536)) \ 256)
                MaskCol.Red = (lngMaskCol - (MaskCol.Blue * 65536) - _
                               (MaskCol.Green * 256))
                
                'now convert rgb value to
                'value between 0 and 1
                '(divide by 3 for the average
                'rgb and 255 to a value between
                '1 and 0 (3 * 255 = 765) )
                sngBlendAmount = (MaskCol.Red + MaskCol.Green + MaskCol.Blue) \ 765
            End If
            
            'convert long values to rgb values
            Col1.Blue = lngCol1 \ 65536
            Col1.Green = ((lngCol1 - (Col1.Blue * 65536)) \ 256)
            Col1.Red = (lngCol1 - (Col1.Blue * 65536) - (Col1.Green * 256))
            Col2.Blue = lngCol2 \ 65536
            Col2.Green = ((lngCol2 - (Col2.Blue * 65536)) \ 256)
            Col2.Red = (lngCol2 - (Col2.Blue * 65536) - (Col2.Green * 256))
    
            'average the colours by blend amount
            If (Col1.Red <> Col2.Red) _
               Or (Col1.Green <> Col2.Green) _
               Or (Col1.Blue <> Col2.Blue) Then
                
                BlendCol.Red = Col1.Red - ((Col1.Red - Col2.Red) * sngBlendAmount)
                BlendCol.Green = Col1.Green - ((Col1.Green - Col2.Green) * sngBlendAmount)
                BlendCol.Blue = Col1.Blue - ((Col1.Blue - Col2.Blue) * sngBlendAmount)
            Else
                'there is no point in blending
                'colours that are the same
                BlendCol = Col1
            End If
            
            'convert the BlendCol RGB values
            'to a long
            lngBlendCol = (CLng(BlendCol.Blue) * 65536) + _
                          (CLng(BlendCol.Green) * 256) + _
                          BlendCol.Red
            
            'set the corresponding pixel colour
            'on the temperory bitmap
            lngResult = SetPixel(TempBmp.lngDC, _
                                 lngCounterX, _
                                 lngCounterY, _
                                 lngBlendCol)
        Next lngCounterY
    Next lngCounterX
    
    'copy the blended picture to the
    'destination bitmap
    lngResult = BitBlt(lngDesthDc, _
                       intDestX, _
                       intDestY, _
                       intWidth, _
                       intHeight, _
                       TempBmp.lngDC, _
                       0, _
                       0, _
                       P_COPY)
    
    'remove the temperory bitmap
    'from memory
    Call DeleteBitmap(TempBmp.lngDC, _
                      TempBmp.lngBitmap, _
                      TempBmp.lngPointer)
End Sub

Public Sub FlipBitmap(ByVal intDestX As Integer, _
                      ByVal intDestY As Integer, _
                      ByVal intWidth As Integer, _
                      ByVal intHeight As Integer, _
                      ByVal lngSourcehDc As Long, _
                      ByVal intSourceX As Integer, _
                      ByVal intSourceY As Integer, _
                      Optional ByVal Orientation As FlipType = FlipHorizontally, _
                      Optional ByVal udtMeasurement As Scaling = InPixels)

    'This procedure will flip a picture either
    'horizontally or vertically. It copies
    'the bitmap either row by row or column
    'by column to improve speed.
    
    Dim intPxlHeight    As Integer      'the height in twips of a pixel
    Dim intPxlWidth     As Integer      'the width in twips of a pixel
    Dim intCounter      As Integer      'used to scan through each row/column
    Dim TempBmp         As BitmapStruc  'holds a temperory bitmap to copy the rows/columns
    Dim intFinish       As Integer      'the row/column to finish at
    Dim lngResult       As Long         'holds any returned error value from an api call
    
    'convert the twips to pixel values if necessary
    If udtMeasurement = InTwips Then
        intPxlWidth = Screen.TwipsPerPixelX
        intPxlHeight = Screen.TwipsPerPixelY
        
        intDestX = intDestX / intPxlWidth
        intWidth = intWidth / intPxlWidth
        intSourceX = intSourceX / intPxlWidth
        intDestY = intDestY / intPxlHeight
        intHeight = intHeight / intPxlHeight
        intSourceY = intSourceY / intPxlHeight
    End If
    
    'create the temperory bitmap
    TempBmp.Area.Right = intWidth
    TempBmp.Area.Bottom = intHeight
    Call CreateNewBitmap(TempBmp.lngDC, _
                         TempBmp.lngBitmap, _
                         TempBmp.lngPointer, _
                         TempBmp.Area, _
                         lngSourcehDc)
    
    'define the bounds of the loop depending on the orientation (do I scan
    'the bitmap row by row or column by column)
    Select Case Orientation
    Case FlipHorizontally
        'scan column by column
        intFinish = intWidth - 1
    Case FlipVertically
        'scan row by row
        intFinish = intHeight - 1
    End Select
    
    For intCounter = 0 To intFinish
        'copy the row or column into the appropiate section of the bitmap
        If Orientation = FlipHorizontally Then
            'horizontal
            lngResult = BitBlt(TempBmp.lngDC, _
                               intFinish - intCounter, _
                               0, _
                               1, _
                               intHeight, _
                               lngSourcehDc, _
                               intSourceX + intCounter, _
                               intSourceY, _
                               P_COPY)
        Else
            'flip vertically
            lngResult = BitBlt(TempBmp.lngDC, _
                               0, _
                               intFinish - intCounter, _
                               intWidth, _
                               1, _
                               lngSourcehDc, _
                               intSourceX, _
                               intSourceY + intCounter, _
                               P_COPY)
        End If
    Next
    
    'copy the flipped bitmap onto the destination bitmap
    lngResult = BitBlt(mudtBitmap.lngDC, _
                       intDestX, _
                       intDestY, _
                       TempBmp.Area.Right, _
                       TempBmp.Area.Bottom, _
                       TempBmp.lngDC, _
                       0, _
                       0, _
                       P_COPY)
                       
    'delete the temperory bitmap
    Call DeleteBitmap(TempBmp.lngDC, _
                      TempBmp.lngBitmap, _
                      TempBmp.lngPointer)
End Sub

Public Sub RotateBitmap(ByVal lngSourcehDc As Long, _
                        ByVal Rotate As RotateType, _
                        ByVal intDestX As Integer, _
                        ByVal intDestY As Integer, _
                        ByVal intSourceX As Integer, _
                        ByVal intSourceY As Integer, _
                        ByVal intWidth As Integer, _
                        ByVal intHeight As Integer, _
                        Optional ByVal udtMeasurement As Scaling = InPixels)

    'This procedure will rotate a bitmap 90, 180 or 270 degrees.
    
    Dim lngResult       As Long         'holds any returned error value from an api call
    Dim intPxlWidth     As Integer      'holds the width in twips of a pixel
    Dim intPxlHeight    As Integer      'holds the height in twips of a pixel
    Dim intCounterX     As Integer      'used to scan through the horizontal pixels
    Dim intCounterY     As Integer      'used to scan through the vertical pixels
    Dim TempBmp         As BitmapStruc  'holds a bitmap to copy the pixels onto
    Dim lngBitCol       As Long         'holds the colour value of a single pixel
    
    'convert twips values to pixels if necessary
    If udtMeasurement = InTwips Then
        intPxlHeight = Screen.TwipsPerPixelY
        intPxlWidth = Screen.TwipsPerPixelX
        
        'convert values
        intDestX = intDestX / intPxlWidth
        intSourceX = intSourceX / intPxlWidth
        intWidth = intWidth / intPxlWidth
        intDestY = intDestY / intPxlHeight
        intSourceY = intSourceY / intPxlHeight
        intHeight = intHeight / intPxlHeight
    End If
    
    'create a temperory bitmap to draw on
    If Rotate = Rotate180 Then
        'the intWidth and intHeight dimensions are the same
        TempBmp.Area.Bottom = intHeight
        TempBmp.Area.Right = intWidth
    Else
        'rotate the dimensions 90 degrees
        TempBmp.Area.Bottom = intWidth
        TempBmp.Area.Right = intHeight
    End If
    Call CreateNewBitmap(TempBmp.lngDC, _
                         TempBmp.lngBitmap, _
                         TempBmp.lngPointer, _
                         TempBmp.Area, _
                         mudtBitmap.lngDC)
    
    Select Case Rotate
    Case RotateRight To RotateLeft
        'rotate bitmap right or left
        For intCounterX = 0 To intWidth
            For intCounterY = 0 To intHeight
                'get the pixel colour
                lngBitCol = GetPixel(lngSourcehDc, _
                                     intSourceX + intCounterX, _
                                     intSourceY + intCounterY)
                
                'copy to appropiate part of the temperory bitmap
                If Rotate = RotateRight Then
                    'rotate right
                    lngResult = SetPixel(TempBmp.lngDC, _
                                         intHeight - intCounterY, _
                                         intCounterX, _
                                         lngBitCol)
                Else
                    'rotate left
                    lngResult = SetPixel(TempBmp.lngDC, _
                                         intCounterY, _
                                         intHeight - intCounterX, _
                                         lngBitCol)
                End If
            Next intCounterY
        Next intCounterX
    
    Case Rotate180
        'rotate bitmap 180 degrees
        
        'we rotate the bitmap 180 degrees by flipping it vertically and
        'horizontally. This is done fastest by calling the FlipBitmap procedure
        Call FlipBitmap(TempBmp.lngDC, _
                        0, _
                        0, _
                        intWidth, _
                        intHeight, _
                        lngSourcehDc, _
                        intSourceX, _
                        intSourceY)
        Call FlipBitmap(TempBmp.lngDC, _
                        0, _
                        0, _
                        intWidth, _
                        intHeight, _
                        TempBmp.lngDC, _
                        0, _
                        0, _
                        FlipVertically)
    End Select
    
    'copy the temperory bitmap to the destination Dc at the specified
    'co-ordinates
    lngResult = BitBlt(mudtBitmap.lngDC, _
                       intDestX, _
                       intDestY, _
                       TempBmp.Area.Right, _
                       TempBmp.Area.Bottom, _
                       TempBmp.lngDC, _
                       0, _
                       0, _
                       P_COPY)
    
    'remove the temperory bitmap from memory and exit
    Call DeleteBitmap(TempBmp.lngDC, _
                      TempBmp.lngBitmap, _
                      TempBmp.lngPointer)
End Sub

Public Sub DrawRect(Optional ByVal lngColour As Long = NO_COLOUR, _
                    Optional ByVal intTop As Integer, _
                    Optional ByVal intLeft As Integer, _
                    Optional ByVal intHeight As Integer = -1, _
                    Optional ByVal intWidth As Integer = -1, _
                    Optional ByVal udtMeasurement As Scaling = InPixels, _
                    Optional ByVal lngStyle As Long = BS_SOLID, _
                    Optional ByVal lngPattern As Long = HS_SOLID)
    
    'this draws a rectangle using the co-ordinates
    'and lngColour given.
    
    Dim StartRect   As Rect     'holds the dimensions of the rectangle to be drawn
    Dim lngResult   As Long     'holds any returned error value from the api calls
    Dim lngJunk     As Long     'holds any junk information from an api call
    Dim lnghBrush   As Long     'holds a pointer to a brush with the colour and pattern information for the rectangle
    Dim BrushStuff  As LogBrush 'holds the brush information
    
    'check the default height and widths
    With mudtBitmap.Area
        If intHeight < 0 Then
            intHeight = .Bottom - .Top
        End If
        If intWidth < 0 Then
            intWidth = .Right - .Left
        End If
    End With
    
    'initalise values
    StartRect.Top = intTop
    StartRect.Left = intLeft
    StartRect.Bottom = intLeft + intWidth
    StartRect.Right = intTop + intHeight
    
    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert to pixels
        With StartRect
            .Left = .Left / Screen.TwipsPerPixelX
            .Top = .Top / Screen.TwipsPerPixelY
            .Right = .Right / Screen.TwipsPerPixelX
            .Bottom = .Bottom / Screen.TwipsPerPixelY
        End With
    End If
    
    'set the colour to the background colour if
    'none specified
    If lngColour = NO_COLOUR Then
        lngColour = mlngBackColour
    End If
    
    'create a brush
    BrushStuff.lbColor = lngColour
    BrushStuff.lbHatch = lngPattern
    BrushStuff.lbStyle = lngStyle
    
    'apply the brush to the device context
    lnghBrush = CreateBrushIndirect(BrushStuff)
    lnghBrush = SelectObject(mudtBitmap.lngDC, lnghBrush)
    
    'draw a rectangle
    lngResult = PatBlt(mudtBitmap.lngDC, _
                       intLeft, _
                       intTop, _
                       intWidth, _
                       intHeight, _
                       PATCOPY)
    
    'A "Brush" object was created. It must be removed from memory.
    lngJunk = SelectObject(mudtBitmap.lngDC, lnghBrush)
    lngJunk = DeleteObject(lngJunk)
End Sub

Public Sub DrawRoundRect(ByVal lngColour As Long, _
                         ByVal intLeft As Integer, _
                         ByVal intTop As Integer, _
                         ByVal intRight As Integer, _
                         ByVal intBottom As Integer, _
                         ByVal intEdgeRadius As Integer, _
                         Optional ByVal udtMeasurement As Scaling = InPixels, _
                         Optional ByVal lngStyle As Long = BS_SOLID, _
                         Optional ByVal lngPattern As Long = HS_SOLID)
                         
    'this draws a rectangle using the co-ordinates
    'and lngColour given.
    
    Const Width = 1 'the pixel width of the edge
    
    Dim lnghPen     As Long     'holds a pointer to a Pen object
    Dim PenStuff    As LogPen   'holds the pen information
    Dim lnghBrush   As Long     'holds a pointer to a Brush object
    Dim BrushStuff  As LogBrush 'holds the brush information
    Dim lngJunk     As Long     'holds "Junk" information from an api call
    Dim lngResult   As Long     'holds any returned error value from an api call
    Dim OffsetRect  As Rect     'holds how much to grow/shrink the original size to accomidate the border (if any)
    
    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert twip values to pixels
        intLeft = intLeft / Screen.TwipsPerPixelX
        intTop = intTop / Screen.TwipsPerPixelX
        intRight = intRight / Screen.TwipsPerPixelY
        intBottom = intBottom / Screen.TwipsPerPixelY
        intEdgeRadius = intEdgeRadius / Screen.TwipsPerPixelX
    End If
    
    'Find out if a specific lngColour is
    'to be set. If so set it.
    
    'set lnghBrush settings (similar to
    'the FillColor property)
    BrushStuff.lbColor = lngColour
    BrushStuff.lbStyle = lngStyle
    BrushStuff.lbHatch = lngPattern  'ignored if lngStyle is solid
    
    'set lnghPen settings (similar to the
    'border properties on controls)
    PenStuff.lopnColor = lngColour
    PenStuff.lopnWidth.X = Width
    PenStuff.lopnStyle = PS_SOLID
    
    'apply the settings to the device context
    lnghPen = CreatePenIndirect(PenStuff)
    lnghPen = SelectObject(mudtBitmap.lngDC, lnghPen)
    lnghBrush = CreateBrushIndirect(BrushStuff)
    lnghBrush = SelectObject(mudtBitmap.lngDC, lnghBrush)
    
    '---------in debug - 18/02/02
    'deflate the rectangle dimensions by
    'the radius amount
    'OffsetRect.intLeft = intLeft
    'OffsetRect.intTop = intTop
    'OffsetRect.intRight = intRight
    'OffsetRect.intBottom = intBottom
    'lngResult = InflateRect(OffsetRect, -intEdgeRadius, -intEdgeRadius)
    
    'draw the rounded rectangle
    lngResult = RoundRect(mudtBitmap.lngDC, _
                          OffsetRect.Left, _
                          OffsetRect.Top, _
                          OffsetRect.Right, _
                          OffsetRect.Bottom, _
                          intEdgeRadius, _
                          intEdgeRadius)
    
    'delete the objects created (lnghPen and lnghBrush objects)
    lngJunk = SelectObject(mudtBitmap.lngDC, lnghPen)
    lngJunk = DeleteObject(lngJunk)
    lngJunk = SelectObject(mudtBitmap.lngDC, lnghBrush)
    lngJunk = DeleteObject(lngJunk)
End Sub

Public Function TitleToTray(ByRef frm As Form)
    'This function will draw the minimize animation from the form to
    'the title tray.
    
    'find the position of the title tray
    mlnghStartMenu = FindWindow("Shell_TrayWnd", vbNullString)
    mlnghChild = GetWindow(mlnghStartMenu, GW_CHILD)
    Do
        mlngClassNameLen = GetClassName(mlnghChild, _
                                       mstrClass, _
                                       Len(mstrClass))
        If InStr(1, mstrClass, "TrayNotifyWnd") Then
            mlnghTray = mlnghChild
            Exit Do
        End If
        mlnghChild = GetWindow(mlnghChild, GW_HWNDNEXT)
    Loop
    
    'animate the title bar to the title tray
    mlngRetVal = GetWindowRect(frm.hwnd, mudtRectFrom)
    mlngRetVal = GetWindowRect(mlnghTray, mudtRectTo)
    mlngRetVal = DrawAnimatedRects(frm.hwnd, _
                                  IDANI_OPEN Or IDANI_CAPTION, _
                                  mudtRectFrom, _
                                  mudtRectTo)
    
    'hide form
    frm.Visible = False
    frm.Hide
End Function

Public Function TrayToTitle(ByRef frm As Form)
    'This function draws the restore animation of the forms title bar, from
    'the system tray to the forms' position.
    
    'find the system trays position
    mlnghStartMenu = FindWindow("Shell_TrayWnd", _
                                vbNullString)
    mlnghChild = GetWindow(mlnghStartMenu, GW_CHILD)
    Do
        mlngClassNameLen = GetClassName(mlnghChild, _
                                       mstrClass, _
                                       Len(mstrClass))
        
        If InStr(1, mstrClass, "TrayNotifyWnd") Then
            mlnghTray = mlnghChild
            Exit Do
        End If
        
        mlnghChild = GetWindow(mlnghChild, _
        GW_HWNDNEXT)
    Loop
    
    'draw the animation
    mlngRetVal = GetWindowRect(frm.hwnd, mudtRectFrom)
    mlngRetVal = GetWindowRect(mlnghTray, mudtRectTo)
    mlngRetVal = DrawAnimatedRects(frm.hwnd, _
                                  IDANI_CLOSE Or IDANI_CAPTION, _
                                  mudtRectTo, _
                                  mudtRectFrom)
    
    'show the window
    frm.Visible = True
    frm.Show
End Function

Public Sub DrawLine(ByVal intX1 As Integer, _
                    ByVal intY1 As Integer, _
                    ByVal intX2 As Integer, _
                    ByVal intY2 As Integer, _
                    Optional ByVal lngColour As Long = vbBlack, _
                    Optional ByVal intWidth As Integer = 1, _
                    Optional ByVal udtMeasurement As Scaling = InPixels)
                    
    'This will draw a line from point1 to point2
    
    Const NUM_OF_POINTS = 2
    
    Dim lngResult               As Long     'holds any returned error value from an api call
    Dim lnghPen                 As Long     'holds a pointer to a Pen object
    Dim lngPointer              As Long     'holds the pointer
    Dim PenStuff                As LogPen   'holds the Pen object information
    Dim lngJunk                 As Long     'holds any useless information returned from an api call
    Dim Points(NUM_OF_POINTS)   As PointAPI 'holds the co-ordinates of the line to be drawn
    
    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert twip values to pixels
        intX1 = intX1 / Screen.TwipsPerPixelX
        intX2 = intX2 / Screen.TwipsPerPixelX
        intY1 = intY1 / Screen.TwipsPerPixelY
        intY2 = intY2 / Screen.TwipsPerPixelY
    End If
    
    'Find out if a specific lngColour is to be set. If so set it.
    PenStuff.lopnColor = lngColour
    PenStuff.lopnStyle = PS_GEOMETRIC
    PenStuff.lopnWidth.X = intWidth
    
    'apply the pen settings to the device context
    lnghPen = CreatePenIndirect(PenStuff)
    lnghPen = SelectObject(mudtBitmap.lngDC, lnghPen)
    
    'set the points
    Points(1).X = intX1
    Points(1).Y = intY1
    Points(2).X = intX2
    Points(2).Y = intY2
    
    'draw the line
    lngResult = Polyline(mudtBitmap.lngDC, Points(1), NUM_OF_POINTS)
    
    'A "Pen" object was created. It must be removed from memory.
    lngJunk = SelectObject(mudtBitmap.lngDC, lnghPen)
    lngJunk = DeleteObject(lngJunk)
End Sub

Public Sub DrawPoly(Optional ByVal lngFillColour As Long = 0, _
                    Optional ByVal lngBorderColour As Long = 0, _
                    Optional ByVal intBorderWidth As Integer = 1, _
                    Optional ByVal udtMeasurement As Scaling = InPixels)
                    
    'This function will draw a polygon of size and colours specified.
    'Modified to draw and fill all the points specified in the array,
    'PolygonPoints(). Use;
    '
    '"ReDim Preserve PolygonPoints(UBound(PolygonPoints) + 1)"
    '
    'To add a point to the array, then specifiy the X and Y values
    'of your point.
    '
    'To delete a point from the array, use the same as above
    'except replace the + sign with a - sign.
    '
    'To clear the array of all data, use;
    '
    '"ReDim PolygonPoints(0)"
    '
    'Also see AddToPoly, DelFromPoly and ClearPoly
    '----------------
    '- Eric  15/12/2001
    '----------------
    
    Dim lngSuccessful   As Long
    Dim intCounter      As Integer
    Dim Temp()          As PointAPI
    Dim intPointNum     As Integer
    Dim lnghBrush       As Long
    Dim BrushStuff      As LogBrush
    Dim lnghPen         As Long
    Dim PenStuff        As LogPen
    
    'find the number of points stored
    intPointNum = UBound(PolygonPoints)
    
    'convert all points to pixels if necessary, otherwise, don't waste time.
    If udtMeasurement = InTwips Then
        'create a temperory array to hold the points
        ReDim Temp(intPointNum)
        
        'convert to pixels
        For intCounter = 0 To intPointNum
            Temp(intCounter).X = PolygonPoints(intCounter).X / Screen.TwipsPerPixelX
            Temp(intCounter).Y = PolygonPoints(intCounter).Y / Screen.TwipsPerPixelY
        Next intCounter
    End If
    
    'apply the border and background colours
    BrushStuff.lbColor = lngFillColour
    BrushStuff.lbHatch = 0
    BrushStuff.lbStyle = BS_SOLID
    
    PenStuff.lopnColor = lngBorderColour
    PenStuff.lopnWidth.X = intBorderWidth
    PenStuff.lopnStyle = PS_SOLID
    
    'create the objects necessary to apply the colour and draw settings
    lnghBrush = CreateBrushIndirect(BrushStuff)
    lnghPen = CreatePenIndirect(PenStuff)
    lngSuccessful = SelectObject(mudtBitmap.lngDC, lnghBrush)
    lngSuccessful = SelectObject(mudtBitmap.lngDC, lnghPen)
    
    'draw the polygon
    If udtMeasurement = InTwips Then
        lngSuccessful = Polygon(mudtBitmap.lngDC, Temp(0), intPointNum)
    Else
        lngSuccessful = Polygon(mudtBitmap.lngDC, PolygonPoints(0), intPointNum)
    End If
    
    'remove the pen and brush objects
    lngSuccessful = SelectObject(mudtBitmap.lngDC, lnghBrush)
    lngSuccessful = DeleteObject(lngSuccessful)
    lngSuccessful = SelectObject(mudtBitmap.lngDC, lnghPen)
    lngSuccessful = DeleteObject(lngSuccessful)
End Sub

Private Sub ClearPoly()
    'This will reset the polgon array to contain no values, with no elements
    '(except element zero by default).
    'Also see DrawPoly, AddToPoly and DelFromPoly
    
    ReDim PolygonPoints(0, 0)
    
    'find the amount of points in the second array dimension
    'clear the points
    PolygonPoints(0).X = 0
    PolygonPoints(0).Y = 0
End Sub

Private Sub AddToPoly(ByVal intX As Integer, _
                     ByVal intY As Integer, _
                     Optional ByVal Point As Integer = -1)
                     
    'This procedure will add a point to the polygon at the specified
    'point. if no point os specified (point -1), then the point is added to
    'the end of the polygon array. It is best to call ClearPoly before using
    'AddToPoly for the first time, but it is not necessary as the appropiate
    'check is made here.
    'Also see DrawPoly, DelFromPoly and ClearPoly
    
    Dim Counter         As Integer
    Dim MaintXPoints    As Integer
    Dim ErrNum          As Integer
    
    'check for error 9 "subscript out of range". This means that the array
    'has not been initalized yet (there are no elements in the array).
    On Error Resume Next
    MaintXPoints = UBound(PolygonPoints)
    
    'trap error
    ErrNum = Err
    
    'resumt normal error handling
    On Error GoTo 0
    
    'set the array if it has not been initialized yet.
    If ErrNum <> 0 Then
        Call ClearPoly
    End If
    
    'eintXpand the array by one
    ReDim Preserve PolygonPoints(MaintXPoints + 1)
    MaintXPoints = MaintXPoints + 1
    
    
    'if no insert point was specified, then add the point to the end of
    'the current maintX array element number
    If Point < 0 Then
        'enter the new values
        PolygonPoints(MaintXPoints).X = intX
        PolygonPoints(MaintXPoints).Y = intY
    Else
        'insert the point into the array at the specified point, moving any
        'consecutive points "up".
        
        'move consecutive values up
        For Counter = Point To (MaintXPoints - 1)
            PolygonPoints(Counter + 1) = PolygonPoints(Counter)
        Next Counter
        
        'enter the new values
        PolygonPoints(Point).X = intX
        PolygonPoints(Point).Y = intY
    End If
End Sub

Private Sub DelFromPoly(Optional ByVal intPoint As Integer = -1)
    'This will remove the specified Point from the polygon array. if
    'no Point was specified (intPoint -1), then the procedure will remove
    'the last array position
    'Also see DrawPoly, AddToPoly and ClearPoly
    
    Dim intCounter      As Integer
    Dim intMaxPoints    As Integer
    
    'find the number of points
    intMaxPoints = UBound(PolygonPoints)
    
    'if there is only one intPoint, then clear the array and exit
    If intMaxPoints < 2 Then
        Call ClearPoly
        Exit Sub
    End If
    
    'move consecutive points in the array "down" on top of the specified
    'intPoint
    If intPoint >= 0 Then
        For intCounter = intPoint To (intMaxPoints - 1)
            PolygonPoints(intCounter) = PolygonPoints(intCounter + 1)
        Next intCounter
    End If
    
    'shrink the array size by one
    ReDim Preserve PolygonPoints(intMaxPoints - 1)
End Sub

Public Sub LockWindow(ByVal FormName As Form)
    'Prevent the form from updating its display
    
    Dim lngResult As Boolean    'holds the returned error value from an api call
    
    lngResult = LockWindowUpdate(FormName.hwnd)
End Sub

Public Sub UnLockWindow()
    'see the procedure LockWindow
    
    Dim lngResult As Long   'holds the returned error value from an api call
    
    'Let the form update its display
    lngResult = LockWindowUpdate(0)
End Sub

Private Sub CreateNewBitmap(ByRef lngDC As Long, _
                            ByRef lngBitmap As Long, _
                            ByRef lngPointer As Long, _
                            ByRef udtBmpArea As Rect, _
                            Optional ByVal lngCompatableWithhDc As Long = 0, _
                            Optional ByVal udtMeasurement As Scaling = InPixels)
                           
    'This procedure will create a new bitmap compatable with a given
    'form (you will also be able to then use this bitmap in a picturebox).
    'The space specified in "udtArea" should be in "Twips" and will be
    'converted into pixels in the following code.
    
    Dim lngResult           As Long             'holds any returned error value from the api calls
    Dim udtArea             As Rect             'holds the size of the bitmap
    'Dim udtColourAdjustment As COLORADJUSTMENT  'sets the colour adjustment values when StretchBlt is used
    
    'scale the bitmap points if necessary
    udtArea = udtBmpArea
    If udtMeasurement = InTwips Then
        Call RectToPixels(udtArea)
    End If
    
    'create the bitmap and its references
    If lngCompatableWithhDc = 0 Then
        'use the Dc for the desktop window
        lngCompatableWithhDc = GetDC(0)
        lngDC = CreateCompatibleDC(lngCompatableWithhDc)
        
        'create the bitmap
        lngBitmap = CreateCompatibleBitmap(lngCompatableWithhDc, _
                                           (udtArea.Right - udtArea.Left), _
                                           (udtArea.Bottom - udtArea.Top))
        
        lngResult = ReleaseDC(0, lngCompatableWithhDc)
    Else
        'make the dc compatable with the specifed window
        lngDC = CreateCompatibleDC(lngCompatableWithhDc)
    
        'create the bitmap
        lngBitmap = CreateCompatibleBitmap(lngCompatableWithhDc, _
                                           (udtArea.Right - udtArea.Left), _
                                           (udtArea.Bottom - udtArea.Top))
    End If
    
    'apply the bitmap to the device context
    lngPointer = SelectObject(lngDC, lngBitmap)
    
'    'set the colour adjustment values
'    With udtColourAdjustment
'
'    End With
    
    'set default colours and clear bitmap to selected colour
    'lngResult = SetColorAdjustment(lngDC, udtColourAdjustment)
    lngResult = SetBkMode(lngDC, OPAQUE)
    lngResult = SetBkColor(lngDC, mlngBackColour)
    Call DrawRect
End Sub

Private Sub DeleteBitmap(ByRef lngDC As Long, _
                         ByRef lngBitmap As Long, _
                         ByRef lngPointer As Long)
                        
    'This will remove the bitmap that stored what was displayed before
    'the text was written to the screen, from memory.
    
    Dim lngJunk As Long     'holds useless information from an api call
    
    If lngDC = 0 Then
        'there is nothing to delete. Exit the sub-routine
        Exit Sub
    End If
    
    'delete the device context
    lngJunk = SelectObject(lngDC, lngPointer)
    lngJunk = DeleteObject(lngBitmap)
    lngJunk = DeleteDC(lngDC)
    
    'show that the device context has been deleted by setting
    'all parameters passed to the procedure to zero
    lngDC = 0
    lngBitmap = 0
    lngPointer = 0
End Sub

Public Sub MergeBitmaps(ByVal hDcTextureBack As Long, _
                        ByVal hDcTextureFore As Long, _
                        ByVal hDcMask As Long, _
                        ByVal intDestX As Integer, _
                        ByVal intDestY As Integer, _
                        ByVal intTextureBackX As Integer, _
                        ByVal intTextureBackY As Integer, _
                        ByVal intTextureForeX As Integer, _
                        ByVal intTextureForeY As Integer, _
                        ByVal intMaskWidth As Integer, _
                        ByVal intMaskHeight As Integer, _
                        Optional ByVal udtMeasurement As Scaling = InPixels)
                        
    'This procedure takes a monochrome (black & white) bitmap as a mask,
    '2 source texture bitmaps and a destination bitmap. It copies to the
    'destination bitmap a merge of the two source bitmaps, filling in the
    'pixels according to the mask. For each white pixel in the mask, it copies
    'the corresponding pixel from TextureBack. For each black pixel in the mask,
    'it copies the corresponding pixel from TextureFore to the destination
    'bitmap.
    'eg. Say the mask picture was a black square with a white letter "A" on
    'it. The first texture was of clouds and the second picture was of a
    'tartan design. The lngResult would be a cloud picture in the shape of an
    '"A" on a tartan background.
    
    Dim lngResult   As Long         'holds the returned error value from an api call
    Dim TempMaskBmp As BitmapStruc  'holds the mask bitmap
    Dim TempBackBmp As BitmapStruc  'holds the bitmap you wish to use as a background
    Dim intConvertX As Integer      'holds the height in twips of a pixel
    Dim intConvertY As Integer      'holds the width in twips of a pixel
    
    'convert passed values if necessary
    If udtMeasurement = InTwips Then
        intConvertX = Screen.TwipsPerPixelX
        intConvertY = Screen.TwipsPerPixelY
        
        intDestX = intDestX / intConvertX
        intDestY = intDestY / intConvertY
        intTextureBackX = intTextureBackX / intConvertX
        intTextureBackY = intTextureBackY / intConvertY
        intTextureForeX = intTextureForeX / intConvertX
        intTextureForeY = intTextureForeY / intConvertY
        intMaskWidth = intMaskWidth / intConvertX
        intMaskHeight = intMaskHeight / intConvertY
    End If
    
    'create temperory bitmaps
    TempMaskBmp.Area.Right = intMaskWidth
    TempMaskBmp.Area.Bottom = intMaskHeight
    TempBackBmp.Area = TempMaskBmp.Area
    Call CreateNewBitmap(TempMaskBmp.lngDC, _
                         TempMaskBmp.lngBitmap, _
                         TempMaskBmp.lngPointer, _
                         TempMaskBmp.Area, _
                         hDcMask, _
                         InPixels)
    Call CreateNewBitmap(TempBackBmp.lngDC, _
                         TempBackBmp.lngBitmap, _
                         TempBackBmp.lngPointer, _
                         TempBackBmp.Area, _
                         hDcMask, _
                         InPixels)
    
    'create a white bitmap with a mask shaped hole onto the
    'destination background
    lngResult = BitBlt(TempMaskBmp.lngDC, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       hDcTextureFore, _
                       intTextureForeX, _
                       intTextureForeY, _
                       P_COPY)
    lngResult = BitBlt(TempBackBmp.lngDC, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       hDcMask, _
                       0, _
                       0, _
                       P_INVERT)
    lngResult = BitBlt(TempMaskBmp.lngDC, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       TempBackBmp.lngDC, _
                       0, _
                       0, _
                       P_MERGE_PAINT)
    
    'draw a white mask shape onto the second texture
    'where the black mask is
    lngResult = BitBlt(TempBackBmp.lngDC, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       hDcTextureBack, _
                       intTextureBackX, _
                       intTextureBackY, _
                       P_COPY)
    lngResult = BitBlt(TempBackBmp.lngDC, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       hDcMask, _
                       0, _
                       0, _
                       P_MERGE_PAINT)
    
    'merge the two masks
    lngResult = BitBlt(TempMaskBmp.lngDC, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       TempBackBmp.lngDC, _
                       0, _
                       0, _
                       P_AND)
    
    'copy the merged lngResult to the destination
    lngResult = BitBlt(mudtBitmap.lngDC, _
                       intDestX, _
                       intDestY, _
                       intMaskWidth, _
                       intMaskHeight, _
                       TempMaskBmp.lngDC, _
                       0, _
                       0, _
                       P_COPY)
    
    'remove the two temperory bitmaps from memory
    Call DeleteBitmap(TempMaskBmp.lngDC, _
                      TempMaskBmp.lngBitmap, _
                      TempMaskBmp.lngPointer)
    Call DeleteBitmap(TempBackBmp.lngDC, _
                      TempBackBmp.lngBitmap, _
                      TempBackBmp.lngPointer)
End Sub

Private Sub RectToTwips(ByRef udtRect As Rect)
    'converts pixels to twips in a rect structure
    
    With udtRect
        .Left = .Left * Screen.TwipsPerPixelX
        .Right = .Right * Screen.TwipsPerPixelX
        .Top = .Top * Screen.TwipsPerPixelY
        .Bottom = .Bottom * Screen.TwipsPerPixelY
    End With
End Sub

Private Sub RectToPixels(ByRef udtRect As Rect)
    'converts twips to pixels in a rect structure
    
    With udtRect
        .Left = .Left / Screen.TwipsPerPixelX
        .Right = .Right / Screen.TwipsPerPixelX
        .Top = .Top / Screen.TwipsPerPixelY
        .Bottom = .Bottom / Screen.TwipsPerPixelY
    End With
End Sub

Public Function IsFormActive(ByVal frmName As Form) _
                             As Boolean
    'This function returns wether or not the window
    'is active
    
    If frmName.hwnd = GetActiveWindow Then
        IsFormActive = True
    Else
        IsFormActive = False
    End If
End Function

Public Sub Gradient(ByVal lngStartCol As Long, _
                    ByVal lngFinishCol As Long, _
                    ByVal intLeft As Integer, _
                    ByVal intTop As Integer, _
                    ByVal intWidth As Integer, _
                    ByVal intHeight As Integer, _
                    Optional ByVal Direction As GradientTo = GradHorizontal, _
                    Optional ByVal bytLineWidth As Byte = 1)
    'ByVal mudtbitmap.hdcmemory As Long, _

    'draws a gradient from colour udtStart to colour udtFinish, and assums
    'that all measurments passed to it are in pixels unless otherwise
    'specified.
    
    Dim intCounter      As Integer  'cycles through each step in the gradient
    Dim intBiggestDiff  As Integer  'the biggest difference in either R,G or B values for the start and finish (the number of steps the gradient has to go through)
    Dim udtColour       As RGBVal   'temperory colour space for each step in the gradient
    Dim udtStart        As RGBVal   'the starting colour
    Dim udtFinish       As RGBVal   'the finishing colour
    Dim sngAddRed       As Single   'temperorily holds the current Red value
    Dim sngAddGreen     As Single   'temperorily holds the current Green value
    Dim sngAddBlue      As Single   'temperorily holds the current Blue value
    Dim intColRed       As Integer  'used to get the starting and finishing colours
    Dim intColGreen     As Integer  'used to get the starting and finishing colours
    Dim intColBlue      As Integer  'used to get the starting and finishing colours
    
    'perform all necessary calculations before drawing gradient
    'such as converting long to rgb values, and getting the correct
    'scaling for the bitmap.
    With udtStart
        Call GetRGB(lngStartCol, _
                    intColRed, _
                    intColGreen, _
                    intColBlue)
        
        'store the colours
        .Red = intColRed
        .Green = intColGreen
        .Blue = intColBlue
    End With
    With udtFinish
        Call GetRGB(lngFinishCol, _
                    intColRed, _
                    intColGreen, _
                    intColBlue)
        
        'store the colours
        .Red = intColRed
        .Green = intColGreen
        .Blue = intColBlue
    End With
    
    'draw the colour gradient
    Select Case Direction
    Case GradVertical
        intBiggestDiff = intWidth
    Case GradHorizontal
        intBiggestDiff = intHeight
    End Select
    
    'calculate how much to increment/decrement each colour per step
    sngAddRed = (bytLineWidth * ((udtFinish.Red) - udtStart.Red) / intBiggestDiff)
    sngAddGreen = (bytLineWidth * ((udtFinish.Green) - udtStart.Green) / intBiggestDiff)
    sngAddBlue = (bytLineWidth * ((udtFinish.Blue) - udtStart.Blue) / intBiggestDiff)
    udtColour = udtStart
    
    'calculate the colour of each line before drawing it on the bitmap
    For intCounter = 0 To intBiggestDiff Step bytLineWidth
        'find the point between colour udtStart and udtColour udtFinish that
        'corresponds to the point between 0 and intBiggestDiff
        
        'check for overflow
        If udtColour.Red > 255 Then
            udtColour.Red = 255
        Else
            If udtColour.Red < 0 Then
                udtColour.Red = 0
            End If
        End If
        If udtColour.Green > 255 Then
            udtColour.Green = 255
        Else
            If udtColour.Green < 0 Then
                udtColour.Green = 0
            End If
        End If
        If udtColour.Blue > 255 Then
            udtColour.Blue = 255
        Else
            If udtColour.Blue < 0 Then
                udtColour.Blue = 0
            End If
        End If
        
        'draw the gradient in the proper orientation in the calculated colour
        Select Case Direction
        Case GradVertical
            Call DrawLine(intCounter + intLeft, _
                          intTop, _
                          intCounter + intLeft, _
                          intHeight + intTop, _
                          RGB(udtColour.Red, udtColour.Green, udtColour.Blue), _
                          bytLineWidth, _
                          InPixels)
        Case GradHorizontal
            Call DrawLine(intLeft, _
                          intCounter + intTop, _
                          intLeft + intWidth, _
                          intTop + intCounter, _
                          RGB(udtColour.Red, udtColour.Green, udtColour.Blue), _
                          bytLineWidth, _
                          InPixels)
        End Select
        
        'set next colour
        udtColour.Red = udtColour.Red + sngAddRed
        udtColour.Green = udtColour.Green + sngAddGreen
        udtColour.Blue = udtColour.Blue + sngAddBlue
    Next intCounter
End Sub

'not currently in use - 2/June/2002
'Private Sub FadeGradient(ByVal intDestLeft As Integer, _
'                        ByVal intDestTop As Integer, _
'                        ByVal intDestWidth As Integer, _
'                        ByVal intDestHeight As Integer, _
'                        ByVal lngGradhDc As Long, _
'                        ByVal lngStartFromA As Long, _
'                        ByVal lngFinishToA As Long, _
'                        ByVal lngStartFromB As Long, _
'                        ByVal lngFinishToB As Long, _
'                        ByVal intLeft As Integer, _
'                        ByVal intTop As Integer, _
'                        ByVal intWidth As Integer, _
'                        ByVal intHeight As Integer, _
'                        ByVal udtDirection As GradientTo, _
'                        Optional ByVal udtMesurement As Scaling = 1, _
'                        Optional ByVal bytLineWidth As Byte = 1)
'
'    'This procedure will call the Gradient function to fade it into
'    'the udtColours specified.
'    'Note : all mesurements must me of the same scale, ie they must all
'    'be in pixels or all in twips.
'
'    Dim udtColour(2) As RGBVal
'    Dim udtStart(2) As RGBVal
'    Dim udtFinish(2) As RGBVal
'    Dim lngGradCol(2) As Long
'    Dim intCounter As Integer
'    Dim intBiggestDiff As Integer
'    Dim intValue As Integer
'    Dim intIndex As Integer
'    Dim lngResult As Long
'
'    Const A = 0
'    Const B = 1
'
'    'convert to RGB values
'    udtStart(A) = GetRGB(lngStartFromA)
'    udtStart(B) = GetRGB(lngStartFromB)
'    udtFinish(A) = GetRGB(lngFinishToA)
'    udtFinish(B) = GetRGB(lngFinishToB)
'
'    'convert to pixels if necessary
'    If udtMesurement = InTwips Then
'        intDestLeft = intDestLeft / Screen.TwipsPerPixelX
'        intDestTop = intDestTop / Screen.TwipsPerPixelY
'        intDestWidth = intDestWidth / Screen.TwipsPerPixelX
'        intDestHeight = intDestHeight / Screen.TwipsPerPixelY
'        intLeft = intLeft / Screen.TwipsPerPixelX
'        intTop = intTop / Screen.TwipsPerPixelY
'        intWidth = intWidth / Screen.TwipsPerPixelX
'        intHeight = intHeight / Screen.TwipsPerPixelY
'    End If
'
'
'    'Find the largest difference between any two corresponding
'    'udtColours, and use that as the number of steps to take in the loop,
'    '(therefore guarenteing that it will cycle through all necessary
'    'udtColours without jumping)
'    For intIndex = A To B
'        'test red
'        intValue = Abs(udtStart(intIndex).Red - udtFinish(intIndex).Red)
'        If intValue > intBiggestDiff Then
'            intBiggestDiff = intValue
'        End If
'
'        'test green
'        intValue = Abs(udtStart(intIndex).Green - udtFinish(intIndex).Green)
'        If intValue > intBiggestDiff Then
'            intBiggestDiff = intValue
'        End If
'
'        'test blue
'        intValue = Abs(udtStart(intIndex).Blue - udtFinish(intIndex).Blue)
'        If intValue > intBiggestDiff Then
'            intBiggestDiff = intValue
'        End If
'    Next intIndex
'
'    'if there is no difference, then just draw one
'    'gradient and exit
'    If intBiggestDiff = 0 Then
'        'Call Gradient(lngGradhDc, _
'                      lngStartFromA, _
'                      lngStartFromB, _
'                      intLeft, _
'                      intTop, _
'                      intWidth, _
'                      intHeight, _
'                      udtDirection, _
'                      InPixels, _
'                      bytLineWidth)
'        Exit Sub
'    End If
'
'    'fade the gradient
'    For intCounter = 0 To intBiggestDiff
'        'find the point between udtColour udtStart and udtColour udtFinish that
'        'corresponds to the point between 0 and intBiggestDiff
'
'        For intIndex = A To B
'            udtColour(intIndex).Red = udtStart(intIndex).Red + (((udtFinish(intIndex).Red - udtStart(intIndex).Red) / intBiggestDiff) * intCounter)
'            udtColour(intIndex).Green = udtStart(intIndex).Green + (((udtFinish(intIndex).Green - udtStart(intIndex).Green) / intBiggestDiff) * intCounter)
'            udtColour(intIndex).Blue = udtStart(intIndex).Blue + (((udtFinish(intIndex).Blue - udtStart(intIndex).Blue) / intBiggestDiff) * intCounter)
'
'            'convert to long intValue and store
'            lngGradCol(intIndex) = RGB(udtColour(intIndex).Red, _
'                                       udtColour(intIndex).Green, _
'                                       udtColour(intIndex).Blue)
'        Next intIndex
'
'        'draw the gradient onto the bitmap
'        'Call Gradient(lngGradhDc, _
'                      lngGradCol(A), _
'                      lngGradCol(B), _
'                      intLeft, _
'                      intTop, _
'                      intWidth, _
'                      intHeight, _
'                      udtDirection, _
'                      InPixels, _
'                      bytLineWidth)
'
'        'blitt the bitmap to the screen
'        lngResult = BitBlt(mudtBitmap.lngDC, _
'                           intDestLeft, _
'                           intDestTop, _
'                           intDestWidth, _
'                           intDestHeight, _
'                           lngGradhDc, _
'                           intLeft, _
'                           intTop, _
'                           P_COPY)
'        DoEvents
'    Next intCounter
'End Sub

Public Function FromRGB(ByVal bytRed As Byte, _
                        ByVal bytGreen As Byte, _
                        ByVal bytBlue As Byte) _
                        As Long
    'Convert RGB to Long
     
    Dim lngMyVal As Long
    
    lngMyVal = (CLng(bytBlue) * 65536) + (CLng(bytGreen) * 256) + bytRed
    FromRGB = lngMyVal
End Function

Public Sub GetRGB(ByVal lngColour As Long, _
                  Optional ByRef intRed As Integer, _
                  Optional ByRef intGreen As Integer, _
                  Optional ByRef intBlue As Integer)

    'Convert Long to RGB:
    
                                                '        Sys Cols,  Blue  , Green  ,  Red
    Const RED_MASK      As Long = 255           'Binary: 00000000,00000000,00000000,11111111
    Const GREEN_MASK    As Long = 65280         'Binary: 00000000,00000000,11111111,00000000
    Const BLUE_MASK     As Long = 16711680      'Binary: 00000000,11111111,00000000,00000000
    Const SYS_MASK      As Double = 2147483648# 'Binary: 01000000,00000000,00000000,00000000
    Const MAX_COLOUR    As Long = 16777215      'Binary: 00000000,11111111,11111111,11111111
    
    If ((lngColour + SYS_MASK) <= 28) Then
        'convert a vb tagged system colour into a valid
        'RGB colour value
        
        'convert to index and get the colour
        lngColour = lngColour + SYS_MASK
        lngColour = GetSysColor(lngColour)
    Else
    
        'if the colour value is greater than acceptable
        'then exit
        If (lngColour > MAX_COLOUR) Or _
           (lngColour < -MAX_COLOUR) Then
            'system colour range
            Exit Sub
        End If
    End If
    
    'get the red, green and blue values
    intRed = (lngColour And RED_MASK)
    intGreen = (lngColour And GREEN_MASK) \ 256 'shift 8 bits to the right
    intBlue = (lngColour And BLUE_MASK) \ 65536 'shift 16 bits to the right
End Sub

Public Sub Pause(ByVal lngTicks As Long, _
                 Optional ByVal blnSleep As Boolean = False)
    'pause execution of the program for a specified
    'number of ticks, or stop execution of the thread
    'entirly for the period of the thread
    
    Dim lngStart As Long    'marks the starting tick
    
    'should we pause the thread of allow events to
    'operate as normal
    If blnSleep Then
        'don't allow events
        If lngTicks < 0 Then
            lngTicks = 0
        End If
        Call Sleep(lngTicks)
    Else
        'allow events
        lngStart = GetTickCount
        Do While GetTickCount < (lngStart + lngTicks)
            DoEvents
        Loop
    End If
End Sub

Public Sub DrawEllipse(ByVal intCenterX As Integer, _
                       ByVal intCenterY As Integer, _
                       ByVal intHeight As Integer, _
                       ByVal intWidth As Integer, _
                       Optional ByVal intTileAngle As Integer = 90, _
                       Optional ByVal lngColour As Long = 0, _
                       Optional ByVal intThickness As Integer = 1, _
                       Optional ByVal blnIsHollow As Boolean = True, _
                       Optional ByVal udtMesurement As Scaling = InPixels)
                       
    'This procedure will draw an ellipse of the specified dimensions and
    'lngColour, by drawing a line between each of the 360 points that make
    'up the ellipse.
    
    Const A = 1
    Const B = 2
    
    Dim sngMoveCenterX  As Single   'holds the co-ordinates of the centre point of the moving circle
    Dim sngMoveCenterY  As Single   'holds the co-ordinates of the centre point of the moving circle
    Dim sngCircleX      As Single   'holds the outer co-ordinates of the ellipse
    Dim sngCircleY      As Single   'holds the outer co-ordinates of the ellipse
    Dim sngCounter      As Single   'cycles through each dot on the ellipse
    Dim sngTiltX        As Single   'holds the co-ordinates of the tilted circle
    Dim sngTiltY        As Single   'holds the co-ordinates of the tilted circle
    Dim sngNumOfPoints  As Single   'the number of points used to draw the ellipse
    Dim udtEllipse()    As PointAPI 'holds a list of points on the ellipse
    Dim udtBrushStuff   As LogBrush 'holds the Brush information
    Dim lnghBrush       As Long     'holds a pointer to the brush
    Dim udtPenStuff     As LogPen   'holds the Pen information
    Dim lnghPen         As Long     'holds a pointer to the pen
    Dim lngResult       As Long     'holds any returned error value from an api call
    Dim lngJunk         As Long     'holds junk information from an api call
    Dim sngDegPerPoint  As Single   'holds the number of degrees to move to a new point on the ellipse
    
    'set scaling values
    If udtMesurement = InTwips Then
        'convert parameters to pixels
        intCenterX = (intCenterX / Screen.TwipsPerPixelX) '- intThickness
        intCenterY = (intCenterY / Screen.TwipsPerPixelY) '- intThickness
        intHeight = (intHeight / Screen.TwipsPerPixelY) - (intThickness * 2)
        intWidth = (intWidth / Screen.TwipsPerPixelX) - (intThickness * 2)
        
        'values are now in pixels
        udtMesurement = InPixels
    End If
    
    'calculate the radius for intWidth and intHeight
    intHeight = intHeight / 2
    intWidth = (intWidth / 2) - intHeight
    
    'calculate the starting point of the ellipse
    sngTiltX = Sin(intTileAngle * PI / 180) * intWidth
    sngTiltY = Cos(intTileAngle * PI / 180) * intWidth
    
    'draw the ellipse using one line for every three pixels
    'This will increase drawing speed on small ellipses and produce
    'detailed ones for large ellipses.
    sngNumOfPoints = (PI * (intWidth + intHeight)) / 3  '2.Pi.r = circumfrence, /3=per 3 pixels
    
    
    'size the array to match the number of points to be calculated
    ReDim udtEllipse(sngNumOfPoints)
    
    'calculate the number of degrees between points
    sngDegPerPoint = (360 / sngNumOfPoints)
    
    For sngCounter = 0 To 360 Step sngDegPerPoint
        sngMoveCenterX = intCenterX + (Cos(sngCounter * PI / 180) * sngTiltX)
        sngMoveCenterY = intCenterY + (Cos(sngCounter * PI / 180) * sngTiltY)
        
        'calculate the new position
        sngCircleX = Sin((sngCounter + intTileAngle) * PI / 180) * intHeight
        sngCircleY = Cos((sngCounter + intTileAngle) * PI / 180) * intHeight
        
        'add the points
        udtEllipse(sngCounter / sngDegPerPoint).X = sngMoveCenterX + sngCircleX
        udtEllipse(sngCounter / sngDegPerPoint).Y = sngMoveCenterY + sngCircleY
    Next sngCounter
    
    'draw the ellipse as a polygon
    
    'first create a brush and pen to display the colours
    udtBrushStuff.lbColor = lngColour
    If blnIsHollow Then
        udtBrushStuff.lbHatch = BS_HOLLOW
    Else
        udtBrushStuff.lbHatch = BS_SOLID
    End If
    udtBrushStuff.lbStyle = 0
    lnghBrush = CreateBrushIndirect(udtBrushStuff)
    
    'apply brush
    lngResult = SelectObject(mudtBitmap.lngDC, lnghBrush)
    
    'create the pen
    udtPenStuff.lopnColor = lngColour
    udtPenStuff.lopnWidth.X = intThickness
    udtPenStuff.lopnStyle = PS_SOLID And PS_INSIDEFRAME
    lnghPen = CreatePenIndirect(udtPenStuff)
    
    'apply pen
    lngResult = SelectObject(mudtBitmap.lngDC, lnghPen)
    
    'now draw the ellipse onto the mudtbitmap.hdcmemory
    lngResult = Polygon(mudtBitmap.lngDC, udtEllipse(0), sngNumOfPoints)
    
    'delete the brush and pen objects from memory
    lngJunk = SelectObject(mudtBitmap.lngDC, lnghBrush)
    lngJunk = DeleteObject(lngJunk)
    lngJunk = SelectObject(mudtBitmap.lngDC, lnghPen)
    lngJunk = DeleteObject(lnghPen)
End Sub

Public Function GetAngle(ByVal sngX1 As Single, _
                         ByVal sngY1 As Single, _
                         ByVal sngX2 As Single, _
                         ByVal sngY2 As Single) _
                         As Integer
                         
    'returns the angle of point1 in relation to point2
    
    Dim sngTempAngle As Single
    
    'if the values are not over the center point, then calculate the angle
    If ((sngY1 - sngY2) <> 0) Or ((sngX1 - sngX2) <> 0) Then
        
        sngTempAngle = (Atn(Slope(sngX1, sngY1, sngX2, sngY2)) * 180 / PI) Mod 360
        If sngTempAngle > 0 Then
            sngTempAngle = 90 - sngTempAngle
        Else
            sngTempAngle = Abs(sngTempAngle) + 90
        End If
        If sngX1 < sngX2 Then
            sngTempAngle = sngTempAngle + 180
        End If
        
        GetAngle = sngTempAngle
    End If
End Function

Public Function Slope(ByVal intX1 As Integer, _
                      ByVal intY1 As Integer, _
                      ByVal intX2 As Integer, _
                      ByVal intY2 As Integer) _
                      As Single
                      
    'This function finds the slope of a line, where the slope, m =
    '       intX1 - inty1
    'm = ------------
    '       intX2 - intY2
    
    Dim intXVal As Integer
    Dim intYVal As Integer
    
    intXVal = intX2 - intX1
    intYVal = intY2 - intY1
    If (intXVal = 0) And (intYVal = 0) Then
        'if both values were zero, then
        Slope = 0
        Exit Function
    Else
        'if only one value was zero then
        If (intXVal = 0) Or (intYVal = 0) Then
            'the slope = the other value
            Select Case 0
            Case intXVal
                Slope = intXVal
            Case intYVal
                Slope = intYVal
            End Select
            Exit Function
        End If
    End If
    
    If (intXVal <> 0) And (intYVal <> 0) Then
        'otherwise the slope = the formula
        Slope = (intY2 - intY1) / (intX2 - intX1)
    End If
End Function

Public Sub DrawString(ByVal strText As String, _
                      ByVal intTop As Integer, _
                      ByVal intLeft As Integer, _
                      ByVal intHeight As Integer, _
                      ByVal intWidth As Integer, _
                      ByRef fntFormat As Font, _
                      Optional ByVal lngColour As Long = vbBlack, _
                      Optional ByVal enmAlignment As AlignText = vbCentreAlign, _
                      Optional ByVal udtMeasurement As Scaling = InPixels)
                    
    'This procedure will draw strText onto the bitmap in the specified udtFont,
    'colour and position.
    
    Dim udtAPIFont      As LogFont      'contains the font information
    Dim lngAlignment    As Long         'holds the text alignment
    Dim udtTextRect     As Rect         'holds the area to draw the text in
    Dim lngResult       As Long         'holds any information returned from the api calls
    Dim lngJunk         As Long         'used to temperorily hold useless information
    Dim hDcFont         As Long         'holds a handle to the font
    Dim hDcOldFont      As Long         'holds a handle to the old font
    Dim intCounter      As Integer      'used to convert a string into a character array
    Dim udtFont         As FontStruc    'holds some font information
    
    'set Measurement values
    udtTextRect.Top = intTop
    udtTextRect.Left = intLeft
    udtTextRect.Right = intLeft + intWidth
    udtTextRect.Bottom = intTop + intHeight
    
    If udtMeasurement = InTwips Then
        'convert to pixels
        Call RectToPixels(udtTextRect)
    End If
    
    'put information into the definted font structure
    With udtFont
        .Alignment = enmAlignment
        .Name = fntFormat.Name
        .Bold = fntFormat.Bold
        .Italic = fntFormat.Italic
        .Underline = fntFormat.Underline
        .StrikeThru = fntFormat.Strikethrough
        .PointSize = fntFormat.Size
        .Colour = lngColour
    End With
    
    'Create details about the udtFont using the udtFont structure
    '====================
    
    'convert point size to pixels
    udtAPIFont.lfHeight = -((udtFont.PointSize * GetDeviceCaps(mudtBitmap.lngDC, LOGPIXELSY)) / 72)
    udtAPIFont.lfCharSet = DEFAULT_CHARSET
    udtAPIFont.lfClipPrecision = CLIP_DEFAULT_PRECIS
    udtAPIFont.lfEscapement = 0
    
    'move the name of the udtFont into the array
    For intCounter = 1 To Len(udtFont.Name)
        udtAPIFont.lfFaceName(intCounter) = Asc(Mid(udtFont.Name, intCounter, 1))
    Next intCounter
    'this has to be a Null terminated string
    udtAPIFont.lfFaceName(intCounter) = 0
    
    udtAPIFont.lfItalic = udtFont.Italic
    udtAPIFont.lfUnderline = udtFont.Underline
    udtAPIFont.lfStrikeOut = udtFont.StrikeThru
    udtAPIFont.lfOrientation = 0
    udtAPIFont.lfOutPrecision = OUT_DEFAULT_PRECIS
    udtAPIFont.lfPitchAndFamily = DEFAULT_PITCH
    udtAPIFont.lfQuality = PROOF_QUALITY
    
    If udtFont.Bold Then
        udtAPIFont.lfWeight = FW_BOLD
    Else
        udtAPIFont.lfWeight = FW_NORMAL
    End If
    
    udtAPIFont.lfWidth = 0
    hDcFont = CreateFontIndirect(udtAPIFont)
    hDcOldFont = SelectObject(mudtBitmap.lngDC, hDcFont)
    '====================
    
    Select Case udtFont.Alignment
    Case vbLeftAlign
        lngAlignment = DT_LEFT Or DT_VCENTER
    Case vbCentreAlign
        lngAlignment = DT_CENTER Or DT_VCENTER
    Case vbRightAlign
        lngAlignment = DT_RIGHT Or DT_VCENTER
    End Select
    
    'Draw the strText into the off-screen bitmap before copying the
    'new bitmap (with the strText) onto the screen.
    lngResult = SetBkMode(mudtBitmap.lngDC, TRANSPARENT)
    lngResult = SetTextColor(mudtBitmap.lngDC, udtFont.Colour)
    lngResult = DrawText(mudtBitmap.lngDC, _
                         strText, _
                         Len(strText), _
                         udtTextRect, _
                         lngAlignment)
    
    'clean up by deleting the off-screen bitmap and udtFont
    lngJunk = SelectObject(mudtBitmap.lngDC, hDcOldFont)
    lngJunk = DeleteObject(lngJunk)
    lngJunk = SelectObject(mudtBitmap.lngDC, hDcFont)
    lngJunk = DeleteObject(lngJunk)
End Sub

Public Function GetTextHeight() As Integer
    'This function will return the height of the text using the point size
    
    Dim udtMetrics  As TEXTMETRIC   'holds the font metrics for the bitmap
    Dim lngResult   As Long         'holds any returned error value from an api call
    
    lngResult = GetTextMetrics(mudtBitmap.lngDC, _
                               udtMetrics)
    
    GetTextHeight = udtMetrics.tmHeight
End Function

Public Sub GetScreenRes(ByRef intWidth As Integer, _
                        ByRef intHeight As Integer)
    'This procedure sets the variable to the current screen dimensions.
    
    intWidth = Screen.Width / Screen.TwipsPerPixelX
    intHeight = Screen.Height / Screen.TwipsPerPixelY
End Sub

Public Sub ReturnOldDisplay()
    'returns the display to what it was originally
    
    If mblnResChanged Then
        Call ChangeDisplaySettings(Null, 0)
        mblnResChanged = False
    End If
End Sub

Public Sub SetDisplay(ByVal intWidth As Integer, _
                      ByVal intHeight As Integer)
    'changes the resolution of the screen to new size
    
    Dim udtDevM     As DEVMODE  'holds the current device settings for the system
    Dim lngResult   As Long     'holds any returned error value from an api call
    
    lngResult = EnumDisplaySettings(0, 0, udtDevM)
    
    With udtDevM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = intWidth  'ScreenWidth
            .dmPelsHeight = intHeight 'ScreenHeight
            .dmBitsPerPel = 32 '(could be 8, 16, 32 or even 4)
    End With
    lngResult = ChangeDisplaySettings(udtDevM, 2)
    
    If lngResult = 0 Then
            Call ChangeDisplaySettings(udtDevM, 4)
    End If
    
    mblnResChanged = True
End Sub

Private Function IsWinNT() As Boolean
    'Detect if the program is running under Windows NT. Different
    'graphics api calls may be necessary under NT based systems
    
    Const VER_PLATFORM_WIN32_NT     As Integer = 2
    
    Dim myOS                        As OSVERSIONINFO
    Dim lngResult                   As Long             'returned error value from the api call
    
    'get version information
    myOS.dwOSVersionInfoSize = Len(myOS)
    lngResult = GetVersionEx(myOS)
    
    'return True if the test of windows NT is positive
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Public Sub LoadBitmap(ByVal strFileName As String, _
                      Optional ByVal enmImageType As LoadType = IMAGE_BITMAP, _
                      Optional ByVal enmMeasurement As Scaling = InPixels)
    'This procedure will load the specified image into a bitmap.
    'Please keep in mind that this function creates a bitmap object. If the
    'handles passed to this procedure already contain a bitmap, then it is
    'automatically deleted. This is assuming that the load operation was
    'successful. If not, then the pointers are left untouched.
    '
    'NOTE: is not currently supported by WinNT based systems
    
    Dim hBitmap         As Long     'holds the pointer to the bitmap itself
    Dim hMemory         As Long     'holds the pointer to the DC
    Dim hPointer        As Long     'holds a pointer to the previously assigned object
    Dim udtLoadArea     As Rect     'holds the size of the bitmap
    Dim lngResult       As Long     'holds any returned error value of api calls
    Dim udtDimensions   As BITMAP   'holds the bitmap dimensions and settings
    
    'delete the bitmap if it already exists
    
    'make sure we load the correct image type
    Select Case LCase(Right(strFileName, 3))
    Case "bmp"  'file is a bitmap
        enmImageType = IMAGE_BITMAP
    
    Case "ico"  'file is an icon
        enmImageType = IMAGE_ICON
    
    Case "cur"  'file is a cursor
        enmImageType = IMAGE_CURSOR
    End Select
    
    'we have to pass a Null terminated string
    strFileName = strFileName & Chr(0)
    hBitmap = LoadImage(0, _
                        strFileName, _
                        enmImageType, _
                        0, _
                        0, _
                        LR_LOADFROMFILE)
    
    'if the load operation was unsuccessful then exit the procedure.
    If hBitmap = Null Then
        'load unsuccessful
        Exit Sub
    End If
    
    'create a device context and assign the bitmap to it
    hMemory = CreateCompatibleDC(mudtBitmap.lngDC)
    hPointer = SelectObject(hMemory, hBitmap)
    
    'get the size of the bitmap loaded
    lngResult = GetObjectAPI(hBitmap, Len(udtDimensions), udtDimensions)
    
    'assign the udtDimensions to a rect structure. If no values
    'were returned, then use the classes dimensions
    udtLoadArea.Left = 0
    udtLoadArea.Top = 0
    If lngResult <> 0 Then
        'everything ok, get the size of the loaded bitmap
        udtLoadArea.Right = udtDimensions.bmWidth
        udtLoadArea.Bottom = udtDimensions.bmHeight
    Else
        'default size
        udtLoadArea.Right = 600
        udtLoadArea.Bottom = 480
    End If
    
    If enmMeasurement = InTwips Then
        'convert pixel values to twips
        Call RectToTwips(udtLoadArea)
    End If
    
    With mudtBitmap
        'Now that a new bitmap has been created, then we have to check to
        'see if the passed parameters are already in use hlding a bitmap. If so
        'then we need to delete it from memory before replacing it.
        Call DeleteBitmap(.lngDC, _
                          .lngBitmap, _
                          .lngPointer)
        
        'return the handles of the bitmap created
        .lngDC = hMemory
        .lngBitmap = hBitmap
        .lngPointer = hPointer
        .Area = udtLoadArea
    End With
End Sub

Public Sub ReSize(Optional ByVal intWidth As Integer = -1, _
                  Optional ByVal intHeight As Integer = -1, _
                  Optional ByVal blnStretch As Boolean = False)
    'This will change the size of the bitmap to the specified dimensions.
    'The bitmap can be stretched into it's new size. If any negative value
    'is passed, then the bitmaps original dimensions are used.
    
    Dim udtTempBitmap   As BitmapStruc  'holds a temperory bitmap
    Dim intNewWidth     As Integer      'holds the width to change to
    Dim intNewHeight    As Integer      'holds the height to change to
    Dim lngResult       As Long         'holds any returned error value
    
    'get the new dimensions
    With mudtBitmap.Area
        If intWidth < 0 Then
            'use original width
            intWidth = .Right - .Left
        End If
        If intHeight < 0 Then
            'use original height
            intHeight = .Bottom - .Top
        End If
    End With
    
    'create a bitmap of the size
    With udtTempBitmap
        .Area.Right = intWidth
        .Area.Bottom = intHeight
        Call CreateNewBitmap(.lngDC, _
                             .lngBitmap, _
                             .lngPointer, _
                             .Area)
    End With
    
    'copy any existing picture from the old to the new
    With mudtBitmap
        If blnStretch Then
            'stretch the old picture
            lngResult = StretchBlt(udtTempBitmap.lngDC, _
                                   0, _
                                   0, _
                                   intWidth, _
                                   intHeight, _
                                   .lngDC, _
                                   .Area.Left, _
                                   .Area.Top, _
                                   .Area.Right - .Area.Left, _
                                   .Area.Bottom - .Area.Top, _
                                   vbSrcCopy)
        Else
            'just copy the old picture. Any extra space will be black
            lngResult = BitBlt(udtTempBitmap.lngDC, _
                               0, _
                               0, _
                               intWidth, _
                               intHeight, _
                               .lngDC, _
                               .Area.Left, _
                               .Area.Top, _
                               vbSrcCopy)
        End If
        
        'adjust the dimensions so that they are in relation to the
        'old bitmaps Top and Left properties
        With .Area
            udtTempBitmap.Area.Left = .Left
            udtTempBitmap.Area.Top = .Top
            udtTempBitmap.Area.Right = .Left + intWidth
            udtTempBitmap.Area.Bottom = .Top + intHeight
        End With
        
        'destroy the old bitmap
        Call DeleteBitmap(.lngDC, .lngBitmap, .lngPointer)
    End With
    
    'change the pionters to the new bitmap
    mudtBitmap = udtTempBitmap
End Sub

Private Function RectIntersects(ByRef udtRect1 As Rect, _
                                ByRef udtRect2 As Rect) _
                                As Boolean
    'This function will return True if the two passed Rect structures
    'intersect each other.
    
    Dim udtTempRect As Rect 'holds the co-ordinates of the intersection rectangle
    Dim lngResult   As Long 'holds any returned error value from an api call
    
    lngResult = IntersectRect(udtTempRect, udtRect1, udtRect2)
    RectIntersects = lngResult
End Function

Public Function MouseKeyPressed(ByVal udtKey As MouseKeys) _
                                As Boolean
    'This will return True if the specified udtKey was pressed
    
    Const KeyPressed    As Integer = -32768
    
    Dim lngResult       As Long     'holds any returned error value from an api call
    Dim lngMyKey        As Long     'holds the code for the mouse key to check
    
    lngMyKey = udtKey
    
    'find the udtKey state of the mouse udtKey specified
    lngResult = GetAsyncKeyState(lngMyKey)
    
    'if the udtKey was pressed, then
    If lngResult = KeyPressed Then
        'the udtKey is pressed
        MouseKeyPressed = True
    End If
End Function

Public Sub GetScreenShot(Optional ByVal intScreenX As Integer = 0, _
                         Optional ByVal intScreenY As Integer = 0, _
                         Optional ByVal intDestWidth As Integer = -1, _
                         Optional ByVal intDestHeight As Integer = -1, _
                         Optional ByVal intDestX As Integer = 0, _
                         Optional ByVal intDestY As Integer = 0, _
                         Optional ByVal udtMeasurement As Scaling = InPixels)
                         
    'This will get a screen shot at the specified co-ordinates and copy
    'them into the specified destination co-ordinates.
    
    Dim lngResult As Long   'holds any returned error value from an api call
    
    'set the default if necessary
    If intDestWidth < 0 Then
        intDestWidth = mudtBitmap.Area.Left + _
                       mudtBitmap.Area.Right
    End If
    If intDestHeight < 0 Then
        intDestHeight = mudtBitmap.Area.Top + _
                        mudtBitmap.Area.Bottom
    End If
    
    'set the scaling mode specified and convert parameters if necessary
    If udtMeasurement = InTwips Then
        intDestX = intDestX / Screen.TwipsPerPixelX
        intDestY = intDestY / Screen.TwipsPerPixelY
        intDestWidth = intDestWidth / Screen.TwipsPerPixelX
        intDestHeight = intDestHeight / Screen.TwipsPerPixelY
        intScreenX = intScreenX / Screen.TwipsPerPixelX
        intScreenY = intScreenY / Screen.TwipsPerPixelY
    End If
    
    'copy the screen shot - GetDesktopWindow was previously
    'used to get the handle on the top window.
    lngResult = BitBlt(mudtBitmap.lngDC, _
                       intDestX, _
                       intDestY, _
                       intDestWidth, _
                       intDestHeight, _
                       GetDC(0), _
                       intScreenX, _
                       intScreenY, _
                       P_COPY)
End Sub

Public Sub SetBmpToMenu(ByVal hwnd As Long, _
                        ByVal lngTopPos As Long, _
                        ByVal lngSubPos1 As Long, _
                        Optional ByVal lngSubPos2 As Long = -1, _
                        Optional ByVal lngSubPos3 As Long = -1, _
                        Optional ByVal lngHDc As Long = 0)
    'This will set the graphic of a menu item
    'to any image in the device context. The menu
    'item must NOT be a top level menu or have a
    'sub menu. Top-level menu positions are from
    'left to right, starting at 0. Sub-level menu
    'positions are from top down, starting at 0.
    'NOTE : mudtBitmap.lngDC MUST be the Picture property
    'of a control
    
    Const MF_BYPOSITION As Long = &H400&
    Const MF_BYCOMMAND  As Long = &H0&
    Const BMP_SIZE      As Long = 14
    Const BF_BITMAP     As Long = &H4
    
    Dim lngResult       As Long     'any error message returned from an api call
    Dim hMenu           As Long     'the handle of the current menu item
    Dim hSubMenu        As Long     'the handle of the current sub menu
    Dim lngID           As Long     'the sub menu's ID
    
    'Get the handle of the form's menu
    hMenu = GetMenu(hwnd)
    
    'Get the handle of the form's submenu
    hSubMenu = GetSubMenu(hMenu, lngTopPos)
    
    'get any sub menu's
    If lngSubPos2 >= 0 Then
        hSubMenu = GetSubMenu(hSubMenu, lngSubPos2)
    End If
    If lngSubPos3 >= 0 Then
        hSubMenu = GetSubMenu(hSubMenu, lngSubPos3)
    End If
    
    'see if we are to use the bitmap graphic or if
    'another has been specified for us
    If lngHDc = 0 Then
        'use the object bitmap
        lngHDc = mudtBitmap.lngDC
    End If
    
    'if we were unable to get the sub menu handle
    'then exit
    If (hMenu = 0) _
       Or (hSubMenu = 0) _
       Or (lngHDc = 0) Then
        Exit Sub
    End If
    
    'set the graphic to the sub menu
    lngID = GetMenuItemID(hSubMenu, lngSubPos1)
    
    lngResult = SetMenuItemBitmaps(hSubMenu, _
                                   lngID, _
                                   MF_BYCOMMAND, _
                                   lngHDc, _
                                   lngHDc)
End Sub

Public Sub SetBitmap(ByVal intHeight As Integer, _
                     ByVal intWidth As Integer, _
                     Optional ByVal lngColour As Long = vbBlack)
    'this will recreate the bitmaps size and colour
    'to the specified dimensions
    
    'set the colour (makeing sure system colours are
    'converted to RGB values)
    mlngBackColour = GetSystemColour(lngColour)
    
    'create the bitmap
    With mudtBitmap
        'delete the old bitmap if it existed
        Call DeleteBitmap(.lngDC, _
                          .lngBitmap, _
                          .lngPointer)
        
        'create the bitmap size (0x0 by default)
        With .Area
            .Bottom = intHeight
            .Right = intWidth
        End With
        
        'create the bitmap compatable with the
        'Desktop device context (usually displayable
        'on a visual basic form or picture box)
        Call CreateNewBitmap(.lngDC, _
                             .lngBitmap, _
                             .lngPointer, _
                             .Area)
    End With
End Sub

Public Sub Paint(ByVal lngDestDc As Long, _
                 Optional ByVal intDestX As Integer = 0, _
                 Optional ByVal Y As Integer = 0, _
                 Optional ByVal intHeight As Integer = -1, _
                 Optional ByVal intWidth As Integer = -1, _
                 Optional ByVal intSourceX As Integer = 0, _
                 Optional ByVal intSourceY As Integer = 0, _
                 Optional ByVal lngPaintMode As PaintMode = P_COPY)
    'This will copy the bitmap onto the specified
    'device context and position. If the height or
    'width values are -1, then this is taken as the
    'height and width of the bitmap
    
    Dim lngResult As Long   'holds any error value returned from the api calls
    
    'if the height or width values are default, then
    'use the bitmaps size as the amount to copy
    If intHeight < 0 Then
        intHeight = mudtBitmap.Area.Bottom
    End If
    If intWidth < 0 Then
        intWidth = mudtBitmap.Area.Right
    End If
    
    'copy the bitmap onto the device context
    lngResult = BitBlt(lngDestDc, _
                       intDestX, _
                       Y, _
                       intWidth, _
                       intHeight, _
                       mudtBitmap.lngDC, _
                       intSourceX, _
                       intSourceY, _
                       lngPaintMode)
End Sub

Public Sub Cls()
    'clears the bitmap by drawing a rectangle covering
    'the entire bitmap with the background colour
    Call DrawRect
End Sub

Public Sub PaintFrom(ByVal lngSourceDC As Long, _
                     Optional ByVal intSrcWidth As Integer = -1, _
                     Optional ByVal intSrcHeight As Integer = -1, _
                     Optional ByVal intSrcLeft As Integer = 0, _
                     Optional ByVal intSrcTop As Integer = 0, _
                     Optional ByVal intDestLeft As Integer = 0, _
                     Optional ByVal intDestTop As Integer = 0, _
                     Optional ByVal intDestWidth As Integer = -1, _
                     Optional ByVal intDestHeight As Integer = -1, _
                     Optional ByVal enmPaintMode As PaintMode = P_COPY)
    'This will copy from a source device context onto
    'the bitmap at the appropiate position
    
    Dim lngResult As Long   'holds any returned error value from the api calls
    
    'validate the source device context
    If lngSourceDC = 0 Then
        Exit Sub
    End If
    
    'validate the measurements
    Select Case 0
    Case intSrcWidth, intSrcHeight, _
         intDestWidth, intDestHeight
        'nothing will be displayed, exit
        Exit Sub
    End Select
    
    'if the destination measurements are set to the
    'default values, then assume the destination
    'co-rodinates are to be the size of the bitmap
    If (intDestWidth < 0) Or (intDestHeight < 0) Then
        'set the default sizes as the size of the
        'bitmap
        intDestWidth = mudtBitmap.Area.Right - mudtBitmap.Area.Left
        intDestHeight = mudtBitmap.Area.Bottom - mudtBitmap.Area.Top
    End If
    
    If (intSrcWidth < 0) Or (intSrcHeight < 0) Then
        'set the default sizes as the size of the
        'bitmap
        intSrcWidth = mudtBitmap.Area.Right - mudtBitmap.Area.Left
        intSrcHeight = mudtBitmap.Area.Bottom - mudtBitmap.Area.Top
    End If
    
    'copy the bitmap using fastest api method possible
    If (intSrcWidth = intDestWidth) And _
       (intSrcHeight = intDestHeight) Then
    
        'use BitBlt to copy
        lngResult = BitBlt(mudtBitmap.lngDC, _
                           intDestLeft, _
                           intDestTop, _
                           intSrcWidth, _
                           intSrcHeight, _
                           lngSourceDC, _
                           intSrcLeft, _
                           intSrcTop, _
                           enmPaintMode)
    Else
        'use StretchBlt to copy
        lngResult = StretchBlt(mudtBitmap.lngDC, _
                               intDestLeft, _
                               intDestTop, _
                               intDestWidth, _
                               intDestHeight, _
                               lngSourceDC, _
                               intSrcLeft, _
                               intSrcTop, _
                               intSrcWidth, _
                               intSrcHeight, _
                               enmPaintMode)
    End If
End Sub

Public Function SupportedFormats() As String
    'Returns a string of all the file formats supported
    'for loading/saving. The string returned is intended
    'for direct use with the common dialog control
    '- currently only bitmaps; 22/8/2002
    
    SupportedFormats = "*.bmp" '- not supported yet -->";*.ico;*.cur"
End Function

Public Function GetSystemColour(ByVal lngIndex As SystemColorConstants) _
                                As Long
    'Convert a vb tagged system colour into a valid
    'RGB colour value. Please note that there are two
    'index values not represented here but can be
    'retrieved; ActiveTitleBar2 and InactiveTitleBar2.
    'These values can be got by using the constants
    '; vbActiveTitleBar2 and vbInactiveTitleBar2. These
    'values are only valid on systems where the title
    'bar on a window can display a gradient from one
    'colour to another
    
    Const SYS_MASK As Double = 2147483648#   'Binary: 01000000,00000000,00000000,00000000
    
    If (lngIndex < 0) Then
        'the colour passed is a system colour index
        lngIndex = lngIndex + SYS_MASK
        
        'use that index to get the RGB Long colour
        GetSystemColour = GetSysColor(lngIndex)
    Else
        'the colour passed is an ordinary colour, just
        'return
        GetSystemColour = lngIndex
    End If
End Function
                    

'------------------------------------------------------
'class events
'------------------------------------------------------

Private Sub Class_Initialize()
    'the defaults are as small as possible to take
    'up less memory until absolutly needed
    Const DEFAULT_HEIGHT = 0
    Const DEFAULT_WIDTH = 0
    
    'set the default values
    Call SetBitmap(DEFAULT_HEIGHT, DEFAULT_WIDTH)
End Sub

Private Sub Class_Terminate()
    'remove the bitmap from memory before exiting
    'the class
    
    With mudtBitmap
        Call DeleteBitmap(.lngDC, _
                          .lngBitmap, _
                          .lngPointer)
    End With
End Sub

'------------------------------------------------------
'Properties
'------------------------------------------------------

Public Property Get BackgroundColour() As Long
    'returns the RGB colour value of the background
    BackgroundColour = mlngBackColour
End Property

Public Property Let BackgroundColour(lngNew As Long)
    'sets the background colour of the bitmap,
    'converting system colour indecaes to RGB values
    'if necessary
    mlngBackColour = GetSystemColour(lngNew)
End Property

Public Property Get hDc() As Long
    'a read-only property that returns a handle to
    'the bitmap in case the user needs to use the
    'bitmap with other functions or procedures not
    'included here or with other instances of this
    'object
    
    hDc = mudtBitmap.lngDC
End Property

Public Property Get Width() As Integer
    'returns the width of the bitmap
    Width = mudtBitmap.Area.Right - mudtBitmap.Area.Left
End Property

Public Property Let Width(ByVal intWidth As Integer)
    'sets the width of the bitmap
    With mudtBitmap.Area
        Call ReSize(intWidth)
    End With
End Property

Public Property Get Height() As Integer
    'returns the height of the bitmap
    Height = mudtBitmap.Area.Bottom - mudtBitmap.Area.Top
End Property

Public Property Let Height(ByVal intHeight As Integer)
    'sets the height of the bitmap
    With mudtBitmap.Area
        Call ReSize(, intHeight)
    End With
End Property

Public Property Get Left() As Integer
    'returns the left position of the bitmap
    Left = mudtBitmap.Area.Left
End Property

Public Property Let Left(ByVal intLeft As Integer)
    'sets the left position of the bitmap
    With mudtBitmap.Area
        .Right = .Right + (.Left - intLeft)
        .Left = intLeft
    End With
End Property

Public Property Get Top() As Integer
    'returns the left position of the bitmap
    Top = mudtBitmap.Area.Top
End Property

Public Property Let Top(ByVal intTop As Integer)
    'sets the left position of the bitmap
    With mudtBitmap.Area
        .Bottom = .Bottom + (.Top - intTop)
        .Top = intTop
    End With
End Property
