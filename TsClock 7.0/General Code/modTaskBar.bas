Attribute VB_Name = "modTaskBar"
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     2 September 1999
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    Taskbar Information Module
' -----------------------------------------------
'COMMENTS :
'This module provides information about the
'screen's work area and the task-bar and monitors.
'=================================================

'all variables must be declared
Option Explicit

'------------------------------------------------
'               API DECLARATIONS
'------------------------------------------------

'get information about the specified monitor
Private Declare Function GetMonitorInfo _
        Lib "user32.dll" _
        Alias "GetMonitorInfoA" _
            (ByVal hMonitor As Long, _
             ByRef lpmi As MONITORINFO) _
             As Long

'find out which monitor the specified point is on
Private Declare Function MonitorFromPoint _
        Lib "user32.dll" _
            (ByVal X As Long, _
             ByVal Y As Long, _
             ByVal dwFlags As Long) _
             As Long

'get a handle to the monitor that has the largest
'section of the specified area
Private Declare Function MonitorFromRect _
        Lib "user32.dll" _
            (ByRef lprc As Rect, _
             ByVal dwFlags As Long) _
             As Long

'get a handle to the monitor that has the largest
'part of the specified window
Private Declare Function MonitorFromWindow _
        Lib "user32.dll" _
            (ByVal hwnd As Long, _
             ByVal dwFlags As Long) _
             As Long

'send information to a specified function, calling
'the function once for each monitor found (dangerous!)
Private Declare Function EnumDisplayMonitors _
        Lib "user32.dll" _
            (ByVal hdc As Long, _
             ByRef lprcClip As Any, _
             ByVal lpfnEnum As Long, _
             ByVal dwData As Long) _
             As Long

'get the bounding dimensions of the specified window
Private Declare Function GetWindowRect _
        Lib "user32" _
            (ByVal hwnd As Long, _
             lpRect As Rect) _
             As Long

'gets some system information
Public Declare Function SystemParametersInfo _
        Lib "user32" _
        Alias "SystemParametersInfoA" _
            (ByVal uAction As Long, _
             ByVal uParam As Long, _
             lpvParam As Any, _
             ByVal fuWinIni As Long) _
             As Long

'gets the position of the cursor
Private Declare Function GetCursorPos _
        Lib "user32" _
            (lpPoint As PointAPI) _
             As Long

'resize the specified window or change its Z order
Private Declare Sub SetWindowPos _
        Lib "user32" _
            (ByVal hwnd As Long, _
             ByVal hWndInsertAfter As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal cx As Long, _
             ByVal cy As Long, _
             ByVal wFlags As Long)

'------------------------------------------------
'               USER-DEFINED TYPES
'------------------------------------------------
Private Type PointAPI
    X           As Long
    Y           As Long
End Type

Private Type Rect
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Type MONITORINFO
    cbSize      As Long
    rcMonitor   As Rect
    rcWork      As Rect
    dwFlags     As Long
End Type

'------------------------------------------------
'                   ENUMERATORS
'------------------------------------------------
Public Enum AlignmentConst
    vbLeft = 0
    vbRight = 1
    vbTop = 2
    vbBottom = 3
End Enum

Public Enum EnumWindowZOrder
    HWND_NOTOPMOST = -2
    HWND_TOPMOST = -1
    HWND_TOP = 0
    HWND_BOTTOM = 1
End Enum

Public Enum EnumWindowMsgFlags
    SWP_DRAWFRAME = &H20
    SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
    SWP_HIDEWINDOW = &H80
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = &H200
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_SHOWWINDOW = &H40
End Enum

'------------------------------------------------
'               MODULE-LEVEL CONSTANTS
'------------------------------------------------
Private Const SPI_GETWORKAREA           As Long = 48
Private Const NIM_ADD                   As Long = 0
Private Const NIM_MODIFY                As Long = &H1
Private Const NIM_DELETE                As Long = &H2
Private Const NIF_MESSAGE               As Long = &H1
Private Const NIF_ICON                  As Long = &H2
Private Const NIF_TIP                   As Long = &H4
Private Const MONITORINFOF_PRIMARY      As Long = &H1
Private Const MONITOR_DEFAULTTONEAREST  As Long = &H2
Private Const MONITOR_DEFAULTTONULL     As Long = &H0
Private Const MONITOR_DEFAULTTOPRIMARY  As Long = &H1

'------------------------------------------------
'                 GLOBAL CONSTANTS
'------------------------------------------------
Public Const WM_LBUTTONDBLCLK           As Long = &H203
Public Const WM_LBUTTONDOWN             As Long = &H201
Public Const WM_LBUTTONUP               As Long = &H202
Public Const WM_RBUTTONDBLCLK           As Long = &H206
Public Const WM_RBUTTONDOWN             As Long = &H204
Public Const WM_RBUTTONUP               As Long = &H205
Public Const WM_MOUSEMOVE               As Long = &H200

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Public Sub MoveWindow(ByVal frmMove As Form, _
                      ByVal intX As Integer, _
                      ByVal intY As Integer, _
                      Optional ByVal intWidth As Long = -1, _
                      Optional ByVal intHeight As Long = -1, _
                      Optional lngZOrderFlags As EnumWindowZOrder = HWND_NOTOPMOST, _
                      Optional lngMsgFlags As EnumWindowMsgFlags = SWP_NOACTIVATE + SWP_NOSIZE)
    'This will update the position of the specified form. The new postion
    'has to be passed in pixels.
    
    'validate the parameters
    If (frmMove Is Nothing) Then
        Exit Sub
    End If
    
    With frmMove
        'see if new width/height parameters were passed
        If (intWidth < 0) Then
            intWidth = .Width
        End If
        If (intHeight < 0) Then
            intHeight = .Height
        End If
        
        Call SetWindowPos(.hwnd, _
                          lngZOrderFlags, _
                          intX, _
                          intY, _
                          intWidth, _
                          intHeight, _
                          lngMsgFlags)
    End With
End Sub


Public Function GetWorkArea() As Rect
    'Get the area the user is working with (minus the
    'task bar) in PIXELS
    
    Dim lngResult   As Long
    Dim udtWorkArea As Rect             'holds information about the current window
    
    lngResult = SystemParametersInfo(SPI_GETWORKAREA, _
                                     0, _
                                     udtWorkArea, _
                                     0)
    GetWorkArea = udtWorkArea
End Function

Public Function GetAlignment() As AlignmentConst
    'Find the alignment of the taskbar
    
    Dim udtWorkArea As Rect
    Dim enmAlign    As AlignmentConst
    
    'get the system work area
    udtWorkArea = GetWorkArea
    
    If udtWorkArea.Left <> 0 Then
        'the taskbar MUST be right aligned
        enmAlign = vbLeft
    Else
        If udtWorkArea.Top <> 0 Then
            'The taskbar MUST be bottom aligned
            enmAlign = vbTop
        Else
            If ((udtWorkArea.Bottom - udtWorkArea.Top) * Screen.TwipsPerPixelY) = Screen.Height Then
                'If the udtWorkArea height is equal to the screen height then
                'the taskbar MUST be right aligned
                enmAlign = vbRight
            Else
                enmAlign = vbBottom
            End If
        End If
    End If
    
    'return the alignment
    GetAlignment = enmAlign
End Function

Public Function TaskBarDimensions() As Rect
    'Find out what the taskbars', left, top, right and
    'bottom values are in TWIPS, not Pixels
    
    Dim udtWorkArea     As Rect
    Dim udtTaskBarDet   As Rect  'task bar details
    Dim bytTwipsPP(2)   As Byte  'Twips Per Pixel
    
    udtWorkArea = GetWorkArea
    bytTwipsPP(X) = Screen.TwipsPerPixelX
    bytTwipsPP(Y) = Screen.TwipsPerPixelY
    
    'set the taskbars' default values to the screen size
    udtTaskBarDet.Top = 0
    udtTaskBarDet.Bottom = Screen.Height
    udtTaskBarDet.Left = 0
    udtTaskBarDet.Right = Screen.Width
    
    'change the appropiate value according to alignment
    'of the taskbar
    Select Case GetAlignment
    Case vbLeft
        udtTaskBarDet.Right = (udtWorkArea.Left * bytTwipsPP(X))
    Case vbRight
        udtTaskBarDet.Left = (udtWorkArea.Right * bytTwipsPP(X))
    Case vbTop
        udtTaskBarDet.Bottom = (udtWorkArea.Top * bytTwipsPP(Y))
    Case vbBottom
        udtTaskBarDet.Top = (udtWorkArea.Bottom * bytTwipsPP(Y))
    End Select
    
    'return lngResult
    TaskBarDimensions = udtTaskBarDet
End Function

Public Sub SnapWindow(ByVal frmSnap As Form, _
                      ByVal intDistance As Integer, _
                      Optional blnCheckBounds As Boolean = True)
    'This procedure will snap the window to the edges
    'of the work area (like winamp) if the form is
    'within a certain intDistance of the edges (measured
    'in pixels).
    
    Dim udtWorkArea     As Rect     'holds the working screen area
    Dim lngDistTwip(1)  As Long     'the number of twips per pixel X/Y
    Dim lngMoveTo(1)    As Long     'the position to move the form to
    
    'if the form is minimized or maximized then don't do this
    '- it will generate an error otherwise.
    If (frmSnap.WindowState = vbMinimized) Or _
       (frmSnap.WindowState = vbMaximized) Then
        Exit Sub
    End If
    
    'make sure the window is not outside the screen
    'bounds
    If blnCheckBounds Then
        Call CheckIfOutsideScreen(frmSnap)
    End If
    
    If intDistance < 1 Then
        'a value of zero is meaningless to this
        'procedure and a value of less than zero is
        'invalid.
        Exit Sub
    End If
    
    'find out if the edge of the window is within
    'snapping distance
    udtWorkArea = GetWorkArea
    lngDistTwip(X) = Screen.TwipsPerPixelX
    lngDistTwip(Y) = Screen.TwipsPerPixelY
    
    With frmSnap
        'record the original position
        lngMoveTo(X) = .Left
        lngMoveTo(Y) = .Top
    
        'check top side
        If WithinDistance((.Top / lngDistTwip(Y)), _
                          udtWorkArea.Top, _
                          intDistance) Then
            'snap window to the top
            lngMoveTo(Y) = udtWorkArea.Top * lngDistTwip(Y)
        End If
        
        'check left side
        If WithinDistance((.Left / lngDistTwip(X)), _
                          udtWorkArea.Left, _
                          intDistance) Then
            'snap window to the left
            lngMoveTo(X) = udtWorkArea.Left * lngDistTwip(X)
        End If
        
        'check botton side
        If WithinDistance(((.Top + .Height) / lngDistTwip(Y)), _
                          udtWorkArea.Bottom, _
                          intDistance) Then
            'snap window to the bottom
            lngMoveTo(Y) = (udtWorkArea.Bottom * lngDistTwip(Y)) - .Height
        End If
        
        'check right side
        If WithinDistance(((.Left + .Width) / lngDistTwip(X)), _
                          udtWorkArea.Right, _
                          intDistance) Then
            'snap window to the right
            lngMoveTo(X) = (udtWorkArea.Right * lngDistTwip(X)) - .Width
        End If
        
        'move the form to its new position
        Call MoveWindow(frmSnap, _
                        lngMoveTo(X) \ lngDistTwip(X), _
                        lngMoveTo(Y) \ lngDistTwip(Y))
    End With
End Sub

Public Function WithinDistance(ByVal lngValue As Long, _
                               ByVal lngEdge As Long, _
                               ByVal lngDistance As Long) _
                               As Boolean
    'Find out if the lngValue is within the Distance of the Edge
    If (lngValue > (lngEdge - lngDistance)) And _
       (lngValue < (lngEdge + lngDistance)) Then
        WithinDistance = True
    Else
        WithinDistance = False
    End If
End Function

Public Sub CheckIfOutsideScreen(ByVal frmCheck As Form)
    'This moves a form inside the work area of the screen if the
    'form is outside the work area.
    
    Dim intLeft     As Integer  'position to move clock to
    Dim intTop      As Integer  'position to move clock to
    Dim intWidth    As Integer  'the width of the form
    Dim intHeight   As Integer  'the height of the form
    Dim udtWorkArea As Rect     'the current monitor screen area
    Dim intTwip(1)  As Integer  'the number of twips per pixel
    Dim blnMoved    As Boolean  'holds whether or not the form has changed position
    
    'if the form is minimized or maximized then don't do this
    '- it will generate an error otherwise.
    If (frmCheck.WindowState = vbMinimized) Or _
       (frmCheck.WindowState = vbMaximized) Then
        Exit Sub
    End If
    
    udtWorkArea = GetWorkArea
    
    'convert udtWorkArea to twips
    intTwip(X) = Screen.TwipsPerPixelX
    intTwip(Y) = Screen.TwipsPerPixelY
    
    With udtWorkArea
        .Top = .Top * intTwip(Y)
        .Bottom = .Bottom * intTwip(Y)
        .Left = .Left * intTwip(X)
        .Right = .Right * intTwip(X)
    End With
    
    'get form dimensions
    With frmCheck
        intLeft = .Left
        intTop = .Top
        intWidth = .Width
        intHeight = .Height
    End With
    
    'horizontal
    If (intLeft + intWidth) > udtWorkArea.Right Then
        intLeft = udtWorkArea.Right - intWidth
        blnMoved = True
    End If
    
    If intLeft < udtWorkArea.Left Then
        intLeft = udtWorkArea.Left
        blnMoved = True
    End If
    
    'vertical
    If (intTop + intHeight) > udtWorkArea.Bottom Then
        intTop = udtWorkArea.Bottom - intHeight
        blnMoved = True
    End If
    
    If intTop < udtWorkArea.Top Then
        intTop = udtWorkArea.Top
        blnMoved = True
    End If
    
    'set the noew position of the form
    If blnMoved Then
        'just reposition
        Call MoveWindow(frmCheck, _
                        intLeft \ Screen.TwipsPerPixelX, _
                        intTop \ Screen.TwipsPerPixelY)
    End If
End Sub

Public Sub CentreForm(ByVal frmName As Form)
    'This procedure will centre the form in the work
    'area. The form parameter
    '"StartUpPosition = CenterScreen" does not centre
    'the form in the work area, ie it will not take into
    'account the position/height of the taskbar when
    'positioning the form
    
    Dim udtWorkArea As Rect
    
    udtWorkArea = AreaToTwips(GetWorkArea)
    
    frmName.Left = ((udtWorkArea.Right - frmName.Width) / 2) - udtWorkArea.Left
    frmName.Top = ((udtWorkArea.Bottom - frmName.Height) / 2) - udtWorkArea.Top
End Sub

Public Function AreaToTwips(ByRef udtWorkArea As Rect) _
                            As Rect
    'This function will convert a rect structure to twips
    
    With udtWorkArea
        .Left = .Left * Screen.TwipsPerPixelX
        .Right = .Right * Screen.TwipsPerPixelX
        .Top = .Top * Screen.TwipsPerPixelY
        .Bottom = .Bottom * Screen.TwipsPerPixelY
    End With
    
    AreaToTwips = udtWorkArea
End Function

'Public Function MonitorEnumProc(ByVal hMonitor As Long, _
'                                ByVal hdcMonitor As Long, _
'                                lprcMonitor As RECT, _
'                                ByVal dwData As Long) _
'                                As Long
'    Dim MI As MONITORINFO, R As RECT
'
'    Debug.Print "Moitor handle: " + CStr(hMonitor)
'    'initialize the MONITORINFO structure
'    MI.cbSize = Len(MI)
'    'Get the monitor information of the specified monitor
'    GetMonitorInfo hMonitor, MI
'    'write some information on teh debug window
'    Debug.Print "Monitor Width/Height: " + CStr(MI.rcMonitor.Right - MI.rcMonitor.Left) + "x" + CStr(MI.rcMonitor.Bottom - MI.rcMonitor.Top)
'    Debug.Print "Primary monitor: " + CStr(CBool(MI.dwFlags = MONITORINFOF_PRIMARY))
'    'check whether Form1 is located on this monitor
'    If MonitorFromWindow(Form1.hwnd, MONITOR_DEFAULTTONEAREST) = hMonitor Then
'        Debug.Print "Form1 is located on this monitor"
'    End If
'    'heck whether the point (0, 0) lies within the bounds of this monitor
'    If MonitorFromPoint(0, 0, MONITOR_DEFAULTTONEAREST) = hMonitor Then
'        Debug.Print "The point (0, 0) lies wihthin the range of this monitor..."
'    End If
'    'check whether Form1 is located on this monitor
'    GetWindowRect Form1.hwnd, R
'    If MonitorFromRect(R, MONITOR_DEFAULTTONEAREST) = hMonitor Then
'        Debug.Print "The rectangle of Form1 lies within this monitor"
'    End If
'    Debug.Print ""
'    'Continue enumeration
'    MonitorEnumProc = 1
'End Function
