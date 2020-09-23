VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     2 September 1999
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    Mouse API Module
' -----------------------------------------------
'COMMENTS :
'This module is for general use with the mouse such
'as moving it, changing the cursor picture etc.
'==================================================

'all variables must be declared
Option Explicit

'------------------------------------------------
'               API DECLARATIONS
'------------------------------------------------
'this will get the window that has the current
'mouse capture
Private Declare Function GetCapture _
        Lib "user32" () As Long

'returns a handle to the current cursor
Private Declare Function GetCursor _
        Lib "user32" _
            () _
             As Long

'stores the cursors position in the PointAPI structure
'The values returned are in pixels
Private Declare Function GetCursorPos _
        Lib "user32" _
            (lpPoint As PointAPI) _
             As Long

'prevents the mouse from moving outside the specified
'co-ordinates
Private Declare Sub ClipCursor _
        Lib "user32" _
            (lpRect As Any)

'stop moving the window with the mouse. See SetCapture
'and GetCapture for more details
Private Declare Function ReleaseCapture _
        Lib "user32" () As Long

'this will tell windows that the specified window
'has captured the mouse. Moving the mouse will now
'move the window. Call ReleaseCapture to stop '
'moving the window.
Private Declare Function SetCapture _
        Lib "user32" _
            (ByVal hWnd As Long) _
             As Long

'set the cursor to the specified attributes.
'The function needs a handle to these attributes
Private Declare Function SetCursor _
        Lib "user32" _
            (ByVal hCursor As Long) _
             As Long

'set the position of the cursor on the screen using
'picel co-ordinates
Private Declare Function SetCursorPos _
        Lib "user32" _
            (ByVal X As Long, _
             ByVal Y As Long) _
             As Long

'show or hide the cursor
Private Declare Function ShowCursor _
        Lib "user32" _
            (ByVal bShow As Long) _
             As Long
             
'returns the amount of time windows has been active for
'in milliseconds (sec/1000)
Private Declare Function GetTickCount _
        Lib "kernel32" _
            () _
             As Long

'sends a message to the specified window
Private Declare Function SendMessage _
        Lib "user32" _
        Alias "SendMessageA" _
            (ByVal hWnd As Long, _
             ByVal wMsg As Long, _
             ByVal wParam As Long, _
             lParam As Any) _
             As Long

'pause execution of the current thread for the specifed
'number of milliseconds. Not recommended for use
Private Declare Sub Sleep _
        Lib "kernel32" _
            (ByVal dwMilliseconds As Long)

'returns the handle of a window from the specified point
Private Declare Function WindowFromPoint _
        Lib "user32" _
            (ByVal xPoint As Long, _
             ByVal yPoint As Long) _
             As Long

'------------------------------------------------
'               USER-DEFINED TYPES
'------------------------------------------------
Private Type PointAPI
    X           As Long
    Y           As Long
End Type

Private Type MosRect
    lngLeft     As Long
    lngTop      As Long
    lngRight    As Long
    lngBottom   As Long
End Type

'------------------------------------------------
'                GLOBAL VARIABLES
'------------------------------------------------
Private mblnMouseMoveCancel As Boolean      'used to cancel the mouse being moved (eg, if the form is unloaded, then you don't want to continue moving the mouse)

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Public Sub CancelMouseMove()
    'sets the cancle mouse move flag
    mblnMouseMoveCancel = True
End Sub

Public Sub Limit(Optional ByVal lngLeft As Long = -1, _
                 Optional ByVal lngTop As Long = -1, _
                 Optional ByVal lngWidth As Long = -1, _
                 Optional ByVal lngHeight As Long = -1, _
                 Optional ByVal blnInPixels As Boolean = True)
    'This will limit the mouse to the specified screen region. If any co-
    'ordinates are less than zero, then the screens measurement is used. If
    'the blnInPixels variable is false, then it is assumed that the measurements
    'used are in Twips.
    
    Dim udtBounds       As MosRect      'holds the bounding measurements
    Dim intTppX         As Integer      'holds the Twips Per Pixel for the X axis
    Dim intTppY         As Integer      'holds the Twips Per Pixel for the Y axis
    Dim lngResult       As Long         'holds any returned value from an api call
    
    intTppX = Screen.TwipsPerPixelX
    intTppY = Screen.TwipsPerPixelY
    
    'are the measurements in pixels
    If Not blnInPixels Then
        lngTop = lngTop / intTppY
        lngLeft = lngLeft / intTppX
        lngHeight = lngHeight / intTppY
        lngWidth = lngWidth / intTppX
    End If
    
    'check if we are to use the screens measurements or the ones specified
    If (lngTop < 0) Then
        lngTop = 0
    End If
    If (lngLeft < 0) Then
        lngLeft = 0
    End If
    If (lngWidth <= 0) Then
        lngWidth = Screen.Width / intTppX
    End If
    If (lngHeight <= 0) Then
        lngHeight = Screen.Height / intTppY
    End If
    
    'enter the specified bounds
    With udtBounds
        .lngTop = lngTop
        .lngLeft = lngLeft
        .lngRight = (lngLeft + lngWidth)
        .lngBottom = (lngTop + lngHeight)
    End With    'udtBounds
    
    Call ClipCursor(udtBounds)
End Sub

Public Sub MoveTo(ByVal intXCo As Integer, _
                  ByVal intYCo As Integer, _
                  Optional ByVal lngDelay As Long = 0, _
                  Optional ByVal intSteps As Integer = 0, _
                  Optional ByVal blnSleep As Boolean = True)
    'This procedure moves the mouse to a specified
    'point on the screen in a given number of steps
    
    Const X             As Integer = 0
    Const Y             As Integer = 1
    Const DEFAULT_DELAY As Integer = 1  'the number of ticks to wait before moving the mouse to the next position
    
    Dim udtGotPos       As PointAPI     'the current position of the mouse
    Dim intCounter      As Integer      'used to animate the mouse across the screen
    Dim intJump         As Integer      'the number of pixels to jump by
    Dim intDirection(1) As Integer      'the direction of the x/y directions
    Dim intDistance(1)  As Integer      'the distance to travel for the x/y directions
    Dim lngMove(1)      As Long         'where to move the mouse for the current step
    Dim lngResult       As Long         'holds any error value returned from the api call
    Dim lngStartTick    As Long         'holds the starting tick of the move
    Dim lngCurrTick     As Long         'holds the current tick before we pause
    Dim lngDelayTick    As Long         'holds how long has elapsed between ticks
    
    'set the cancel flag to false - this will allow the mouse
    'to be moved at least once.
    mblnMouseMoveCancel = False
    
    'set the frame rate
'    If lngDelay = 0 Then
'        lngDelay = DEFAULT_DELAY
'    End If
    
    'get the current mouse positions
    udtGotPos.X = Me.XPos
    udtGotPos.Y = Me.YPos
    
    'get the distance to travel in each axis
    intDistance(X) = intXCo - udtGotPos.X
    intDistance(Y) = intYCo - udtGotPos.Y
    
    'find the direction to move the mouse cursor in
    For intCounter = X To Y
        If intDistance(intCounter) < 0 Then
            'travel left, up
            intDirection(intCounter) = -1
        Else
            'travel right, down
            intDirection(intCounter) = 1
        End If
    Next intCounter
    
    'make sure the distance to travel is
    'always a positive value
    intDistance(X) = Abs(intDistance(X))
    intDistance(Y) = Abs(intDistance(Y))
    
    If intSteps <> 0 Then
        'the the number of intSteps was
        'specified then, use them
        intJump = intSteps
    Else
        'The number of steps = the largest
        'distance to lngMove (by pixel)
        If Abs(intDistance(X)) > Abs(intDistance(Y)) Then
            'the largest intDistance between
            'the two points is horizontal
            intJump = Abs(intDistance(X))
        Else
            'else the largest intDistance
            'is vertical or is equal
            intJump = Abs(intDistance(Y))
        End If
    End If
    
    For intCounter = 1 To intJump
        'Move the mouse
        
        lngStartTick = GetTickCount
        
        'if the cancel flag has been set, then exit
        If mblnMouseMoveCancel Then
            Exit Sub
        End If
        
        'Move horizontal
        lngMove(X) = udtGotPos.X + (((intDistance(X) / intJump) _
                                     * intCounter) _
                                    * intDirection(X))
        
        'Move vertical
        lngMove(Y) = udtGotPos.Y + (((intDistance(Y) / intJump) _
                                     * intCounter) _
                                    * intDirection(Y))
        
        'set the cursor position but don't allow the user to use it
        'while we move it
        Call Limit(lngMove(X), lngMove(Y), 1, 1)
        
        lngCurrTick = GetTickCount
        lngDelayTick = lngDelay - (lngCurrTick - lngStartTick)
        
        'pause before moving to the next position
        If (lngDelayTick > 0) Then
            Call Pause(lngDelayTick, blnSleep)
        End If
    Next intCounter
    
    'allow the user to move the mouse freely now that we have moved it
    Call Limit
End Sub

Private Sub SwapVal(ByRef intVal1 As Integer, _
                    ByRef intVal2 As Integer)
    'This procedure will swap the two values around
    
    Dim intTemp As Integer
    
    intTemp = intVal1
    intVal1 = intVal2
    intVal2 = intTemp
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

Public Sub GrabWindow(ByVal lnghWindow As Long)
    'grab the mouse to drag the specified form
    
    Const WM_NCLBUTTONDOWN  As Long = &HA1
    Const HTCAPTION         As Long = 2
    
    Dim lngResult   As Long     'holds any returned error value returned from an api call
    
    'make sure the specified window exists
    If (lnghWindow <> 0) Then
        lngResult = SendMessage(lnghWindow, _
                                WM_NCLBUTTONDOWN, _
                                HTCAPTION, _
                                0)
    End If
End Sub

Public Sub ReleaseWindow()
    'release the current window from being dragged
    
    Dim lngResult   As Long     'holds any returned error value returned from an api call
    
    lngResult = ReleaseCapture
End Sub

Public Function GetGrabbedWindow() As Long
    'returns a handle to the window currently grabbed
    'by the mouse
    
    GetGrabbedWindow = GetCapture
End Function

Public Sub SetPos(Optional ByVal intX As Integer = 0, _
                  Optional ByVal intY As Integer = 0)
    'sets the position of the mouse to the point specified
    'on the screen.
    
    Dim lngResult   As Long     'holds any returned error value from an api call
    
    lngResult = SetCursorPos(intX, intY)
End Sub

Public Sub Show()
    'This will show the mouse on the screen if not
    'already visible.
    
    Dim lngResult   As Long     'holds any returned error value from an api call
    
    lngResult = ShowCursor(True)
End Sub

Public Sub Hide()
    'This will hide the mouse on the screen. This
    'does not prevent the user from still using
    'the mouse.
    
    Dim lngResult   As Long     'holds any returned error value from an api call
    
    lngResult = ShowCursor(False)
End Sub

Public Function hWndFromPoint(Optional ByVal intX As Integer = -1, _
                              Optional ByVal intY As Integer = -1) _
                              As Long
    'This will return the handle from the specified mouse position. If no
    'mouse screen position is specified, the current mouse position is used
    
    Dim hWnd    As Long     'holds the handle of the window that we have found
    
    'set defaults
    If (intX < 0) Then
        intX = Me.XPos
    End If
    If (intY < 0) Then
        intY = Me.YPos
    End If
    
    'get the window handle
    hWnd = WindowFromPoint(Me.XPos, Me.YPos)
    
    'return the handle
    hWndFromPoint = hWnd
End Function


'------------------------------------------------
'                   PROPERTIES
'------------------------------------------------
Public Property Get XPos() As Integer
    'returns the current horizontal position of
    'the mouse in pixels. This property is read-
    'only. The position of the mouse can be set
    'using the SetPos method.
    
    Dim lngResult   As Long     'holds any returned error value from an api call
    Dim udtPosition As PointAPI 'holds the current mouse position on the screen
    
    lngResult = GetCursorPos(udtPosition)
    
    XPos = udtPosition.X
End Property

Public Property Get YPos() As Integer
    'returns the current vertical position of
    'the mouse in pixels. This property is read-
    'only. The position of the mouse can be set
    'using the SetPos method.
    
    Dim lngResult   As Long     'holds any returned error value from an api call
    Dim udtPosition As PointAPI 'holds the current mouse position on the screen
    
    lngResult = GetCursorPos(udtPosition)
    
    YPos = udtPosition.Y
End Property

'------------------------------------------------
'                   CLASS EVENTS
'------------------------------------------------
Private Sub Class_Terminate()
    'make sure we are not still trying to move the mouse
    Call CancelMouseMove
End Sub
