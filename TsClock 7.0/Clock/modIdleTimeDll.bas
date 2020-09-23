Attribute VB_Name = "modIdleTime"
'=================================================
'AUTHOR : Eric O'Sullivan
' -----------------------------------------------
'DATE : 2/Aug/2002
' -----------------------------------------------
'CONTACT: DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE : Track Idle Time Module
' -----------------------------------------------
'COMMENTS :
'This module is used to keep track of the mouse
'and keyboard. The procedure NotIdle is intended
'for use as an "event" when the user presses a
'key or moves the mouse. To initiate the tracking,
'call the SysStartTracking procedure. To cancel the
'tracking, call the SysStopTracking procedure.
'
'NB, Please note the following;
'1) If SysStartTracking has been called within your
'   program, SysStopTracking MUST also be called
'   within your program.
'
'2) You must NOT use End within your program
'   anywhere. Doing so can cause the vb ide to
'   crash.
'
'3) You must NOT press the Stop button on the
'   visual basic toolbar when debugging. Doing so
'   can cause the vb ide to crash.
'
'4) To exit your program, make a call to
'   SysStopTracking in the QueryUnload event on one
'   of your forms, and unload all forms to exit
'   the program. THIS IS THE ONLY SAFE WAY TO EXIT
'   YOUR PROGRAM!
'
'5) If you are making any calls to external
'   procedures/functions from within this module,
'   be aware that any errors raised within that
'   code is liable to cause the vb ide to crash.
'   Make sure any external code is BUG FREE before
'   making the call.
'
'6) To be certain there is no loss of data when
'   using this module, go to the Tools menu in
'   Visual Basic. Select Options and go to the
'   Environment tab. In the frame "When Program
'   Starts", select the "Save Changes" option
'   and then press Ok. This will automatically
'   save your program when ever you try to run it
'
'DISCLAIMER:
' I am NOT responsible for any loss of data or
' any form of system related problems as a direct
' or indirect result of using this module. You use
' this module ENTIRLY AT YOUR OWN RISK. You should
' not have any problems provided that you follow
' the instructions EXACTLY. I am also not
' responsible for any modifications made to the
' code by anyone except myself.
'=================================================

Option Explicit

'-----------------------------------------------------
' API Declarations
'-----------------------------------------------------

'make sure other hooks in the queue are called after
'our one has been processed
Private Declare Function CallNextHookEx _
        Lib "user32" _
            (ByVal hHook As Long, _
             ByVal ncode As Long, _
             ByVal wParam As Long, _
             ByRef lParam As Any) _
             As Long

'this will get the state of any specified key - even
'the mouse buttons
Private Declare Function GetAsyncKeyState _
       Lib "user32" _
            (ByVal vKey As Long) _
             As Integer

'stores the curosrs position in the PointAPI structure
'The values returned are in pixels
Private Declare Function GetCursorPos _
        Lib "user32" _
            (ByRef lpPoint As PointAPI) _
             As Long

'returns the amount of time windows has been active for
'in milliseconds (sec/1000)
Private Declare Function GetTickCount _
        Lib "kernel32" _
            () _
             As Long

'create a new hook in the specified hook queue for
'a function either for a thread or the system
Private Declare Function SetWindowsHookEx _
        Lib "user32" _
        Alias "SetWindowsHookExA" _
            (ByVal idHook As Long, _
             ByVal lpfn As Long, _
             ByVal hmod As Long, _
             ByVal dwThreadId As Long) _
             As Long

'unhook a specified hook from the hook queue
Private Declare Function UnhookWindowsHookEx _
        Lib "user32" _
            (ByVal hHook As Long) _
             As Long

'-----------------------------------------------------
' Constants
'-----------------------------------------------------
Private Const WH_KEYBOARD = 2   'The hook code for the keyboard
Private Const WH_MOUSE = 7      'The hook code for the mouse

'-----------------------------------------------------
' Enumerators
'-----------------------------------------------------
'tells where the NotIdle is activate from
Public Enum EnumCalledFrom
    FROM_KEYBOARD = 0
    FROM_MOUSE = 1
End Enum

'-----------------------------------------------------
' Types
'-----------------------------------------------------
'Holds the position of the mouse
Private Type PointAPI
    X As Long
    Y As Long
End Type

'-----------------------------------------------------
' Variables
'-----------------------------------------------------
Private mlngStartTick As Long       'the starting tick that this was activated on
Private mblnTracking As Boolean     'a flag that is set when we are tracking the idle time
Private mlngKeyboardHook As Long    'the windows hook for the keyboard
Private mlngMouseHook As Long       'the windows hook for the mouse


'-----------------------------------------------------
' Procedures (Events)
'-----------------------------------------------------
Public Sub NotIdle(ByVal enmFrom As EnumCalledFrom, _
                   Optional ByVal intMouseX As Integer = -1, _
                   Optional ByVal intMouseY As Integer = -1)
    'This procedure is programmer defined. Ideally, I'd
    'like to put this into a class, but unfortunatly
    'this just isn't possible. This procedure basically
    'is to be treated as an Event and can be copied into
    'any other module as long as it does not affect the
    'calls to it made below.
    Debug.Print "New Idle Time Triggered At " & Time
End Sub

'-----------------------------------------------------
' Procedures
'-----------------------------------------------------
Public Sub SysStartTracking()
    'This will start tracking the idle time
    
    'make sure that we only have one set of hooks
    'active at any one time
    Call SysStopTracking
    
    'set the hooks on the mouse and keyboard
    mlngKeyboardHook = SetWindowsHookEx(WH_KEYBOARD, _
                                        AddressOf KeyboardProc, _
                                        App.hInstance, _
                                        App.ThreadID)
    If mlngKeyboardHook = 0 Then
        Exit Sub
    End If
    
    mlngMouseHook = SetWindowsHookEx(WH_MOUSE, _
                                     AddressOf MouseProc, _
                                     App.hInstance, _
                                     App.ThreadID)
    
    If mlngMouseHook = 0 Then
        Call SysStopTracking
        Exit Sub
    End If
    
    'we are tracking the idle time
    mlngStartTick = GetTickCount
    mblnTracking = True
End Sub

Public Sub SysStopTracking()
    'this will stop tracking the idle time
    
    Dim lngResult As Long   'any returned error value from the api call
    
    If mlngKeyboardHook <> 0 Then
        lngResult = UnhookWindowsHookEx(mlngKeyboardHook)
        mlngKeyboardHook = 0
    End If
    If mlngMouseHook <> 0 Then
        lngResult = UnhookWindowsHookEx(mlngMouseHook)
        mlngMouseHook = 0
    End If
    
    'we are no longer tracking the idle time
    mblnTracking = False
End Sub

Public Function SysCurrentIdleTime() As Long
    'The current amount of time spent idle in Ticks
    '(milliseconds, second/1000)
    
    If mblnTracking Then
        SysCurrentIdleTime = GetTickCount - mlngStartTick
    End If
End Function

Private Function KeyboardProc(ByVal lngCode As Long, _
                              ByVal lngWParam As Long, _
                              ByVal lngLParam As Long) _
                              As Long
    'This procedure is called when the keyboard has
    'been activated. We must update the idle time and
    'make sure that other hooks in the queue are also
    'passed the same information.
    
    'process other hooks in the queue
    KeyboardProc = CallNextHookEx(mlngKeyboardHook, _
                                  lngCode, _
                                  lngWParam, _
                                  ByVal lngLParam)
    
    'update the idle time
    mlngStartTick = GetTickCount
    
    'trigger the event with new information
    Call NotIdle(FROM_KEYBOARD)
End Function

Private Function MouseProc(ByVal lngCode As Long, _
                           ByVal lngWParam As Long, _
                           ByVal lngLParam As Long) _
                           As Long
    'This procedure is called when the keyboard has
    'been activated. We must update the idle time and
    'make sure that other hooks in the queue are also
    'passed the same information.
    
    Dim udtMouse As PointAPI    'the current co-ordinates of the mouse
    Dim lngResult As Long       'holds any returned error value from the api calls
    
    'process other hooks in the queue
    MouseProc = CallNextHookEx(mlngMouseHook, _
                               lngCode, _
                               lngWParam, _
                               ByVal lngLParam)
    
    'update the idle time
    mlngStartTick = GetTickCount
    
    'trigger the event with new information
    lngResult = GetCursorPos(udtMouse)
    Call NotIdle(FROM_MOUSE, _
                 udtMouse.X, _
                 udtMouse.Y)
End Function

Private Function GetMouseButtonState() As MouseButtonConstants
    'This will return which mouse buttons are currently
    'being pressed
    
    'Left   = 1  - vbLeftButton
    'Right  = 2  - vbRightButton
    'Middle = 4  - vbMiddleButton
    
    Const KEY_DOWN = -32767
    Const KEY_PRESSED = -32768
    Const KEY_NOT_PRESSED = 0
    
    Dim lngResult As Long   'holds the keystate of the mouse button
    Dim lngState As Long    'holds the state of all three buttons. The values are totaled
    
    'check the left button
    lngResult = GetAsyncKeyState(vbLeftButton)
    If lngResult <> KEY_NOT_PRESSED Then
        'the key has been/is being used
        lngState = vbLeftButton
    End If
    
    'check the right button
    lngResult = GetAsyncKeyState(vbRightButton)
    If lngResult <> KEY_NOT_PRESSED Then
        'the key has been/is being used
        lngState = lngState Or vbRightButton
    End If
    
    'check the middle button
    lngResult = GetAsyncKeyState(vbMiddleButton)
    If lngResult <> KEY_NOT_PRESSED Then
        'the key has been/is being used
        lngState = lngState Or vbMiddleButton
    End If
    
    'return the state of all three buttons
    GetMouseButtonState = lngState
End Function


