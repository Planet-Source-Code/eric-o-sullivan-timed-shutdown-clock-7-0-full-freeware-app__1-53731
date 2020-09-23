VERSION 5.00
Object = "{072A0CD7-4439-46FD-9BC0-FF8959716B3B}#2.0#0"; "SYSTEM~1.OCX"
Begin VB.Form frmClock 
   Caption         =   "Clock"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1815
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   121
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin SystemTrayIcon.ctlSysTray sysIcon 
      Left            =   0
      Top             =   1320
      _ExtentX        =   529
      _ExtentY        =   503
      ToolTip         =   ""
   End
   Begin VB.Timer timRefresh 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   720
      Top             =   1200
   End
   Begin VB.Label lblEggIcon 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblEggRun 
      BackStyle       =   0  'Transparent
      Caption         =   "  Egg1 - Show's a Run dialog box"
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "&System Tray Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuSysShow 
         Caption         =   "&Show Clock"
      End
      Begin VB.Menu mnuSysQuitBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuClock 
      Caption         =   "&Clock"
      Visible         =   0   'False
      Begin VB.Menu mnuClock24Hour 
         Caption         =   "24 &Hour"
      End
      Begin VB.Menu mnuClockAnalog 
         Caption         =   "Analo&g"
      End
      Begin VB.Menu mnuClockAlarmBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClockAlarm 
         Caption         =   "A&larm Settings..."
      End
      Begin VB.Menu mnuClockDisplay 
         Caption         =   "&Display Settings..."
      End
      Begin VB.Menu mnuClockPasswordBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClockPassword 
         Caption         =   "&Password"
         Begin VB.Menu mnuClockPassEnter 
            Caption         =   "&Enter Password"
         End
         Begin VB.Menu mnuClockPassEnable 
            Caption         =   "&Password Enabled"
         End
         Begin VB.Menu mnuClockPassLock 
            Caption         =   "&Lock Program"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuClockAdvanced 
         Caption         =   "Advanced"
         Begin VB.Menu mnuClockAdvOnTop 
            Caption         =   "&Always On Top"
         End
         Begin VB.Menu mnuClockAdvStartup 
            Caption         =   "&Run At Startup"
         End
         Begin VB.Menu mnuClockAdvSnap 
            Caption         =   "S&nap Clock"
         End
         Begin VB.Menu mnuClockAdvStartMin 
            Caption         =   "Start &Minimized"
         End
         Begin VB.Menu mnuClockAdvHidden 
            Caption         =   "Start &Hidden"
         End
         Begin VB.Menu mnuClockAdvTrans 
            Caption         =   "Trans&parent"
            Begin VB.Menu mnuClockAdvTransOpaque 
               Caption         =   "&Opaque"
               Index           =   0
            End
            Begin VB.Menu mnuClockAdvTransOpaque 
               Caption         =   "&10%"
               Index           =   1
            End
            Begin VB.Menu mnuClockAdvTransOpaque 
               Caption         =   "&20%"
               Index           =   2
            End
            Begin VB.Menu mnuClockAdvTransOpaque 
               Caption         =   "&30%"
               Index           =   3
            End
            Begin VB.Menu mnuClockAdvTransOpaque 
               Caption         =   "&40%"
               Index           =   4
            End
            Begin VB.Menu mnuClockAdvTransOpaque 
               Caption         =   "&50%"
               Index           =   5
            End
            Begin VB.Menu mnuClockAdvTransOpaque 
               Caption         =   "&60%"
               Index           =   6
            End
            Begin VB.Menu mnuClockAdvTransOpaque 
               Caption         =   "&70%"
               Index           =   7
            End
            Begin VB.Menu mnuClockAdvTransOpaque 
               Caption         =   "&80%"
               Index           =   8
            End
            Begin VB.Menu mnuClockAdvTransOpaque 
               Caption         =   "&90%"
               Index           =   9
            End
         End
         Begin VB.Menu mnuClockAdvReloadBreak 
            Caption         =   "-"
         End
         Begin VB.Menu mnuClockAdvChangeTime 
            Caption         =   "&Change Local Time"
         End
         Begin VB.Menu mnuClockAdvIcon 
            Caption         =   "Rel&oad Tray Icon"
         End
         Begin VB.Menu mnuClockAdvShutdownBreak 
            Caption         =   "-"
         End
         Begin VB.Menu mnuClockAdvShutdown 
            Caption         =   "&Log Off"
            Index           =   0
         End
         Begin VB.Menu mnuClockAdvShutdown 
            Caption         =   "&Force Close"
            Index           =   1
         End
         Begin VB.Menu mnuClockAdvShutdown 
            Caption         =   "&Shutdown"
            Index           =   2
         End
         Begin VB.Menu mnuClockAdvShutdown 
            Caption         =   "&Power Down"
            Index           =   3
         End
         Begin VB.Menu mnuClockAdvShutdown 
            Caption         =   "R&eboot"
            Index           =   4
         End
         Begin VB.Menu mnuClockAdvLock 
            Caption         =   "Loc&k Workstation"
         End
      End
      Begin VB.Menu mnuClockAboutBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClockAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuClockExitBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClockExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     2 August 2002
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    Timed Shutdown Clock
' -----------------------------------------------
'COMMENTS :
'This is the main form for the program
'Timed Shutdown Clock, and is dependant on several
'modules and classes to work correctly. It is
'intended only for use within this program. It's
'purpose is to shutdown the computer either at a
'predefined time or at a specified time.
'=================================================

'require variable declaration
Option Explicit

'------------------------------------------------
'               GLOBAL-LEVEL VARIABLES
'------------------------------------------------
'This variable is public so that it's settings can
'be set from the module modMain
Public WithEvents mtscAnalog    As clsTimedClock    'used to display the analog time and to trap events
Attribute mtscAnalog.VB_VarHelpID = -1

'------------------------------------------------
'               MODULE-LEVEL VARIABLES
'------------------------------------------------
'local variables
Private mblnLoaded              As Boolean          'used to notify other methods that the form is loading

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Private Sub Form_Activate()
    'display the clocks
    Call ShowClock
    DoEvents
    mblnLoaded = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    'check for the menu key
    
    Const MENU_KEY  As Integer = 93
    
    If KeyCode = MENU_KEY Then
        'display the normal menu
        Call Me.PopupMenu(mnuClock)
    End If
End Sub

Private Sub Form_Load()
    'load all the necessary settings before displaying
    'the form
    
    Dim bmpTitle As clsBitmap   'to display a maximise animation
    
    'load the system tray icon
    With sysIcon
        Set .Icon = Me.Icon
        .ToolTip = Now
        
        'display the icon
        .Show
    End With
    
    'create new clock objects
    Set mtscAnalog = New clsTimedClock
    
    'apply all appropiate settings
    Call SetAllSettings
    
    'set the menu's to display the appropiate clock
    'information
    mnuClock24Hour.Checked = gbln24Hour
    mnuClockAnalog.Checked = gblnShowAnalog
    mnuClockAdvStartMin.Checked = gblnStartMin
    mnuClockAdvStartup.Checked = gblnRunAtStartup
    mnuClockAdvHidden.Checked = gblnStartHidden
    mnuClockAdvSnap.Checked = gblnSnapClock
    mnuClockAdvOnTop.Checked = gblnIsOnTop
    
    'if this is NOT a windows NT machine then we cannot "Lock" the computer so disable this option
    If Not IsWinNT Then
        mnuClockAdvLock.Enabled = False
    End If
    
    'enable/disable the menu's is the password is
    'active
    Call EnableMenus(Not gblnPassOn)
    
    'minimize the clock
    If gblnStartMin Then
        Me.WindowState = vbMinimized
    End If
    
    'move the clock to the bottom right of the screen
    If Me.WindowState <> vbMinimized Then
        Call ResizeForAnalog
        Call Me.Move(gintLastPos(x), _
                     gintLastPos(Y), _
                     gintWidth, _
                     gintHeight)
        Call CheckIfOutsideScreen(Me)
    End If
    
    'display the form
    If Not gblnStartHidden Then
        Set bmpTitle = New clsBitmap
        Call bmpTitle.TrayToTitle(Me)
        Call Me.Show
    End If
    
    'set the form fonts to the current system default
    Call SetFormFontsToSystem(Me, FNT_MESSAGE)
    
    'start keeping track of time
    timRefresh.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           Y As Single)
    'display the pop-up menu when the right mouse
    'button is clicked anywhere on the form. If
    'a left mouse button is used, assume that the
    'user is trying to drag the clock.
    
    Dim mosGrab     As clsMouse     'used to grab the clock at the current point the mouse is. This is like the click-drag operation of the title bar
    
    Select Case Button
    Case vbRightButton
        'display menu
        Call Me.PopupMenu(mnuClock)
    
    Case vbLeftButton
        'drag the clock
        Set mosGrab = New clsMouse
        
        With mosGrab
            Call .GrabWindow(Me.hWnd)
        End With    'mosGrab
    End Select
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'this will release the click-drag operation of the left button
    'if the clock is currently being dragged
    
    Dim mosGrab         As clsMouse         'used to query information about the clock
    
    Select Case Button
    Case vbLeftButton
        Set mosGrab = New clsMouse
        With mosGrab
            If (.GetGrabbedWindow = Me.hWnd) Then
                'release the window
                Call .ReleaseWindow
            End If
        End With    'mosGrab
    End Select
End Sub

Private Sub Form_Paint()
    'repaint the display
    Call mtscAnalog.PaintClock
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    'check to see if we should have exclusive rights
    'to shut windows down
    
    Dim bmpTitle As clsBitmap   'used to animate the title bar should we need to "minimize" the form (ie, hide it)
    
    Select Case UnloadMode
    Case vbFormCode, vbAppWindows, vbAppTaskManager
        'windows is closing
        Call PreUnload
        Exit Sub
    End Select
    
    'if the timed shutdown is active, then disregard
    'the close
    If gudtDay(Today).blnDoShutdown Then
        'don't unload the form, just hide it
        
        'display the animation
        Set bmpTitle = New clsBitmap
        Call bmpTitle.TitleToTray(Me)
        
        'cancel the program shut down
        Me.Visible = False
        mtscAnalog.Visible = False
        Cancel = True
        Exit Sub
    End If
    
    'should we prevent closing
    If (gblnExclusiveShut) Then
        'prevent shutdown
        Cancel = True
        Exit Sub
    End If
    
    'the form is closing
    Call PreUnload
End Sub

Private Sub Form_Resize()
    'make appropiate changes when the clock is resized
    
    Static intPrevState As Integer  'holds the windows previous WindowState prior to this call

    'has it been minimized or restored
    If intPrevState <> Me.WindowState Then
        'the form has been minimized or restored
        Select Case Me.WindowState
        Case vbMinimized
            Me.Caption = Time
        Case vbNormal
            Me.Caption = "Clock"
        End Select

        'update the current state
        intPrevState = Me.WindowState
    End If
    
    'adjust the size of the clock being displayed
    With mtscAnalog
        Call .ReSize(Me.ScaleWidth, Me.ScaleHeight)
    End With
    
    'make sure the labels triggering the eggs are on the form
    With lblEggIcon
        .Left = Me.ScaleWidth - .Width
    End With
    
    'make sure the clock does not go outside the users work area
    Call CheckIfOutsideScreen(Me)
    
    'update the size and dimensions
    gintWidth = Me.Width
    gintHeight = Me.Height
    gintLastPos(x) = Me.Left
    gintLastPos(Y) = Me.Top
End Sub

Private Sub lblEggIcon_MouseDown(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 Y As Single)
    'change the system tray icon to a smiley face :) if
    'the egg is activated
    
    Static blnEggOn As Boolean  'is the egg active or not
    
    If (Button = vbLeftButton) And _
       (Shift = (vbAltMask + vbShiftMask)) Then
        If Not blnEggOn Then
            'activate the egg
            Set sysIcon.Icon = LoadResPicture("EggTray", vbResIcon)
        Else
            'deactivate the egg
            Set sysIcon.Icon = Me.Icon
        End If
        blnEggOn = Not blnEggOn
    Else
        'pass event control to the form for processing
        Call Form_MouseDown(Button, Shift, x, Y)
    End If
End Sub

Private Sub lblEggRun_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                Y As Single)
    'display the run dialog box if necessary
    
    'do we display the run dialog box?
    If (Button = vbRightButton) And _
       (Shift = (vbShiftMask Or vbCtrlMask)) Then
        'if the right mouse button is clicked while
        'holding the Shift and Ctrl keys, then display
        'the program egg (a Run dialog box)
        frmEggRun.Show
    
    ElseIf (Button = vbLeftButton) And _
           (Shift = (vbAltMask Or vbCtrlMask)) Then
        'if holding the Control and Alt buttons and
        'while clicking the Left mouse button, then
        'display the Eyes egg form
        frmEggEyes.Show
    
    Else
        'pass mouse info to the form for handling
        Call Form_MouseDown(Button, Shift, x, Y)
    End If
End Sub

Private Sub mnuClock24Hour_Click()
    'do we display the 24 hour clock or am/pm
    gbln24Hour = Not gbln24Hour
    mnuClock24Hour.Checked = gbln24Hour
    mtscAnalog.Time24Hour = gbln24Hour
End Sub

Private Sub mnuClockAdvChangeTime_Click()
    'display the Change Time screen
    DoEvents
    frmChangeTime.Show
End Sub

Private Sub mnuClockAdvHidden_Click()
    'Should the clock be displayed immediatly when it starts
    gblnStartHidden = Not gblnStartHidden
    mnuClockAdvHidden.Checked = gblnStartHidden
End Sub

Private Sub mnuClockAdvIcon_Click()
    'reload the icon in the system tray
    Call sysIcon.Show
End Sub

Private Sub mnuClockAdvLock_Click()
    'this will lcok the workstation as long as this is a windows NT based machine
    If IsWinNT Then
        Call WINLock
    End If
End Sub

Private Sub mnuClockAdvOnTop_Click()
    'is the window the top-most window or not
    gblnIsOnTop = Not gblnIsOnTop
    mnuClockAdvOnTop.Checked = gblnIsOnTop
    If gblnIsOnTop Then
        Call StayOnTop(Me)
    Else
        Call NotOnTop(Me)
    End If
End Sub

Private Sub mnuClockAdvShutdown_Click(Index As Integer)
    'shutdown the computer the appropiate way
    DoEvents
    Call DoShutMethod(Index)
End Sub

Private Sub mnuClockAdvSnap_Click()
    'do we snap the clock to the side of the screen
    gblnSnapClock = Not gblnSnapClock
    mnuClockAdvSnap.Checked = gblnSnapClock
    
    'check clock bound if turned on
    If gblnSnapClock Then
        Call SnapWindow(Me, SNAP_DISTANCE)
    End If
End Sub

Private Sub mnuClockAdvStartMin_Click()
    'should the clock start minimized
    gblnStartMin = Not gblnStartMin
    mnuClockAdvStartMin.Checked = gblnStartMin
    'takes effect on next program start up
End Sub

Private Sub mnuClockAdvStartup_Click()
    'should the program start when windows starts
    gblnRunAtStartup = Not gblnRunAtStartup
    mnuClockAdvStartup.Checked = gblnRunAtStartup
    Call RunAtStartup
End Sub

Private Sub mnuClockAdvTransOpaque_Click(Index As Integer)
    'set the transparent level of the form
    
    'uncheck the last item
    mnuClockAdvTransOpaque(gsngTransLvl \ 10).Checked = False
    
    gsngTransLvl = (Index * 10)
    Call MakeWndTransparent(Me.hWnd, gsngTransLvl)
    
    'check the new tranparent level
    mnuClockAdvTransOpaque(gsngTransLvl \ 10).Checked = True
End Sub

Private Sub mnuClockAnalog_Click()
    'do we display the analog clock or not
    gblnShowAnalog = Not gblnShowAnalog
    mnuClockAnalog.Checked = gblnShowAnalog
    mtscAnalog.ShowAnalog = gblnShowAnalog
    Call ResizeForAnalog
End Sub

Private Sub mnuClockPassEnable_Click()
    'enable or disable the menus
    With mnuClockPassEnable
        .Checked = Not .Checked
        gblnPassOn = Not gblnPassOn
        Call EnableMenus(Not gblnPassOn)
    End With
End Sub

Private Sub mnuClockPassEnter_Click()
    'allow the user to enter the password
    DoEvents
    Load frmPassword
    Call frmPassword.SetScreen
End Sub

Private Sub mnuClockPassLock_Click()
    'lock all the menus
    Call EnableMenus(False)
End Sub

Private Sub mtscAnalog_NewDay(ByVal NewDate As String)
    'update the alarm time
    mtscAnalog.AlarmTime = GetShutdownTime
End Sub

Private Sub mtscAnalog_AlarmActivate()
    'try shut down the computer
    
    Dim strAlarmTime As String  'the time we are meant to shutdown the computer at
    
    'get the correct time we are meant to shut down the
    'computer at
    strAlarmTime = GetShutdownTime
    
    'make sure that this is the allotted time
    If DateDiff("s", Time, strAlarmTime) = 0 Then
        'this is the appropiate time
        Load frmShut
        Call frmShut.Start
    Else
        'The allotted time is not at this moment.
        'Update the alarm
        mtscAnalog.AlarmTime = strAlarmTime
    End If
End Sub

Private Sub mtscAnalog_NewTime(ByVal NewTime As Date)
    'update the tooltip to display the current date time
    sysIcon.ToolTip = Format(Date, "Long Date") + " " + _
                      Format(Time, "h:nn am/pm")
    
    'check the idle time
    Call UpdateClockAlarm
    
    'check if we need to free up memory
    If gblnDoCleanUp Then
        Call DoCleanUp(Me)
    End If
    
    'update the caption if minimized
    If Me.WindowState = vbMinimized Then
        Me.Caption = Time
    Else
        'repaint the display
        Call mtscAnalog.PaintClock
    End If
End Sub

Private Sub mnuClockAbout_Click()
    'display the about form
    
    Dim strWavPath      As String       'holds the complete path to the sound file
    
    'continue graphic for the menu before tying up the cpu with other things
    DoEvents
    
    'get the location to the sound file
    strWavPath = AddToPath(App.Path, "TaDa.wav")
    
    'play the sound if it exists in the programs directory
    If (Dir(strWavPath) <> "") Then
        'play file asynchrinosly
        Call PlaySound(strWavPath)
    End If
    
    'display the about acreen
    frmAboutScreen.Show
End Sub

Private Sub mnuClockAlarm_Click()
    'display the alarm settings
    DoEvents
    frmOptions.Show
End Sub

Private Sub mnuClockDisplay_Click()
    'show the Display Settings form to change the
    'visual appearance of the form
    DoEvents
    frmDisplay.Show
End Sub

Private Sub mnuClockExit_Click()
    'pass unloading control to the appropiate event to
    'test if we exit the program or stay active
    DoEvents
    Unload Me
End Sub

Private Sub mnuSysQuit_Click()
    'exit the program
    DoEvents
    Unload Me
End Sub

Private Sub mnuSysShow_Click()
    'display the clock
    
    Dim bmpTitle As clsBitmap   'display animation if necessary
    
    DoEvents
    
    'if the form was hidden, the display an animation to
    'show where the form is going to appear
    If Not Me.Visible Then
        Set bmpTitle = New clsBitmap
        Call bmpTitle.TrayToTitle(Me)
    End If
    
    'make sure the application is "activated". This may
    'cause an error if another application is registered
    'and running with the same application title (I only
    'encounter this when using the vb development
    'environment)
    On Error Resume Next
        Call AppActivate(App.Title)
    On Error GoTo 0
    
    'show the form
    Call ShowClock
End Sub

Private Sub sysIcon_DblClick(ByVal Button As Integer, _
                             ByVal Shift As Integer, _
                             ByVal x As Integer, _
                             ByVal Y As Integer)
    'do the default menu
    Call mnuSysShow_Click
End Sub

Private Sub sysIcon_MouseDown(ByVal Button As Integer, _
                              ByVal Shift As Integer, _
                              ByVal x As Integer, _
                              ByVal Y As Integer)
    'display the menu
    If Button = vbRightButton Then
        'make sure that the focus is on the application
        'before trying to display the popup menu
        
        'make sure the application is "activated". This may
        'cause an error if another application is registered
        'and running with the same application title (I only
        'encounter this when using the vb development
        'environment)
        On Error Resume Next
            Call AppActivate(App.Title)
        On Error GoTo 0
        
        'display the popup menu
        Call PopupMenu(mnuSysTray, _
                       DefaultMenu:=mnuSysShow)
    End If
End Sub

Private Sub timRefresh_Timer()
    'This updates the actual clock display
    If Not mtscAnalog Is Nothing Then
        mtscAnalog.Refresh
    End If

'    -- debug for wallpaper on win95/98/Xp : 2/jan/2003
    'Dim test As clsBitmap
    'Set test = New clsBitmap
    'Call test.DrawBorder(15, 15, 25, 25, BDR_SUNKEN, BF_RECT, picTest.hDc)
    'Call test.DrawBorder(15, 15, 25, 25, BDR_RAISED, BF_RECT, picTest.hDc)
    'Call test.RotateBitmap(Me.hDc, picTest.hDc, 45, 0, 0, 0, 0, 20, 20)
    'Call test.Paint(picTest.hDc)
    
    'Call test.FlashWindow(Me.hwnd)
    'Call test.SetBitmap(picTest.ScaleHeight, _
                        picTest.ScaleWidth, _
                        picTest.BackColor, _
                        picTest.hdc)
    'Call test.GetWallpaper
    'Set picTest.Picture = test.Picture
    'Call test.Paint(picTest.hDc)
End Sub

Private Sub ShowClock()
    'display the clocks on the form
    
    Dim bmpMoveWin      As clsBitmap        'used to move the window to the top of the zorder
    
    DoEvents
    Set bmpMoveWin = New clsBitmap
    Me.Visible = True
    mtscAnalog.Visible = True
    
    'move the window to the top of the zorder
    Call bmpMoveWin.PutWindowInFore(Me.hWnd)
    
    'do we force the program to be the top-most window
    If gblnIsOnTop Then
        Call StayOnTop(Me)
    End If
End Sub

Public Sub UpdateClockAlarm()
    'This is used so that if settings are changed, this
    'form can react to the new changes
    With mtscAnalog
        If gudtDay(Today).blnDoShutdown Then
            'set the alarm time
            .AlarmTime = GetShutdownTime
        Else
            'turn off the alarm
            .AlarmTime = ""
        End If
    End With
End Sub

Private Sub PreUnload()
    'ok to shut down. clean up memory
    
    'make sure that all timers are switched off
    timRefresh.Enabled = False
    
    'remove objects from memory (trigger the
    'Class_Terminate events)
    Set mtscAnalog = Nothing
    
    With Me
        If (.WindowState <> vbMinimized) Then
            gintLastPos(x) = .Left
            gintLastPos(Y) = .Top
            gintHeight = .Height
            gintWidth = .Width
        End If
    End With
    
    'make sure that all current settings are saved
    Call SaveSettings
    
    'stop tracking idle time
    Call StopTracking
    
    'make sure that the form is not on top anymore
    'if active
    If gblnIsOnTop Then
        Call NotOnTop(Me)
    End If
    
    'remove all forms from memory to end the program.
    'This is done to ensure that memory is cleaned up
    'appropiatly before exiting the program
    Call UnloadAll(Me)
End Sub

Private Sub ResizeForAnalog()
    'This will resize the form depending on whether or
    'not the analog clock is to be displayed or not
    
    Dim intBorderSize As Integer
    
    'we can't change the size of a window if it is
    'maximized or minimized
    If Me.WindowState = vbNormal Then
        'get the size of the border and title bar
        intBorderSize = (Me.Height - (CLng(Me.ScaleHeight) * Screen.TwipsPerPixelY))
        
        If gblnShowAnalog Then
             'size for an analog clock
            Me.Height = (mtscAnalog.Height * _
                         CLng(Screen.TwipsPerPixelY)) _
                        + intBorderSize
            
            'make sure the clock is not out of the screen
            Call SnapWindow(Me, SNAP_DISTANCE)
       Else
            'size for no analog clock
            Me.Height = ((mtscAnalog.Height - _
                          mtscAnalog.AnalogHeight) _
                         * CLng(Screen.TwipsPerPixelY)) _
                        + intBorderSize
        End If
        
        'after the form has been resize, redisplay the
        'clocks graphic
        Call mtscAnalog.PaintClock
    End If
End Sub

Public Sub EnableMenus(Optional ByVal blnEnable As Boolean = True)
    'This will enable or disable certain menus that
    'access any settings for the clock. This is
    'included as a minor security feature.
    
    'set the main menus
    mnuClock24Hour.Enabled = blnEnable
    mnuClockAnalog.Enabled = blnEnable
    mnuClockAlarm.Enabled = blnEnable
    mnuClockDisplay.Enabled = blnEnable
    mnuClockAdvanced.Enabled = blnEnable
    
    'set the sub menus of "Password >"
    If gblnPassOn And blnEnable Then
        'only enabled when the password has been entered
        'and all menu's are being shown
        mnuClockPassLock.Enabled = True
    Else
        mnuClockPassLock.Enabled = False
    End If
    'only allow the user to activate/deactivate the
    'password once it has been entered
    mnuClockPassEnable.Enabled = gblnPassOn And blnEnable
    mnuClockPassEnable.Checked = gblnPassOn
End Sub

Public Sub CheckDrag()
    'This procedure will check to see if the form is
    'being dragged, and will "snap" the window if the
    'option is turned on. This procedure is called
    'whenever the mouse is moved (see procedure NotIdle
    'in modIdleTime)
    
    Static intX As Integer
    Static intY As Integer

    With Me
        'make sure that the clock is not minimized
        If .WindowState = vbMinimized Then
            Exit Sub
        End If

        If (.Left = intX) And (.Top = intY) Then
            'the form has not changed position, exit
            Exit Sub
        End If

        'snap the clock to the side of the screen if we
        'need to
        If gblnSnapClock Then
            Call SnapWindow(Me, SNAP_DISTANCE)
        End If
        
        'update the position of the form
        intX = .Left
        intY = .Top
    End With
    
    'if the display is set to Wallpaper, then update the
    'background picture to match the area the clock is
    'being dragged over
    With mtscAnalog
        If .BackgroundStyle = clkWallpaper Then
            Call mtscAnalog.GetScreenPos(Me)
            Call .Refresh
        End If
    End With
End Sub
