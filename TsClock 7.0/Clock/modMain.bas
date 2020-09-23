Attribute VB_Name = "modMain"
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     2 August 2002
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    Timed Shutdown Clock Main Module
' -----------------------------------------------
'COMMENTS :
'This is used to manage the start up of the
'program and any global data available throughout
'the project.
'=================================================

Option Explicit

'------------------------------------------------
'               API DECLARATIONS
'------------------------------------------------
'returns the amount of time windows has been active for
'in milliseconds (sec/1000)
Public Declare Function GetTickCount _
        Lib "kernel32" _
            () _
             As Long

'this will force the program to use the current windows themed controls
Private Declare Function InitCommonControls _
        Lib "comctl32.dll" _
            () As Long

'------------------------------------------------
'                 GLOBAL CONSTANTS
'------------------------------------------------
'declare constants for co-ordinates
Public Const X              As Integer = 0
Public Const Y              As Integer = 1

'the range of idle times in which you can shut down the computer
Public Const MIN_IDLE_TIME  As Integer = 300    '5 minutes (in seconds)
Public Const MAX_IDLE_TIME  As Integer = 18000  '5 hours (in seconds)

Public Const SNAP_DISTANCE  As Integer = 10     'snap the clock if withing x pixels
Public Const SCHEME_NAME    As String = "Schemes.col" 'the name of the file that holds the colour schemes

'BitBlt constants
Public Const SRCAND         As Long = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY        As Long = &HCC0020  ' (DWORD) dest = source
Public Const SRCERASE       As Long = &H440328  ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT      As Long = &H660046  ' (DWORD) dest = source XOR dest
Public Const SRCPAINT       As Long = &HEE0086  ' (DWORD) dest = source OR dest
Public Const SRCMERGE_COPY  As Long = &HC000CA  ' (DWORD) dest = (source AND pattern)
Public Const SRCMERGE_PAINT As Long = &HBB0226  ' (DWORD) dest = (NOT source) OR dest
Public Const SRCNOT_COPY    As Long = &H330008  ' (DWORD) dest = (NOT source)
Public Const SRCNOT_ERASE   As Long = &H1100A6  ' (DWORD) dest = (NOT src) AND (NOT dest)

Public Const PI             As Single = 3.14159265358979

'the flag returned from a common dialog when the user pressed the
'Cancel button
Public Const DLG_CANCEL     As Long = 0

'------------------------------------------------
'            MODULE-LEVEL CONSTANTS
'------------------------------------------------
'File constants
Private Const INI_NAME = "Settings.ini"     'the name of the ini file with the settings in it

'declare the ini file headers and setting names
Private Const HDR_MAIN                          As String = "CLOCK VALUES"
    Private Const SET_SHUTDOWN_METHOD           As String = "Shutdown Method"
    Private Const SET_PREVENT_OTHER_SHUTDOWN    As String = "Exclusive Shutdown"
    Private Const SET_DO_IDLE_SHUTDOWN          As String = "Idle Shutdown"
    Private Const SET_IDLE_TIME                 As String = "If Idle For"
    Private Const SET_ANALOG                    As String = "Show Analog"
    Private Const SET_24_HOUR                   As String = "Is 24 Hour"

Private Const HDR_PASSWORD                      As String = "PASSWORD SETTINGS"
    Private Const SET_PASSWORD                  As String = "User Password"
    Private Const SET_PASS_ACTIVE               As String = "Password Active"

Private Const HDR_DAY                           As String = "DAY SETTING FOR "
    Private Const SET_DELAY_TIME                As String = "Delay Time"
    Private Const SET_AT_HOUR                   As String = "Shutdown Hour"
    Private Const SET_AT_MINUTE                 As String = "Shutdown Minute"
    Private Const SET_AT_SECOND                 As String = "Shutdown Second"
    Private Const SET_DO_TIMED_SHUTDOWN         As String = "Do Timed Shutdown"
    Private Const SET_DO_DELAY                  As String = "Do Delay Time"

Private Const HDR_COLOUR                        As String = "COLOUR SETTINGS"
    Private Const SET_COL_HOUR                  As String = "Hour Hand"
    Private Const SET_COL_MINUTE                As String = "Minute Hand"
    Private Const SET_COL_SECOND                As String = "Second Hand"
    Private Const SET_COL_DOTS                  As String = "Dots"
    Private Const SET_COL_ANA_BACK              As String = "Analog Background"
    Private Const SET_COL_TIME_FONT             As String = "Time Font"
    Private Const SET_COL_TIME_BACK             As String = "Time Background"
    Private Const SET_COL_DAY_FONT              As String = "Day Font"
    Private Const SET_COL_DAY_BACK              As String = "Day Background"
    Private Const SET_COL_DATE_FONT             As String = "Date Font"
    Private Const SET_COL_DATE_BACK             As String = "Date Background"
    Private Const SET_COL_BORDER                As String = "Border"

Private Const HDR_BACKGROUND                    As String = "BACKGROUND SETTINGS"
    Private Const SET_BACKGROUND_PATH           As String = "Picture Path"
    Private Const SET_BACKGROUND_MODE           As String = "Mode"
    
Private Const HDR_MISC                          As String = "MISCELLANIOUS SETTINGS"
    Private Const SET_LAST_POS_X                As String = "Last Position X"
    Private Const SET_LAST_POS_Y                As String = "Last Position Y"
    Private Const SET_CLOCK_WIDTH               As String = "Clock Width"
    Private Const SET_CLOCK_HEIGHT              As String = "Clock Height"
    Private Const SET_IS_ON_TOP                 As String = "Put Clock On Top"
    Private Const SET_START_MIN                 As String = "Start Clock Minimized"
    Private Const SET_START_HIDDEN              As String = "Start Clock Hidden"
    Private Const SET_SNAP_WINDOW               As String = "Snap Clock"
    Private Const SET_OWNER                     As String = "Licenced to"
    Private Const SET_RUN_AT_STARTUP            As String = "Run At Startup"
    Private Const SET_TRANSPARENT               As String = "Transparent Level"
    Private Const SET_BORDER_WIDTH              As String = "Border Width"

Private Const HDR_DEBUG                         As String = "DEBUG INFO"
    Private Const SET_LAST_SAVE_TIME            As String = "Last Save Time"
    Private Const SET_LAST_SAVE_DATE            As String = "Last Save Date"
    Private Const SET_SHUTDOWN_AT               As String = "Shutdown Set For"
    Private Const SET_CURRENT_IDLE_TIME         As String = "Current Idle Time"

'general private constants
'------
Private Const STARTUP_LABEL As String = "Timed Shutdown Clock"   'a registry label used when the app should start up when windows starts
Private Const DEFAULT_OWNER As String = "Unknown"                'the default text to display if there is no licenced owner

'------------------------------------------------
'                   ENUMERATORS
'------------------------------------------------
'declare enumerators
Public Enum EnumShutdown        'the different methods of shutting the computer down
    shtLogOut = 0
    shtForceClose = 1
    shtShutdown = 2
    shtPowerDown = 3
    shtRestart = 4
    shtLockWorkstation = 5
End Enum

'------------------------------------------------
'               USER-DEFINED TYPES
'------------------------------------------------
'declare types
Public Type TypeShutdown        'shutdown information for each day
    intHour         As Integer  'the shutdown hour
    intMinute       As Integer  'the shutdown minute
    intSecond       As Integer  'the shutdown second
    intDelay        As Integer  'the amount of time to wait for a user response to cancel shutdown
    blnDoDelay      As Boolean  'do we automatically shutdown the computer after a specified time
    blnDoShutdown   As Boolean  'do we shutdown the computer
End Type

Public Type TypeSchemes         'the colour data for a colour scheme for the clock
    strName         As String * 50
    lngBorderColour As Long
    lngAnalogBack   As Long
    lngDots         As Long
    lngHourHand     As Long
    lngMinuteHand   As Long
    lngSecondHand   As Long
    lngTimeFont     As Long
    lngTimeBack     As Long
    lngDayFont      As Long
    lngDayBack      As Long
    lngDateFont     As Long
    lngDateBack     As Long
End Type

Public Type Rect
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

'------------------------------------------------
'               GLOBAL VARIABLES
'------------------------------------------------
'declare variables to hold the clock settings
Public genmShutdownMethod   As EnumShutdown         'what method to user shutting down the computer
Public gblnExclusiveShut    As Boolean              'shut we prevent windows from closing until the clock is ready
Public gblnIdleShut         As Boolean              'activate idle shutdown
Public glngIdleTime         As Long                 'the amount fo idle time to wait before shutting down
Public gblnShowAnalog       As Boolean              'display the analog clock
Public gbln24Hour           As Boolean              'display time in 24 hour format
Public gstrPassword         As String               'the password that unlocks the program features
Public gblnPassOn           As Boolean              'is the password currently activated
Public gudtDay(6)           As TypeShutdown         'the timed shutdown setting for each day
Public glngColHour          As Long                 'the colour of the hour hand
Public glngColMinute        As Long                 'the colour of the minute hand
Public glngColSecond        As Long                 'the colour of the second hand
Public glngColDots          As Long                 'the colour of the dots
Public glngColAnaBack       As Long                 'the colour of the analog background
Public glngColTimeFont      As Long                 'the colour of the time font
Public glngColTimeBack      As Long                 'the colour of the time background
Public glngColDayFont       As Long                 'the colour of the day font
Public glngColDayBack       As Long                 'the colour of the day background
Public glngColDateFont      As Long                 'the colour of the date font
Public glngColDateBack      As Long                 'the colour of the date background
Public glngColBorder        As Long                 'the colour of the border around the clock (if any)
Public gstrBackPath         As String               'the location of the background picture
Public genmBackStyle        As EnmCBackgroundStyle  'what way to display the background picture
Public gintLastPos(1)       As Integer              'the last known co-ordinate of the clock (usually so we can startup where the user put the clock last
Public gintWidth            As Integer              'the width of the clock
Public gintHeight           As Integer              'the height of the clock
Public gblnIsOnTop          As Boolean              'keep the clock the top-most window
Public gblnStartMin         As Boolean              'start the clock minimized
Public gblnStartHidden      As Boolean              'start the clock, but don't display it
Public gblnSnapClock        As Boolean              'snap the clock window to the edge of the screen/work area
Public gblnRunAtStartup     As Boolean              'run the clock at windows startup
Public gstrOwner            As String               'the licenced owner of the machine
Public gsngTransLvl         As Single               'the transparent level that the clock is displayed at
Public gintBorderWidth      As Integer              'the width of the border around the clock
Public gstrSchemePath       As String               'the path to the colour schemes file (always in the current directory)
Public gblnDoCleanUp        As Boolean              'flags if the program should check all forms and make sure that they are set to nothing except for the main form

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Public Sub Main()
    'This main starting procedure for this
    'project
    
    Dim udtColours As TypeSchemes   'used to create the colour schemes file (if necessary)
    Dim lngSuccess As Long         'holds wehther or not the api was successful
    
    
    'make sure there is no currently running
    'instance of this application
    If App.PrevInstance Then
        End
    End If
    
    'increase the thread priority temperorily
    Call SetAppThreadPriority(THREAD_PRIORITY_HIGHEST, False)
    
    'This will force the program to use the Manifest file for windows controls. What this means
    'is that on a xp machine, if an xp theme is being used, then the controls will look like the theme controls
    If IsXp Then
        lngSuccess = InitCommonControls
    End If
    
    'hide the program from the Application list on windows NT based machines as it'll show up on the
    'process list anyway. Leave it there on w9x systems as there is no default process list
    If IsWinNT Then
        App.TaskVisible = False
    End If
    
    'get the path to the colour schemes
    gstrSchemePath = AddToPath(App.Path, SCHEME_NAME)
    
    If (Dir(gstrSchemePath) = "") Then
        'create the colour schemes file using the
        'default colours
        Call CreateScheme(udtColours, True)
    End If
    
    'get the programs settings
    Call LoadSettings
    
    'start tracking user response
    Call StartTracking
    
    'show the main form
    Load frmClock
    DoEvents
    
    'set the applications thread priority to "Low" because this app is
    'meant to be running in the background
    Call SetAppThreadPriority(THREAD_PRIORITY_BELOW_NORMAL, True)
End Sub

Public Sub SaveSettings(Optional ByVal blnReset As Boolean = False)
    'This procedure will save all the clocks settings
    
    Dim strFilePath     As String   'the location of the ini file
    Dim strEncrypted    As String   'the encrypted password
    Dim intCounter      As Integer  'used for cycling through each day
    Dim strDayHeader    As String   'used to hold a header for each day
    Dim blnNew          As Boolean  'holds whether or not the ini is just being created
    
    'get the path to save the ini file to
    strFilePath = AddToPath(App.Path, INI_NAME)
    
    'check if the ini file already exists
    If (Dir(strFilePath) = "") Or (blnReset) Then
        'allow formatting of the ini file
        blnNew = True
    End If
    
    'encrypt the password
    strEncrypted = EncryptText(gstrPassword)
    
    'save all settings
    
    'general settings
    Call WriteToIni(strFilePath, _
                    HDR_MAIN, _
                    SET_SHUTDOWN_METHOD, _
                    genmShutdownMethod)
    Call WriteToIni(strFilePath, _
                    HDR_MAIN, _
                    SET_PREVENT_OTHER_SHUTDOWN, _
                    gblnExclusiveShut)
    Call WriteToIni(strFilePath, _
                    HDR_MAIN, _
                    SET_DO_IDLE_SHUTDOWN, _
                    gblnIdleShut)
    Call WriteToIni(strFilePath, _
                    HDR_MAIN, _
                    SET_IDLE_TIME, _
                    glngIdleTime)
    Call WriteToIni(strFilePath, _
                    HDR_MAIN, _
                    SET_ANALOG, _
                    gblnShowAnalog)
    If blnNew Then  'new section after this line (a new line for formatting purposes if the ini file was reset or deleted)
        Call WriteToIni(strFilePath, _
                        HDR_MAIN, _
                        SET_24_HOUR, _
                        gbln24Hour & vbCrLf)
    Else
        Call WriteToIni(strFilePath, _
                        HDR_MAIN, _
                        SET_24_HOUR, _
                        gbln24Hour)
    End If
    
    'password settings
    Call WriteToIni(strFilePath, _
                    HDR_PASSWORD, _
                    SET_PASSWORD, _
                    strEncrypted)
    If blnNew Then
        Call WriteToIni(strFilePath, _
                        HDR_PASSWORD, _
                        SET_PASS_ACTIVE, _
                        gblnPassOn & vbCrLf)
    Else
        Call WriteToIni(strFilePath, _
                        HDR_PASSWORD, _
                        SET_PASS_ACTIVE, _
                        gblnPassOn)
    End If
    
    'daily settings
    For intCounter = 0 To 6
        'build the header
        strDayHeader = HDR_DAY & WeekdayName(intCounter + 1)
        
        With gudtDay(intCounter)
            Call WriteToIni(strFilePath, _
                            strDayHeader, _
                            SET_DO_DELAY, _
                            .blnDoDelay)
            Call WriteToIni(strFilePath, _
                            strDayHeader, _
                            SET_DELAY_TIME, _
                            .intDelay)
            Call WriteToIni(strFilePath, _
                            strDayHeader, _
                            SET_DO_TIMED_SHUTDOWN, _
                            .blnDoShutdown)
            Call WriteToIni(strFilePath, _
                            strDayHeader, _
                            SET_AT_HOUR, _
                            .intHour)
            Call WriteToIni(strFilePath, _
                            strDayHeader, _
                            SET_AT_MINUTE, _
                            .intMinute)
            If blnNew Then
                Call WriteToIni(strFilePath, _
                                strDayHeader, _
                                SET_AT_SECOND, _
                                .intSecond & vbCrLf)
            Else
                Call WriteToIni(strFilePath, _
                                strDayHeader, _
                                SET_AT_SECOND, _
                                .intSecond)
            End If
        End With
    Next intCounter
    
    'colour settings
    Call WriteToIni(strFilePath, _
                    HDR_COLOUR, _
                    SET_COL_HOUR, _
                    glngColHour)
    Call WriteToIni(strFilePath, _
                    HDR_COLOUR, _
                    SET_COL_MINUTE, _
                    glngColMinute)
    Call WriteToIni(strFilePath, _
                    HDR_COLOUR, _
                    SET_COL_SECOND, _
                    glngColSecond)
    Call WriteToIni(strFilePath, _
                    HDR_COLOUR, _
                    SET_COL_DOTS, _
                    glngColDots)
    Call WriteToIni(strFilePath, _
                    HDR_COLOUR, _
                    SET_COL_ANA_BACK, _
                    glngColAnaBack)
    Call WriteToIni(strFilePath, _
                    HDR_COLOUR, _
                    SET_COL_TIME_FONT, _
                    glngColTimeFont)
    Call WriteToIni(strFilePath, _
                    HDR_COLOUR, _
                    SET_COL_TIME_BACK, _
                    glngColTimeBack)
    Call WriteToIni(strFilePath, _
                    HDR_COLOUR, _
                    SET_COL_DAY_FONT, _
                    glngColDayFont)
    Call WriteToIni(strFilePath, _
                    HDR_COLOUR, _
                    SET_COL_DAY_BACK, _
                    glngColDayBack)
    Call WriteToIni(strFilePath, _
                    HDR_COLOUR, _
                    SET_COL_DATE_FONT, _
                    glngColDateFont)
    Call WriteToIni(strFilePath, _
                    HDR_COLOUR, _
                    SET_COL_DATE_BACK, _
                    glngColDateBack)
    If blnNew Then
        Call WriteToIni(strFilePath, _
                        HDR_COLOUR, _
                        SET_COL_BORDER, _
                        glngColBorder & vbCrLf)
    Else
        Call WriteToIni(strFilePath, _
                        HDR_COLOUR, _
                        SET_COL_BORDER, _
                        glngColBorder)
    End If
    
    'background settings
    Call WriteToIni(strFilePath, _
                    HDR_BACKGROUND, _
                    SET_BACKGROUND_PATH, _
                    gstrBackPath)
    If blnNew Then
        Call WriteToIni(strFilePath, _
                        HDR_BACKGROUND, _
                        SET_BACKGROUND_MODE, _
                        genmBackStyle & vbCrLf)
    Else
        Call WriteToIni(strFilePath, _
                        HDR_BACKGROUND, _
                        SET_BACKGROUND_MODE, _
                        genmBackStyle)
    End If
    
    'miscellanious settings
    Call WriteToIni(strFilePath, _
                    HDR_MISC, _
                    SET_LAST_POS_X, _
                    gintLastPos(X))
    Call WriteToIni(strFilePath, _
                    HDR_MISC, _
                    SET_LAST_POS_Y, _
                    gintLastPos(Y))
    Call WriteToIni(strFilePath, _
                    HDR_MISC, _
                    SET_CLOCK_WIDTH, _
                    gintWidth)
    Call WriteToIni(strFilePath, _
                    HDR_MISC, _
                    SET_CLOCK_HEIGHT, _
                    gintHeight)
    Call WriteToIni(strFilePath, _
                    HDR_MISC, _
                    SET_IS_ON_TOP, _
                    gblnIsOnTop)
    Call WriteToIni(strFilePath, _
                    HDR_MISC, _
                    SET_START_MIN, _
                    gblnStartMin)
    Call WriteToIni(strFilePath, _
                    HDR_MISC, _
                    SET_START_HIDDEN, _
                    gblnStartHidden)
    Call WriteToIni(strFilePath, _
                    HDR_MISC, _
                    SET_SNAP_WINDOW, _
                    gblnSnapClock)
    Call WriteToIni(strFilePath, _
                    HDR_MISC, _
                    SET_RUN_AT_STARTUP, _
                    gblnRunAtStartup)
    Call WriteToIni(strFilePath, _
                    HDR_MISC, _
                    SET_TRANSPARENT, _
                    gsngTransLvl)
    Call WriteToIni(strFilePath, _
                    HDR_MISC, _
                    SET_BORDER_WIDTH, _
                    gintBorderWidth)
    
    If blnNew Then
        Call WriteToIni(strFilePath, _
                        HDR_MISC, _
                        SET_OWNER, _
                        gstrOwner & vbCrLf)
    Else
        Call WriteToIni(strFilePath, _
                        HDR_MISC, _
                        SET_OWNER, _
                        gstrOwner)
    End If
    
    'debug information
    Call WriteToIni(strFilePath, _
                    HDR_DEBUG, _
                    SET_LAST_SAVE_TIME, _
                    Time)
    Call WriteToIni(strFilePath, _
                    HDR_DEBUG, _
                    SET_LAST_SAVE_DATE, _
                    Date)
    Call WriteToIni(strFilePath, _
                    HDR_DEBUG, _
                    SET_SHUTDOWN_AT, _
                    GetShutdownTime)
    Call WriteToIni(strFilePath, _
                    HDR_DEBUG, _
                    SET_CURRENT_IDLE_TIME, _
                    CurrentIdleTime)
End Sub

Public Sub LoadSettings()
    'This procedure will load all the clocks settings
    
    Const STARTUP_KEY   As String = "Software\Microsoft\Windows\CurrentVersion\Run"
    
    Dim strFilePath     As String       'holds the path to the ini file
    Dim intCounter      As Integer      'used to cycle through the days of the week to get the settings from the ini file
    Dim strDayHeader    As String       'holds the header name for the given day
    Dim strValue        As String       'holds the registry value for the application starting up when windows starts
    
    'get the path of the ini file
    strFilePath = AddToPath(App.Path, INI_NAME)
    
    'See if the ini file exists. This won't cause an
    'error directly, but will affect the program
    'variables
    If (Dir(strFilePath) = "") Then
        'reset to default before exiting
        Call SetDefault
        Call SaveSettings
        Exit Sub
        
    ElseIf (FileLen(strFilePath) < 100) Then
        'the file size is way too small, reset
        Call SetDefault
        Call SaveSettings
        Exit Sub
    End If
    
    'load all the values
    
    'general settings
    genmShutdownMethod = Val(GetFromIni(strFilePath, _
                                        HDR_MAIN, _
                                        SET_SHUTDOWN_METHOD))
    gblnExclusiveShut = GetFromIni(strFilePath, _
                                   HDR_MAIN, _
                                   SET_PREVENT_OTHER_SHUTDOWN)
    gblnIdleShut = GetFromIni(strFilePath, _
                              HDR_MAIN, _
                              SET_DO_IDLE_SHUTDOWN)
    glngIdleTime = Val(GetFromIni(strFilePath, _
                                  HDR_MAIN, _
                                  SET_IDLE_TIME))
    gblnShowAnalog = GetFromIni(strFilePath, _
                                HDR_MAIN, _
                                SET_ANALOG)
    gbln24Hour = GetFromIni(strFilePath, _
                            HDR_MAIN, _
                            SET_24_HOUR)
    
    'password settings
    gstrPassword = DecryptText(GetFromIni(strFilePath, _
                                          HDR_PASSWORD, _
                                          SET_PASSWORD))
    gblnPassOn = GetFromIni(strFilePath, _
                            HDR_PASSWORD, _
                            SET_PASS_ACTIVE)
    
    'daily settings
    For intCounter = 0 To 6
        'set today's header
        strDayHeader = HDR_DAY & WeekdayName(intCounter + 1)

        With gudtDay(intCounter)
            .blnDoShutdown = GetFromIni(strFilePath, _
                                        strDayHeader, _
                                        SET_DO_TIMED_SHUTDOWN)
            .blnDoDelay = GetFromIni(strFilePath, _
                                     strDayHeader, _
                                     SET_DO_DELAY)
            .intDelay = Val(GetFromIni(strFilePath, _
                                       strDayHeader, _
                                       SET_DELAY_TIME))
            .intHour = Val(GetFromIni(strFilePath, _
                                      strDayHeader, _
                                      SET_AT_HOUR))
            .intMinute = Val(GetFromIni(strFilePath, _
                                        strDayHeader, _
                                        SET_AT_MINUTE))
            .intSecond = Val(GetFromIni(strFilePath, _
                                        strDayHeader, _
                                        SET_AT_SECOND))
        End With
    Next intCounter
    
    'colour settings
    glngColHour = Val(GetFromIni(strFilePath, _
                                 HDR_COLOUR, _
                                 SET_COL_HOUR))
    glngColMinute = Val(GetFromIni(strFilePath, _
                                   HDR_COLOUR, _
                                   SET_COL_MINUTE))
    glngColSecond = Val(GetFromIni(strFilePath, _
                                   HDR_COLOUR, _
                                   SET_COL_SECOND))
    glngColDots = Val(GetFromIni(strFilePath, _
                                 HDR_COLOUR, _
                                 SET_COL_DOTS))
    glngColAnaBack = Val(GetFromIni(strFilePath, _
                                    HDR_COLOUR, _
                                    SET_COL_ANA_BACK))
    glngColTimeFont = Val(GetFromIni(strFilePath, _
                                     HDR_COLOUR, _
                                     SET_COL_TIME_FONT))
    glngColTimeBack = Val(GetFromIni(strFilePath, _
                                     HDR_COLOUR, _
                                     SET_COL_TIME_BACK))
    glngColDayFont = Val(GetFromIni(strFilePath, _
                                    HDR_COLOUR, _
                                    SET_COL_DAY_FONT))
    glngColDayBack = Val(GetFromIni(strFilePath, _
                                    HDR_COLOUR, _
                                    SET_COL_DAY_BACK))
    glngColDateFont = Val(GetFromIni(strFilePath, _
                                     HDR_COLOUR, _
                                     SET_COL_DATE_FONT))
    glngColDateBack = Val(GetFromIni(strFilePath, _
                                     HDR_COLOUR, _
                                     SET_COL_DATE_BACK))
    glngColBorder = Val(GetFromIni(strFilePath, _
                                   HDR_COLOUR, _
                                   SET_COL_BORDER))
    
    'background settings
    gstrBackPath = GetFromIni(strFilePath, _
                              HDR_BACKGROUND, _
                              SET_BACKGROUND_PATH)
    genmBackStyle = Val(GetFromIni(strFilePath, _
                                   HDR_BACKGROUND, _
                                   SET_BACKGROUND_MODE))
    
    'miscellanious settings
    gintLastPos(X) = Val(GetFromIni(strFilePath, _
                                    HDR_MISC, _
                                    SET_LAST_POS_X))
    gintLastPos(Y) = Val(GetFromIni(strFilePath, _
                                    HDR_MISC, _
                                    SET_LAST_POS_Y))
    
    gintWidth = Val(GetFromIni(strFilePath, _
                               HDR_MISC, _
                               SET_CLOCK_WIDTH))
    
    gintHeight = Val(GetFromIni(strFilePath, _
                                HDR_MISC, _
                                SET_CLOCK_HEIGHT))
    
    gblnIsOnTop = GetFromIni(strFilePath, _
                             HDR_MISC, _
                             SET_IS_ON_TOP)
    gblnStartMin = GetFromIni(strFilePath, _
                              HDR_MISC, _
                              SET_START_MIN)
    gblnStartHidden = GetFromIni(strFilePath, _
                                 HDR_MISC, _
                                 SET_START_HIDDEN)
    gblnSnapClock = GetFromIni(strFilePath, _
                               HDR_MISC, _
                               SET_SNAP_WINDOW)
    gblnRunAtStartup = GetFromIni(strFilePath, _
                                  HDR_MISC, _
                                  SET_RUN_AT_STARTUP)
    gstrOwner = GetFromIni(strFilePath, _
                           HDR_MISC, _
                           SET_OWNER)
    gsngTransLvl = Val(GetFromIni(strFilePath, _
                                  HDR_MISC, _
                                  SET_TRANSPARENT))
    gintBorderWidth = Val(GetFromIni(strFilePath, _
                                     HDR_MISC, _
                                     SET_BORDER_WIDTH))
    
    
    'we need to confirm if the registry entry for starting
    'the clock exists and is correct
    strValue = ReadRegString(HKEY_LOCAL_MACHINE, _
                             STARTUP_KEY, _
                             STARTUP_LABEL)
    
    'is the program meant to run at startup
    If Not gblnRunAtStartup Then
        'make sure that the value is deleted if it exists
        If (strValue <> "") Then
            Call DeleteValue(HKEY_LOCAL_MACHINE, _
                             STARTUP_KEY, _
                             STARTUP_LABEL)
        End If
        
    Else
        'make sure that the value exists and is correct
        If (UCase(Trim(strValue)) <> UCase(Trim(AddToPath(App.Path, App.ExeName + ".exe")))) Then
            'create the  registry entry so that the program
            'starts when windows does
            Call PutAppInStartup(STARTUP_LABEL, blnOverwrite:=True)
        End If
    End If  'is the program meant to run at startup
End Sub

Public Sub SetAllSettings()
    'This will set all the settings for the global variables.
    'This procedure assumes that the form frmClock is already
    'loaded into memory as some of the code affects this form
    'directly.
    
    Dim intCounter      As Integer      'used to cycle through the clocks transparent menu array
    Dim intIndex        As Integer      'holds the menu index to set
    
    'should the program run at startup
    Call RunAtStartup
    
    With frmClock
        'set the transparent state of the form
        Call MakeWndTransparent(.hWnd, gsngTransLvl)
        
        'check the appropiate menu item based on the transparent amount
        Select Case gsngTransLvl
        Case Is < 0
            gsngTransLvl = 0
        
        Case Is > 100
            'we don't want to hide the form completely otherwise the
            'user will not be able to set it back to a more reasonable
            'level
            gsngTransLvl = 90
        End Select
        
        'round the value to the nearist 10 to get the proper menu item
        intIndex = gsngTransLvl \ 10
        
        For intCounter = .mnuClockAdvTransOpaque.LBound To .mnuClockAdvTransOpaque.UBound
            'set the menu checked item
            .mnuClockAdvTransOpaque(intCounter).Checked = (intIndex = intCounter)
        Next intCounter
        
        'enable/disable the Transparent menu based on the operating system
        .mnuClockAdvTrans.Enabled = IsW2000
        
        'set the border width
        If (gintBorderWidth < 0) Then
            gintBorderWidth = 0
        
        ElseIf (gintBorderWidth > 9) Then
            gintBorderWidth = 9
        End If
        
        'set the size and position of the clock
        Call .Move(gintLastPos(X), _
                   gintLastPos(Y), _
                   gintWidth, _
                   gintHeight)
        
        If gblnStartMin Then
            .WindowState = vbMinimized
        End If
        
        If gblnSnapClock Then
            'make sure that the window is visible
            Call CheckIfOutsideScreen(frmClock)
            
            'double check the co-ordinates
            gintLastPos(X) = .Left
            gintLastPos(Y) = .Top
        End If
        
        'set the clock settings directly
        With .mtscAnalog
            'colours
            .AnalogBackColour = glngColAnaBack
            .HandHourColour = glngColHour
            .HandMinuteColour = glngColMinute
            .HandSecondColour = glngColSecond
            .DotColour = glngColDots
            .DateBackColour = glngColDateBack
            .DateFontColour = glngColDateFont
            .DayBackColour = glngColDayBack
            .DayFontColour = glngColDayFont
            .TimeBackColour = glngColTimeBack
            .TimeFontColour = glngColTimeFont
            .BorderColour = glngColBorder
            
            'display settings
            .SurphaseDC = frmClock.hdc
            .Height = frmClock.ScaleHeight
            .Width = frmClock.ScaleWidth
            Set .Font = frmClock.Font
            Call .GetScreenPos(frmClock)
            .BackgroundStyle = genmBackStyle
            .PicturePath = gstrBackPath
            .Time24Hour = gbln24Hour
            .DisplayBackground = True
            .ShowAnalog = gblnShowAnalog
            .BorderWidth = gintBorderWidth
            
            'the current shutdown time
            .AlarmTime = GetShutdownTime
        End With
    End With
End Sub

Public Sub SetDefault()
    'resets all program variables to their default
    'values
    
    Const vbPurple      As Long = &HC000C0
    Const vbDarkPurple  As Long = &H800080
    Const vbLightRed    As Long = &H8080FF
    Const vbPink        As Long = &HFF00FF
    
    Dim intCounter  As Integer   'used to cycle through the daily settings
    
    'general values
    genmShutdownMethod = shtShutdown
    gblnExclusiveShut = False
    gblnIdleShut = False
    glngIdleTime = 3600 '1 hour
    gblnShowAnalog = True
    gbln24Hour = False
    
    'password values
    gstrPassword = ""
    gblnPassOn = False
    
    'daily settings
    For intCounter = 0 To 6
        With gudtDay(intCounter)
            .blnDoDelay = True
            .intDelay = 15  'seconds
            .blnDoShutdown = False
            .intHour = 0
            .intMinute = 0
            .intSecond = 0
        End With
    Next intCounter
    
    'colour settings
    glngColHour = vbPink
    glngColMinute = vbPurple
    glngColSecond = vbLightRed
    glngColAnaBack = vbYellow
    glngColDots = vbBlack
    glngColTimeFont = vbDarkPurple
    glngColTimeBack = vbYellow
    glngColDayFont = vbYellow
    glngColDayBack = vbPurple
    glngColDateFont = vbYellow
    glngColDateBack = vbPurple
    glngColBorder = vbBlack
    
    'background settings
    gstrBackPath = ""
    genmBackStyle = clkWallpaper
    
    'miscellanious settings
    gintLastPos(X) = Screen.Width
    gintLastPos(Y) = Screen.Height
    gintWidth = 1935
    gintHeight = 3060
    gblnIsOnTop = False
    gblnStartMin = False
    gblnStartHidden = False
    gblnSnapClock = True
    gblnRunAtStartup = True
    gsngTransLvl = 0    'fully opaque
    gintBorderWidth = 2
    
    'try and find the owner
    gstrOwner = GetRegisteredOwner
    
    'if unable to retrieve the licensed owner of this machine, then
    If UCase(Left(gstrOwner, 5)) = "ERROR" Then
        'reset to default
        gstrOwner = DEFAULT_OWNER
    End If
End Sub

Private Function GetTimeForDay(ByVal intDayNum As Integer) _
                               As String
    'returns the shutdown time for the specified day
    With gudtDay(intDayNum Mod 7)
        GetTimeForDay = .intHour & ":" & _
                        Format(.intMinute, "00") & ":" & _
                        Format(.intSecond, "00")
    End With
End Function

Public Function GetShutdownTime() As String
    'returns the closest shutdown time from the daily
    'settings and the timed shutdown
    
    Dim strShutTime As String   'holds the time to shutdown at
    Dim strIdleTime As String   'the predicted idle time
        
    'do we test tomorrows shutdowntime
    strShutTime = GetTimeForDay(Today)
    If DateDiff("s", strShutTime, Time) > 0 Then
        'the current shutdown time for today has
        'already passed. Test for tomorrows time
        strShutTime = DateAdd("d", 1, Date) & " " & _
                      GetTimeForDay(Today + 1)
    Else
        'make sure we are testing the right day
        strShutTime = Date & " " & strShutTime
    End If
    
    'do we account for the idle time
    If gblnIdleShut Then
        'get the predicted idle time from from now
        strIdleTime = DateAdd("s", _
                              glngIdleTime - (CurrentIdleTime / 1000), _
                              Now)
        
        'if the predicted idle time occurs AFTER the set
        'shutdown time, then ignore, otherwise set the
        'shutdown time for sooner
        If (DateDiff("s", Now, strShutTime) > _
            DateDiff("s", Now, strIdleTime)) Then
            
            'strip the date from the shutdown time
            '(a necessary addition when adding time
            'that goes over the date mark, 12:00)
            strIdleTime = Trim(Mid(strIdleTime, _
                                   InStr(strIdleTime, _
                                         " ")))
            
            'shutdown sooner
            GetShutdownTime = strIdleTime
            Exit Function
        End If
    End If
    
    'strip the date from the shutdown time
    '(a necessary addition when adding time
    'that goes over the date mark, 12:00)
    strShutTime = Trim(Mid(strShutTime, _
                           InStr(strShutTime, _
                                 " ")))
    
    'return today's shut down time
    GetShutdownTime = strShutTime
End Function

Public Function GetShutText(ByVal enmMethod As EnumShutdown) As String
    'return the text of the shut down method
    
    Select Case enmMethod
    Case shtLogOut
        GetShutText = "Log Off"
    
    Case shtForceClose
        GetShutText = "Force Close"
    
    Case shtShutdown
        GetShutText = "Shut Down"
    
    Case shtPowerDown
        GetShutText = "Power Down"
    
    Case shtRestart
        GetShutText = "Restart"
        
    Case shtLockWorkstation
        'applies to W2k only
        GetShutText = "Lock Computer"
    End Select
End Function

Public Sub DoShutMethod(Optional ByVal enmShutMethod As EnumShutdown = -1)
    'This will shut down the computer in the specified
    'method
    
    Dim lngIndex As Long    'the shutdown method
    
    'did the programmer specify a shutdown method?
    If enmShutMethod = -1 Then
        'default to shut down using the current method
        lngIndex = genmShutdownMethod
    Else
        'shut down using the specified method
        lngIndex = enmShutMethod
    End If
    
    'when in design mode, stop execution here
    Debug.Assert False
    
    Select Case lngIndex
    Case shtLogOut
        Call WINLogUserOff
    
    Case shtForceClose
        Call WINForceClose
    
    Case shtShutdown
        Call WINShutdown
    
    Case shtPowerDown
        Call WINPowerDown
    
    Case shtRestart
        Call WINReboot
    
    Case shtLockWorkstation
        Call WINLock
    End Select
End Sub

Public Function Today() As Integer
    'returns the array index for today
    Today = Weekday(Date, vbMonday) - 1
End Function

Public Function TicksToTime(ByVal lngTicks) As String
    'converts Ticks (milliseconds) to a usable time
    
    Dim intHours    As Integer
    Dim intMinutes  As Integer
    Dim intSeconds  As Integer
    
    'you can't have a negative time
    If lngTicks < 0 Then
        Exit Function
    End If
    
    'convert ticks to seconds
    lngTicks = lngTicks / 1000
    
    'split seconds into hh:mm:ss
    intSeconds = lngTicks Mod 60
    lngTicks = (lngTicks - intSeconds) / 60
    intMinutes = lngTicks Mod 60
    lngTicks = (lngTicks - intMinutes) / 60
    intHours = lngTicks
    
    'format the time
    TicksToTime = Format(intHours, "00") & ":" & _
                  Format(intMinutes, "00") & ":" & _
                  Format(intSeconds, "00")
End Function

Public Sub RunAtStartup()
    'This will adjust the registry enteries as necessary
    'for the program
    
    If gblnRunAtStartup Then
        'create a registry entry that will
        'start this program when windows starts
        Call PutAppInStartup(STARTUP_LABEL)
    Else
        Call RemoveAppFromStartup(STARTUP_LABEL)
    End If
End Sub

Public Sub CreateScheme(ByRef udtColours As TypeSchemes, _
                        Optional ByVal blnNewFile As Boolean = False)
    'This procedure will either add a new scheme to the
    'colour schemes file or create the file with the
    'default colour scheme (the udtColours parameter is
    'overwritten)
    
    'default colours
    Const vbPurple      As Long = &HC000C0
    Const vbDarkPurple  As Long = &H800080
    Const vbLightRed    As Long = &H8080FF
    Const vbPink        As Long = &HFF00FF
    
    'if a new file was specified, then create the file
    'using the default colours
    If blnNewFile Then
        'enter the default scheme
        With udtColours
            .strName = "[Default]"
            .lngBorderColour = vbBlack
            .lngAnalogBack = vbYellow
            .lngDots = vbBlack
            .lngHourHand = vbPink
            .lngMinuteHand = vbPurple
            .lngSecondHand = vbLightRed
            .lngTimeFont = vbDarkPurple
            .lngTimeBack = vbYellow
            .lngDayFont = vbYellow
            .lngDayBack = vbPurple
            .lngDateFont = vbYellow
            .lngDateBack = vbPurple
            .lngBorderColour = vbBlack
        End With
    End If
    
    'add the new scheme to the file
    Call AddRecord(gstrSchemePath, _
                   udtColours)
End Sub

Public Function DeleteScheme(ByVal lngRecordNum As Long) _
                             As Boolean
    'This will delete the specified record from the
    'colour schemes file. The function will only return
    'True if the record was deleted. The record number
    'specified is assumed to start from 0, not 1.
    
    Dim udtColours()    As TypeSchemes  'holds all the colour schemes in the file
    Dim intFileNum      As Integer      'holds a handle to the file
    Dim intCounter      As Integer      'used to remove the specified record from the array
    
    'get all data in the file
    If Not GetAllRecords(udtColours, gstrSchemePath) Then
        'Unable to load the schemes. Don't delete
        'anything
        DeleteScheme = False
        Exit Function
    End If
    
    'check for invalid record number
    If (lngRecordNum > UBound(udtColours)) Or _
       (lngRecordNum < 0) Then
        'the specified record doesn't exist
        DeleteScheme = False
        Exit Function
    End If
    
    'remove the specified record from the array is it is not the
    'last record (this is automatically removed)
    If lngRecordNum < UBound(udtColours) Then
        For intCounter = lngRecordNum To (UBound(udtColours) - 1)
            udtColours(intCounter) = udtColours(intCounter + 1)
        Next intCounter
    End If
    ReDim Preserve udtColours(UBound(udtColours) - 1)
    
    'set the error handler
    On Error GoTo ErrHandler
    
    'wipe all data in the file
    intFileNum = FreeFile
    Open gstrSchemePath For Output As #intFileNum
    Close #intFileNum
    
    'rewrite the file with only the correct data
    DeleteScheme = CreateFile(udtColours, _
                              gstrSchemePath)
    Exit Function
ErrHandler:
    On Error Resume Next
    'Close the file and exit the function
    Close #intFileNum
    
    DeleteScheme = False
    On Error GoTo 0
End Function

Public Sub AddRecord(ByVal strFilePath As String, _
                     ByRef udtColours As TypeSchemes, _
                     Optional ByVal lngRecordNum As Long = -1)
    'This procedure will add a record with the
    'information specified (ideally a user defined
    'type), and add it to the file. Any errors are
    'ignored.
    
    Dim intFileNum As Integer   'holds a handle to the file
    
    'reset the error handler to ignore errors
    On Error Resume Next
    
    'open the file for writing
    intFileNum = FreeFile
    Open strFilePath For Random As #intFileNum Len = Len(udtColours)
        'add the record to the file
        If lngRecordNum >= 0 Then
            'a record number was specified
            Put #intFileNum, lngRecordNum, udtColours
        Else
            'append the record to the end of the file
            Put #intFileNum, _
                (LOF(intFileNum) \ Len(udtColours)) + 1, _
                udtColours
        End If
    Close #intFileNum
    
    'set the error handler for normal handling
    On Error GoTo 0
End Sub

Public Function GetRecord(ByVal strFilePath As String, _
                          ByVal lngRecordNum As Long) _
                          As TypeSchemes
    'This procedure will get a record with the
    'information specified (ideally a user defined
    'type), and add it to the file. Any errors are
    'ignored.
    
    Dim intFileNum As Integer       'holds a handle to the file
    Dim udtColours As TypeSchemes   'holds the colour scheme specified
    
    'reset the error handler to ignore errors
    On Error Resume Next
    
    'open the file for writing
    intFileNum = FreeFile
    Open strFilePath For Random As #intFileNum Len = Len(udtColours)
        'get the record from the file
        Get #intFileNum, lngRecordNum, udtColours
    Close #intFileNum
    
    'return the record before exiting
    GetRecord = udtColours
    
    'set the error handler for normal handling
    On Error GoTo 0
End Function

Private Function CreateFile(ByRef udtColours() As TypeSchemes, _
                            ByVal strFilePath As String) _
                            As Boolean
    'This function will create a random file based on
    'the data in the array. If the file was created
    'successfully, then the function will return True.
    'A one dimensional array is assumed
    
    Dim intFileNum  As Integer  'holds a handler to the file
    Dim intLBound   As Integer  'the lower bound of the array
    Dim intUBound   As Integer  'the upper bound of the array
    Dim intRec      As Integer  'the current record number
    Dim intCounter  As Integer  'used to cycle through the array
    
    'set the error handler
    On Error GoTo ErrHandler
    
    'get the array bounds
    intLBound = LBound(udtColours)
    intUBound = UBound(udtColours)
    
    'create the file and enter the records
    intFileNum = FreeFile
    Open strFilePath For Random As #intFileNum Len = Len(udtColours(LBound(udtColours)))
        For intCounter = intLBound To intUBound
            'get a valid record number (numbers start
            'from 1 and increate incrementally)
            intRec = (intCounter - intLBound) + 1
            
            'create the record
            Put #intFileNum, intRec, udtColours(intCounter)
        Next intCounter
    Close #intFileNum
    
    'return True and exit
    CreateFile = True
    Exit Function
    
ErrHandler:
    'return False and exit
    CreateFile = False
    Exit Function
End Function

Public Function GetAllRecords(ByRef udtColours() As TypeSchemes, _
                              ByVal strFilePath As String) _
                              As Boolean
    'This function will return all contents based on
    'the record type specified. If the file was read
    'successfully, then the function will return True.
    'A one dimensional dynamic array is assumed
    
    Dim intFileNum As Integer       'holds a handler to the file
    Dim intCounter As Integer       'used to help enter records into the array
    
    'set the error handler
    On Error GoTo ErrHandler
    
    'intitialise the array
    intCounter = 0
    ReDim udtColours(intCounter)
    
    'open the file and read the records
    intFileNum = FreeFile
    Open strFilePath For Random As #intFileNum Len = Len(udtColours(LBound(udtColours)))
        Do While Not EOF(intFileNum)
            'read a record
            Get #intFileNum, (intCounter + 1), udtColours(intCounter)
            
            'we only enter the scheme into the array
            'if it has a valid name
            If udtColours(intCounter).strName <> String(50, vbNullChar) Then
                'enter a new element
                intCounter = intCounter + 1
            
                'resize the array to hold a new record
                ReDim Preserve udtColours(intCounter)
            End If
        Loop
    Close #intFileNum
    
    'The last array element is always blank. Remove it from the
    'array (if at least one valid record was found)
    If UBound(udtColours) > 0 Then
        ReDim Preserve udtColours(UBound(udtColours) - 1)
    End If
    
    'return True and exit
    GetAllRecords = True
    Exit Function
    
ErrHandler:
    'return False and exit
    GetAllRecords = False
    Exit Function
End Function

Private Function GetArrayDimensions(ByRef varArray() As TypeSchemes) _
                                    As Integer
    'Get the number of dimensions in an array. The
    'function will return 0 if the array has not been
    'initialised
    
    Const MAX_DIMENSIONS    As Integer = 60  'this is the maximum number of dimensions that vb allows for an array
    
    Dim intNum      As Integer  'holds the number of dimensions
    Dim intCounter  As Integer  'used to test for each dimension
    Dim lngErrNum   As Long     'holds an error number. I use this to stop testing array dimensions
    
    'set the error handler for this function
    On Error Resume Next
    
    'scan through each dimension
    For intCounter = 0 To MAX_DIMENSIONS
        'test the dimension
        intNum = LBound(varArray, intCounter)
        
        'check if an error occured
        lngErrNum = Err
        
        'If no error occurred, then the dimension
        'is valid. Continue
        If lngErrNum > 0 Then
            'return the number of dimensions
            GetArrayDimensions = intCounter
            Exit Function
        End If
    Next intCounter
    
    'the array is at maximum size
    GetArrayDimensions = MAX_DIMENSIONS
End Function

Public Sub HighLight(ByRef txtBox As TextBox)
    'This will try and highlight the text in a text box.
    
    'make sure that if this was called from a Form_Load
    'event (or similar), that it does not cause the
    'program to crash.
    On Error Resume Next
        With txtBox
            'highlight the text
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    On Error GoTo 0
End Sub

Public Sub UnloadAll(Optional ByVal frmLast As Form)
    'This will unload all the forms in the project, with the specified
    'form being unloaded last
    
    Dim frm As Form
    
    'make sure that all open file handles are closed
    Call Reset
    
    For Each frm In Forms
        If Not frmLast Is Nothing Then
            'make sure we dont unload the specified form until last
            If frmLast.Name <> frm.Name Then
                Unload frm
                Set frm = Nothing
            End If
        Else
            'unload every form
            Unload frm
            Set frm = Nothing
        End If
    Next frm
    
    'unload the last form if one was specified
    If Not frmLast Is Nothing Then
        Unload frmLast
        Set frmLast = Nothing
    End If
End Sub

Public Function InDebug() As Boolean
    'This will return True only while the program is being run from
    'the vb ide. This works because the compiler automatically
    'removes all Debug statements during compile. We can automatically
    'generate an error on a Debug line, and testing to see if the
    'error occured. The error will only occur when the program is
    'being run from the ide.
    '
    'NOTE: you might want to make sure that the option in the ide
    '      to "Break on all errors" is turned off. Go to "Tools",
    '      "Options" and click on the "General" tab to set this.
    
    'reset the error handler
    Call Err.Clear
    On Local Error Resume Next
        Debug.Assert (1 / 0)
        
        'return the result of this error
        InDebug = (Err.Number <> 0)
        
        'reset back to normal error handling
        Call Err.Clear
    On Local Error GoTo 0
End Function

Public Sub DoCleanUp(ByRef frmExcept As Form)
    'this will check through all forms in the project and unload them except for the form passed through
    'the parameter which is used to keep the program active.
    
    Dim frmCounter      As Form         'used to cycle through all forms in the project
    
    'no exception form was passed - we cannot run through this procedure as it will cause the program to end
    If (frmExcept Is Nothing) Then 'Or (Not gblnDoCleanUp) Then
        Exit Sub
    End If
    
    For Each frmCounter In Forms
        
        'is the form loaded
        If Not frmCounter Is Nothing Then
            
            'is this form the exception form
            If (frmCounter.Name <> frmExcept.Name) Then
                
                'is the form still visible to the user
                If Not frmCounter.Visible Then
                    Unload frmCounter
                    Set frmCounter = Nothing
                End If  'is the form still visible to the user
            End If  'is this the exception form
        End If  'is the form unloaded
    Next frmCounter
    
    'we have cleared memory as much as possible
    gblnDoCleanUp = False
End Sub
