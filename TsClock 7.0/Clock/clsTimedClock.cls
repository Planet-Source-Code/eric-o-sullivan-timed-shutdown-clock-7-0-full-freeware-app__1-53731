VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimedClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This will holds four classes of the type clsClock to control a complete clock application."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     2 August 2002
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    Timed Shutdown Clock Class
' -----------------------------------------------
'COMMENTS :
'This clock will display an analog clock, a digital
'clock, the day and the date, vertically, with the
'appropiate settings
'=================================================

'require varaible declaration
Option Explicit

'------------------------------------------------
'                     EVENTS
'------------------------------------------------
'This event fires when the time has changed
Public Event NewTime(ByVal NewTime As Date)

'This triggers when the a new day occurs (when the date
'changes)
Public Event NewDay(ByVal NewDate As String)
Attribute NewDay.VB_Description = "A new day has occurred"

'This will trigger when the alarm is set
Public Event AlarmActivate()
Attribute AlarmActivate.VB_Description = "Triggers when the alarm time matches the time properties"

'------------------------------------------------
'             MODULE-LEVEL VARIABLES
'------------------------------------------------
'local variable(s) to hold property value(s)
Private mstrPicturePath     As String               'holds the file path to the picture specified
Private menmBackgroundStyle As EnmCBackgroundStyle  'how to display the background picture
Private mintTop             As Integer              'the vertical position of the top of the clock
Private mintLeft            As Integer              'the horizontal position of the left of the clock
Private mintWidth           As Integer              'the width of the clock
Private mintHeight          As Integer              'the total height of the clock
Private mblnVisible         As Boolean              'is the clock visible
Private mblnShowAnalog      As Boolean              'whether or not to show the analog clock
Private mintBorderWidth     As Integer              'holds the size of the border seperating the different panels
Private mlngBorderColour    As Long                 'holds the colour for the border
Private mlngSurphaseDc      As Long                 'holds the location of the area to paint onto
Private mintAnaHeight       As Integer              'the new height of the analog clock
Private mintScreenX         As Integer              'holds the position left side of the clock in screen co-ordinates
Private mintScreenY         As Integer              'holds the position top side of the clock in screen co-ordinates

'clock objects that this class will encapsulate
Private WithEvents mclkTime As clsClock
Attribute mclkTime.VB_VarHelpID = -1
Private mclkAnalog          As clsClock
Private mclkDay             As clsClock
Private mclkDate            As clsClock

'used to build the background picture
Private mudtBackground      As clsBitmap
Private mudtBackPic         As clsBitmap

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Public Property Get AnalogHeight() As Integer
    'Returns the current size of the analog clock. This
    'property is Read-Only
    AnalogHeight = mintAnaHeight
End Property

Public Property Let PicturePath(ByVal strPath As String)
    'This will set the path to the specified picture
    
    'validate the path
    If (Dir(strPath) = "") Or (Trim(strPath) = "") Then
        'invalid path, don't load picture and clear
        'any existing picture and show background
        'colours instead
        
        'if we are using a background style that
        'uses the picture, then we must reset the
        'background style
        If Me.BackgroundStyle <> clkWallpaper Then
            Me.BackgroundStyle = clkNone
        End If
        Exit Property
    End If
    
    'load the new picture
    mstrPicturePath = strPath
    Call BuildBackground
    Call LoadBackgroundPicture
    Call Refresh
End Property

Public Property Get PicturePath() As String
    'This will return the path to the specified picture
    PicturePath = mstrPicturePath
End Property

Public Property Let BorderColour(ByVal lngColour As Long)
    'sets the new colour for the border
    mlngBorderColour = lngColour
    
    'update the display
    Call BuildBackground
    Call Refresh
End Property

Public Property Get BorderColour() As Long
    'This will return the colour of the border along the edge of the clock
    BorderColour = mlngBorderColour
End Property

Public Property Let BorderWidth(ByVal intWidth As Integer)
    'Sets the width of the border seperating the
    'different panels on the clock. This proecdure
    'will accept a border width between 0 and 9
    
    'set the valid range for the border width
    Const BORDER_MIN    As Integer = 0
    Const BORDER_MAX    As Integer = 9
    
    'validate the width of the border
    If intWidth < BORDER_MIN Then
        intWidth = BORDER_MIN
    Else
        If intWidth > BORDER_MAX Then
            intWidth = BORDER_MAX
        End If
    End If
    
    'set the new size
    mintBorderWidth = intWidth
    
    'update the new display
    Call ReSize(mintWidth, mintHeight)
End Property

Property Get BorderWidth() As Integer
    'returns the size of the border
    BorderWidth = mintBorderWidth
End Property

Public Property Let DateFontColour(ByVal vData As Long)
    'the font colour of the date panel
    mclkDate.ForeColour = vData
End Property

Public Property Get DateFontColour() As Long
    'the font colour of the date panel
    DateFontColour = mclkDate.ForeColour
End Property

Public Property Let DateBackColour(ByVal vData As Long)
    'the background colour of the date panel
    mclkDate.BackgroundColour = vData
End Property

Public Property Get DateBackColour() As Long
    'the background colour of the date panel
    DateBackColour = mclkDate.BackgroundColour
End Property

Public Property Let DayFontColour(ByVal vData As Long)
    'the font colour of the day panel
    mclkDay.ForeColour = vData
End Property

Public Property Get DayFontColour() As Long
    'the font colour of the day panel
    DayFontColour = mclkDay.ForeColour
End Property

Public Property Let DayBackColour(ByVal vData As Long)
Attribute DayBackColour.VB_Description = "The background colour of the day panel"
    'the background colour of the day panel
    mclkDay.BackgroundColour = vData
End Property

Public Property Get DayBackColour() As Long
    'the bakground colour of the day panel
    DayBackColour = mclkDay.BackgroundColour
End Property

Public Property Let TimeFontColour(ByVal vData As Long)
Attribute TimeFontColour.VB_Description = "The colour of the time text"
    'the font colour of the time panel
    mclkTime.ForeColour = vData
End Property

Public Property Get TimeFontColour() As Long
    'the font colour of the time panel
    TimeFontColour = mclkTime.ForeColour
End Property

Public Property Let TimeBackColour(ByVal vData As Long)
Attribute TimeBackColour.VB_Description = "The background colour of the time panel"
    'the background colour of the time panel
    mclkTime.BackgroundColour = vData
End Property

Public Property Get TimeBackColour() As Long
    'the background colour of the time panel
    TimeBackColour = mclkTime.BackgroundColour
End Property

Public Property Let AnalogBackColour(ByVal vData As Long)
Attribute AnalogBackColour.VB_Description = "The background of the analog clock"
    'the bakground colour of the analog clock
    mclkAnalog.BackgroundColour = vData
End Property

Public Property Get AnalogBackColour() As Long
    'the background colour of the analog clock
    AnalogBackColour = mclkAnalog.BackgroundColour
End Property

Public Property Let Time24Hour(ByVal vData As Boolean)
Attribute Time24Hour.VB_Description = "Whether or not to display the time in 24 hour format or 12 hour format"
    'should the time being displayed, be in 24 hour
    'format
    mclkTime.Time24Hour = vData
End Property

Public Property Get Time24Hour() As Boolean
    'is the time beign displayed in 24 hour format
    Time24Hour = mclkTime.Time24Hour
End Property

Public Property Let SurphaseDC(ByVal vData As Long)
Attribute SurphaseDC.VB_Description = "This is the surphase that we are supposed to copy to"
    'set a handle to the surphase picture
    mlngSurphaseDc = vData
    Call BuildBackground
    Call Refresh
End Property

Public Property Get SurphaseDC() As Long
    'return a handle to the surphase picture
    SurphaseDC = mlngSurphaseDc
End Property

Public Property Set Font(ByVal fntNewFont As StdFont)
Attribute Font.VB_Description = "The font settings of any text displayed"
    'set the font being used
    Set mclkTime.Font = fntNewFont
    Set mclkDay.Font = fntNewFont
    Set mclkDate.Font = fntNewFont
End Property

Public Property Get Font() As StdFont
    'return the font being used to display the time
    Set Font = mclkTime.Font
End Property

Public Property Let ShowAnalog(ByVal blnShow As Boolean)
    'should we display the analog clock
    mblnShowAnalog = blnShow
    
    If mblnShowAnalog And mblnVisible Then
        'display the analog clock
        mclkAnalog.Visible = True
    Else
        mclkAnalog.Visible = False
    End If
    
    Call SetDimensions(mintHeight)
    Call BuildBackground
    Call Refresh
End Property

Public Property Get ShowAnalog() As Boolean
    'is the analog clock being displayed
    ShowAnalog = mblnShowAnalog
End Property

Public Property Let Visible(ByVal vData As Boolean)
Attribute Visible.VB_Description = "Returns or sets whether or not the class has a visual interface."
    'Should the clock be visible on the surphase
    'picture
    mblnVisible = vData
    
    If mblnShowAnalog And mblnVisible Then
        'display the analog clock
        mclkAnalog.Visible = True
    Else
        mclkAnalog.Visible = False
    End If
    
    mclkTime.Visible = mblnVisible
    mclkDay.Visible = mblnVisible
    mclkDate.Visible = mblnVisible
    
    'make sure the clock is shown
    Call PaintClock
End Property

Public Property Get Visible() As Boolean
    'should the clock be visible on the surphase
    'picture
    Visible = mblnVisible
End Property

Public Sub Cls()
Attribute Cls.VB_Description = "This wipes the clock area on the Background with the Background colour"
    'clear all graphical data on the clocks
    mclkAnalog.Cls
    mclkTime.Cls
    mclkDay.Cls
    mclkDate.Cls
End Sub

Public Property Let DotColour(ByVal vData As Long)
Attribute DotColour.VB_Description = "The colour of the dots displayed when the clock style is set to Analog"
    'The colour of the dots
    mclkAnalog.DotColour = vData
End Property

Public Property Get DotColour() As Long
    'The colour of the dots
    DotColour = mclkAnalog.DotColour
End Property

Public Property Let AlarmTime(ByVal vData As String)
Attribute AlarmTime.VB_Description = "This will trigger the Alarm event when the specified time matches the AlarmTime. Only active when the AlarmTime is a valid time format, otherwise the Alarm event does NOT trigger."
    'The time the alarm should trigger at
    mclkTime.AlarmTime = vData
End Property

Public Property Get AlarmTime() As String
    'The time the alarm event should trigger at
    AlarmTime = mclkTime.AlarmTime
End Property

Public Property Let HandSecondColour(ByVal vData As Long)
Attribute HandSecondColour.VB_Description = "The colour of the second hand on the clock"
    'The colour of the second hand
    mclkAnalog.HandSecondColour = vData
End Property

Public Property Get HandSecondColour() As Long
    'The colour of the second hand
    HandSecondColour = mclkAnalog.HandSecondColour
End Property

Public Property Let HandMinuteColour(ByVal vData As Long)
Attribute HandMinuteColour.VB_Description = "The colour of the minute hand on the clock"
    'The colour of the minute  hand
    mclkAnalog.HandMinuteColour = vData
End Property

Public Property Get HandMinuteColour() As Long
    'The colour of the minute hand
    HandMinuteColour = mclkAnalog.HandMinuteColour
End Property

Public Property Let HandHourColour(ByVal vData As Long)
Attribute HandHourColour.VB_Description = "The colour of the hour hand on the analog clock"
    'The colour of the hour hand
    mclkAnalog.HandHourColour = vData
End Property

Public Property Get HandHourColour() As Long
    'The clour of the hour hand
    HandHourColour = mclkAnalog.HandHourColour
End Property

Public Property Let Left(ByVal vData As Integer)
Attribute Left.VB_Description = "Returns or sets the horizontal position of the left side of the clock"
    'the X co-ordiante of the position of the clock
    mintLeft = vData
End Property

Public Property Get Left() As Integer
    'the X co-ordinate of the position of the clock
    Left = mintLeft
End Property

Public Property Let Top(ByVal vData As Integer)
Attribute Top.VB_Description = "Returns or sets the vertical position of the top of the class"
    'the Y co-ordinate of the position of the clock
    mintTop = vData
End Property

Public Property Get Top() As Integer
    'the Y co-ordinate of the position of the clock
    Top = mintTop
End Property

Public Property Let DisplayBackground(ByVal vData As Boolean)
    'should we display the background picture
    mclkAnalog.DisplayBackground = vData
    mclkTime.DisplayBackground = vData
    mclkDay.DisplayBackground = vData
    mclkDate.DisplayBackground = vData
    
    'make sure that the clock will display the graphic
    'we create as normal for that particular panel
    If vData Then
        mclkAnalog.BackgroundStyle = clkNormal
        mclkTime.BackgroundStyle = clkNormal
        mclkDay.BackgroundStyle = clkNormal
        mclkDate.BackgroundStyle = clkNormal
    Else
        mclkAnalog.BackgroundStyle = clkNone
        mclkTime.BackgroundStyle = clkNone
        mclkDay.BackgroundStyle = clkNone
        mclkDate.BackgroundStyle = clkNone
    End If
End Property

Public Property Get DisplayBackground() As Boolean
    'should we display the background picture
    DisplayBackground = mclkTime.DisplayBackground
End Property

Public Property Let BackgroundStyle(ByVal vData As EnmCBackgroundStyle)
    'The background style of the clock
    menmBackgroundStyle = vData
    
    Call BuildBackground
    Call LoadBackgroundPicture
    Call Refresh
End Property

Public Property Get BackgroundStyle() As EnmCBackgroundStyle
    'the background style of the clock
    BackgroundStyle = menmBackgroundStyle
End Property

Public Property Let Height(ByVal vData As Integer)
Attribute Height.VB_Description = "The height of the area to display"
    'The height of the clock
    mintHeight = vData
    
    'adjust the clock to the new height
    Call SetDimensions(vData)
    
    'update the display
    Call BuildBackground
    Call Refresh
End Property

Public Property Get Height() As Integer
    'The height of the clock
    Height = mintHeight
End Property

Public Property Let Width(ByVal intNewWidth As Integer)
Attribute Width.VB_Description = "The width of the area to display"
    'The width of the clock
    mintWidth = intNewWidth
    mclkAnalog.ClockWidth = intNewWidth - mintBorderWidth
    mclkTime.ClockWidth = intNewWidth - mintBorderWidth
    mclkDay.ClockWidth = intNewWidth - mintBorderWidth
    mclkDate.ClockWidth = intNewWidth - mintBorderWidth
    
    'update the display
    Call SetDimensions(mintHeight)
    Call BuildBackground
    Call Refresh
End Property

Public Property Get Width() As Integer
    'the width of the clock
    Width = mintWidth
End Property

Public Property Let AutoDisplay(ByVal blnEnabled As Boolean)
Attribute AutoDisplay.VB_Description = "Returns or set whether or not the class is to automatically update the display when the time changes"
    'should the display be automatically updated
    mclkAnalog.AutoDisplay = blnEnabled
    mclkTime.AutoDisplay = blnEnabled
    mclkDay.AutoDisplay = blnEnabled
    mclkDate.AutoDisplay = blnEnabled
End Property

Public Property Get AutoDisplay() As Boolean
    'should the display be automatically updated
    AutoDisplay = mclkTime.AutoDisplay
End Property

Public Property Let AutoUpdateData(ByVal blnEnabled As Boolean)
Attribute AutoUpdateData.VB_Description = "Returns or sets whether or not to get the settings from the system clock and to update the properties when they change"
    ''should the time and date be automatically updated
    mclkAnalog.AutoUpdateData = blnEnabled
    mclkTime.AutoUpdateData = blnEnabled
    mclkDay.AutoUpdateData = blnEnabled
    mclkDate.AutoUpdateData = blnEnabled
End Property

Public Property Get AutoUpdateData() As Boolean
    'should the time and date be automatically updated
    AutoUpdateData = mclkTime.AutoUpdateData
End Property

Public Property Let ClockTime(ByVal vData As String)
Attribute ClockTime.VB_Description = "Returns or sets the time"
    'sets the current time of the clock
    mclkTime.ClockTime = vData
End Property

Public Property Get ClockTime() As String
    'returns the current time of the clock
    ClockTime = mclkTime.ClockTime
End Property

Public Property Let ScreenPosX(ByVal intX As Integer)
    'This will set the X location of the left side of the clock
    'in pixels, relative to the screen dimensions. This is only
    'used when the BackgroundStyle property is set to clkWallpaper
    mintScreenX = intX
End Property

Public Property Get ScreenPosX() As Integer
    'This will return the X location of the left side of the clock
    'in pixels, relative to the screen dimensions. This is only
    'used when the BackgroundStyle property is set to clkWallpaper
    ScreenPosX = mintScreenX
End Property

Public Property Let ScreenPosY(ByVal intY As Integer)
    'This will set the Y location of the top side of the clock
    'in pixels, relative to the screen dimensions. This is only
    'used when the BackgroundStyle property is set to clkWallpaper
    mintScreenY = intY
End Property

Public Property Get ScreenPosY() As Integer
    'This will return the Y location of the top side of the clock
    'in pixels, relative to the screen dimensions. This is only
    'used when the BackgroundStyle property is set to clkWallpaper
    ScreenPosY = mintScreenY
End Property

Public Sub Refresh()
    'refresh all the clocks
    mclkAnalog.Refresh
    mclkTime.Refresh
    mclkDate.Refresh
    mclkDay.Refresh
    
    'copy the contents of the bitmap onto the source
    'Dc
    'check if we need to display the clock
'    If mblnVisible Then
'        Call mudtBackground.Paint(mlngSurphaseDc, _
'                                  mintLeft, _
'                                  mintTop, _
'                                  mintHeight, _
'                                  mintWidth)
'    End If
End Sub

Public Sub PaintClock()
    'display the clock on the surphase
    
    'repaint all the clocks
    mclkAnalog.PaintClock
    mclkTime.PaintClock
    mclkDay.PaintClock
    mclkDate.PaintClock
    
    'check if we need to display the clock
    If (mblnVisible) Or (mlngSurphaseDc = 0) Then
        'copy the contents of the bitmap onto the source
        'Dc
        Call mudtBackground.Paint(mlngSurphaseDc, _
                                  mintLeft, _
                                  mintTop, _
                                  mintHeight, _
                                  mintWidth)
    End If
End Sub

Private Sub Class_Initialize()
    'create instances of the different clocks and
    'set their initial positon and default values
    
    Const DEFAULT_HEIGHT    As Integer = 180
    Const DEFAULT_WIDTH     As Integer = 117
    
    'create the background surphase and the
    'background picture
    Set mudtBackground = New clsBitmap
    Set mudtBackPic = New clsBitmap
    
    'create the clocks
    Set mclkAnalog = New clsClock
    Set mclkTime = New clsClock
    Set mclkDay = New clsClock
    Set mclkDate = New clsClock
    
    'get any default values from the clsClock classes
    mclkAnalog.DisplayType = clkAnalog
    mclkDay.DisplayType = clkDay
    mclkDate.DisplayType = clkDate
    
    mintHeight = DEFAULT_HEIGHT
    mintWidth = DEFAULT_WIDTH
    
    With mclkTime
            Me.BackgroundStyle = clkNormal
            Me.BorderColour = vbBlack
            Me.BorderWidth = 2  'pixel
            Me.Top = 0
            Me.Left = 0
            Me.Width = DEFAULT_WIDTH
            Me.Visible = False
            Me.SurphaseDC = 0
            Me.AnalogBackColour = .BackgroundColour
            Me.DotColour = .DotColour
            Me.TimeBackColour = .BackgroundColour
            Me.TimeFontColour = .ForeColour
            Me.DayBackColour = .BackgroundColour
            Me.DayFontColour = .ForeColour
            Me.DateBackColour = .BackgroundColour
            Me.DateFontColour = .ForeColour
            Me.ShowAnalog = True
    End With
    
    'Set the heights for each panel
    Call SetDimensions(DEFAULT_HEIGHT)
    Call BuildBackground
End Sub

Private Sub Class_Terminate()
    'make sure the class removes all data appropiatly
    Set mclkAnalog = Nothing
    Set mclkTime = Nothing
    Set mclkDay = Nothing
    Set mclkDate = Nothing
End Sub

Private Sub mclkTime_AlarmActivate()
    'trigger the appropiate event
    RaiseEvent AlarmActivate
End Sub

Private Sub mclkTime_NewDay(ByVal NewDate As Date)
    'update display changes
    mclkDay.Refresh
    mclkDate.Refresh
    
    'trigger the appropiate event
    RaiseEvent NewDay(NewDate)
End Sub

Private Sub mclkTime_NewTime(ByVal NewTime As Date)
    'trigger the NewTime event for this class
    RaiseEvent NewTime(NewTime)
End Sub

Private Sub BuildBackground()
    'This will rebuild the background from scratch
    'based on the current settings
    
    'create the background bitmap
    With mudtBackground
        Call .SetBitmap(mintWidth, _
                        mintHeight, _
                        mlngBorderColour)
        
        'do we show a background or not
        If menmBackgroundStyle = clkNone Then
            'no background
            mclkAnalog.BackgroundStyle = clkNone
            mclkTime.BackgroundStyle = clkNone
            mclkDay.BackgroundStyle = clkNone
            mclkDate.BackgroundStyle = clkNone
        Else
            'display a background
            mclkAnalog.BackgroundStyle = clkNormal
            mclkTime.BackgroundStyle = clkNormal
            mclkDay.BackgroundStyle = clkNormal
            mclkDate.BackgroundStyle = clkNormal
        End If
        
        'set the surphase of the clocks to this bitmap
        mclkAnalog.SurphaseDC = .hdc
        mclkTime.SurphaseDC = .hdc
        mclkDay.SurphaseDC = .hdc
        mclkDate.SurphaseDC = .hdc
    End With
End Sub

Public Sub ReSize(ByVal intWidth As Integer, _
                  ByVal intHeight As Integer)
    'This will resize the clock and reload the
    'background picture if any was specified
    
    mintWidth = intWidth
    mintHeight = intHeight
    
    'adjust the clock to the new height
    Call SetDimensions(intHeight)
    
    'update the display
    Call BuildBackground
    Call LoadBackgroundPicture
    Call Refresh
End Sub

Private Sub SetDimensions(ByVal intHeight As Integer)
    'The analog clock will take up 2/3 of the total
    'height, the remaining space will be broken up
    'into 3 sections, one for each remaining panel.
    'This procedure will also adjust the appropiate
    'heights for the panels
    
    Dim intPanelHeight As Integer   'the height of the remaining panels
    
    'set the left and width positions according to the
    'border width
    mclkAnalog.ClockLeft = mintBorderWidth
    mclkTime.ClockLeft = mintBorderWidth
    mclkDay.ClockLeft = mintBorderWidth
    mclkDate.ClockLeft = mintBorderWidth
    mclkAnalog.ClockWidth = mintWidth - (mintBorderWidth * 2)
    mclkTime.ClockWidth = mintWidth - (mintBorderWidth * 2)
    mclkDay.ClockWidth = mintWidth - (mintBorderWidth * 2)
    mclkDate.ClockWidth = mintWidth - (mintBorderWidth * 2)
    
    'get the height of the analog clock and
    'the remaining panels
    intHeight = intHeight - mintBorderWidth
    If mblnShowAnalog Then
        'size for analog
        mintAnaHeight = Round((intHeight / 3) * 2)
        intPanelHeight = Round(intHeight / 9)
    
    Else
        'size for digital display only
        mintAnaHeight = Round((intHeight / 3) * 2)
        intPanelHeight = Round(intHeight / 3)
    End If
    
    'set the heights for the panels (adjust for
    'borders)
    mclkAnalog.ClockHeight = mintAnaHeight - mintBorderWidth
    mclkTime.ClockHeight = intPanelHeight - mintBorderWidth
    mclkDay.ClockHeight = intPanelHeight - mintBorderWidth
    
    'set the appropiate positions for the
    'panels depending on whether or not the
    'analog clock should be displayed
    If Not mblnShowAnalog Then
        'adjust the time panel to display at the
        'top instead of after the analog
        mclkTime.ClockTop = mintBorderWidth
    Else
        'display the time panel after the analog
        'clock
        mclkAnalog.ClockTop = mintBorderWidth
        mclkTime.ClockTop = mintAnaHeight + mintBorderWidth
    End If
    
    'the remaining panels are calculated in
    'position relative to the time panel, so
    'these do not need to be accounted for,
    '(ie, the panels are "below" the time panel)
    mclkDay.ClockTop = mclkTime.ClockTop + _
                       mclkTime.ClockHeight + _
                       mintBorderWidth
    mclkDate.ClockTop = mclkDay.ClockTop + _
                        mclkDay.ClockHeight + _
                        mintBorderWidth
    mclkDate.ClockHeight = (intHeight - mclkDate.ClockTop)
    
    'set the co-ordinates for the background picture
    mclkAnalog.BackgroundY = mclkAnalog.ClockTop
    mclkTime.BackgroundY = mclkTime.ClockTop
    mclkDay.BackgroundY = mclkDay.ClockTop
    mclkDate.BackgroundY = mclkDate.ClockTop
    
    mclkAnalog.BackgroundX = mclkAnalog.ClockLeft
    mclkTime.BackgroundX = mclkTime.ClockLeft
    mclkDay.BackgroundX = mclkDay.ClockLeft
    mclkDate.BackgroundX = mclkDate.ClockLeft
End Sub

Private Sub LoadBackgroundPicture()
    'This will load the background picture into an
    'off-screen dc.
    
    Dim udtTemp As clsBitmap    'temperorily holds the background picture
    
    'load the new picture
    With mudtBackPic
        'create a new bitmap of the appropiate
        'size
        Set udtTemp = New clsBitmap
        
        'make sure the path is ok, otherwise, just draw
        If (Dir(mstrPicturePath) <> "") And _
           (Trim(mstrPicturePath) <> "") Then
            'valid path
            Call udtTemp.LoadBitmap(mstrPicturePath)
        End If
        
        'if the picture wasn't loaded, then
        'set to wallpaper
        If udtTemp.hdc = 0 Then
            menmBackgroundStyle = clkWallpaper
        End If
        
        'create a new bitmap size
        Call .SetBitmap(mintWidth, _
                        mintHeight, _
                        mlngBorderColour)
        
        With udtTemp
            'draw the picture onto the
            'background surphase
            Call DrawPicture(.hdc, _
                             .Width, _
                             .Height)
        End With
    End With
End Sub

Private Sub DrawPicture(ByVal lngSourceDC As Long, _
                        ByVal intWidth As Integer, _
                        ByVal intHeight As Integer)
    'This will redraw the picture with the given
    'picture mode (eg, tile, stretch, centre, normal)
    
    Dim intRow          As Integer      'the number of rows to tile
    Dim intCol          As Integer      'the number of columns to tile
    Dim intX            As Integer      'co-ordinates to tile picture
    Dim intY            As Integer      'co-ordinate to tile picture
    Dim bmpTemp         As clsBitmap    'holds the bitmap being drawn
    Dim intMinWidth     As Integer      'the minimum bitmap width to display
    Dim intMinHeight    As Integer      'the minimum bitmap height to display
    
    'check the background style
    If (menmBackgroundStyle = clkNone) Then
        'if the mode is None, then exit
        Exit Sub
    End If
    
    'create a temperory bitmap
    Set bmpTemp = New clsBitmap
    Call bmpTemp.SetBitmap(mintWidth, mintHeight)
    
    'draw the picture according to the set style
    Select Case menmBackgroundStyle
    Case clkNormal
        'get the minimum amount to display
        intMinWidth = IIf(intWidth >= mintWidth, mintWidth, intWidth)
        intMinHeight = IIf(intHeight >= mintHeight, mintHeight, intHeight)
        
        'just draw what can fit onto the background
        Call bmpTemp.PaintFrom(lngSourceDC, _
                               intMinWidth, _
                               intMinHeight, _
                               0, _
                               0, _
                               0, _
                               0, _
                               intMinWidth, _
                               intMinHeight)
    
    Case clkTile
        With bmpTemp
            'get the number of rows and columns we have
            'to display
            If intWidth <> 0 Then
                intRow = (.Width \ intWidth)
            End If
            If intHeight <> 0 Then
                intCol = (.Height \ intHeight)
            End If
            
            For intX = 0 To intRow
                For intY = 0 To intCol
                    'copy one tile to the appropiate area
                    Call bmpTemp.PaintFrom(lngSourceDC, _
                                           intWidth, _
                                           intHeight, _
                                           0, _
                                           0, _
                                           (intX * CLng(intWidth)), _
                                           (intY * CLng(intHeight)), _
                                           intWidth, _
                                           intHeight)
                Next intY
            Next intX
        End With
    
    Case clkStretch
        'stretch the picture to fit the size
        Call bmpTemp.PaintFrom(lngSourceDC, _
                               intWidth, _
                               intHeight, _
                               0, _
                               0, _
                               0, _
                               0, _
                               mintWidth, _
                               mintHeight)
    
    Case clkCentre
        'centre the picture in the bitmap
        Call bmpTemp.PaintFrom(lngSourceDC, _
                               intWidth, _
                               intHeight, _
                               0, _
                               0, _
                               (bmpTemp.Width - intWidth) / 2, _
                               (bmpTemp.Height - intHeight) / 2, _
                               intWidth, _
                               intHeight)
    
    Case clkWallpaper
        'capture the wallpaper behind the clock and use that as the
        'background picture. This only appears to work on certain
        'operating systems though.
        If (Not IsWinNT) Or (IsWinNT And (Not IsXp)) Then
            '95,NT,2000
            Call bmpTemp.GetWallpaper(mintScreenX, _
                                      mintScreenY)
        
        Else
            'assume "None"
            menmBackgroundStyle = clkNone
            
            Call bmpTemp.PaintFrom(lngSourceDC, _
                                   intWidth, _
                                   intHeight, _
                                   0, _
                                   0, _
                                   0, _
                                   0, _
                                   IIf(mintWidth >= intWidth, mintWidth, intWidth), _
                                   IIf(mintHeight >= intHeight, mintHeight, intHeight))
        End If
        
    Case Else   'default to the Normal background mode
        menmBackgroundStyle = clkNormal
        
        'just draw what can fit onto the background
        Call bmpTemp.PaintFrom(lngSourceDC, _
                               intWidth, _
                               intHeight, _
                               0, _
                               0, _
                               0, _
                               0, _
                               IIf(mintWidth >= intWidth, mintWidth, intWidth), _
                               IIf(mintHeight >= intHeight, mintHeight, intHeight))
    End Select
    
    'copy the drawn picture and delete the temperory
    'bitmap from memory
    Call bmpTemp.Paint(mudtBackPic.hdc)
    
    'set the handle of the background picture to
    'the clocks so that they know to draw a background
    'picture
    mclkAnalog.BackgroundDc = mudtBackPic.hdc
    mclkTime.BackgroundDc = mudtBackPic.hdc
    mclkDay.BackgroundDc = mudtBackPic.hdc
    mclkDate.BackgroundDc = mudtBackPic.hdc
End Sub

Public Sub GetScreenPos(ByVal frmParent As Form)
    'This will automatically update the screen position of the
    'clock based on the form it is on
    
    Dim intFormLeft     As Integer  'the left position of the form in pixels
    Dim intFormTop      As Integer  'the top position of the form in pixels
    Dim intFormWidth    As Integer  'the width of the form in pixels
    Dim intFormHeight   As Integer  'the height of the form in pixels
    Dim intLeftOffset   As Integer  'the size of the left form border in pixels
    Dim intTopOffset    As Integer  'the size of the border + titlebar + menus in pixels
    Dim intScaleWidth   As Integer  'the forms internal width in pixels
    Dim intScaleHeight  As Integer  'the forms internal height in pixels
    
    'make sure a form was passed
    If frmParent Is Nothing Then
        Exit Sub
    End If
    
    With frmParent
        'get the internal size of the form
        Select Case .ScaleMode
        Case vbTwips
            intScaleWidth = .ScaleWidth \ Screen.TwipsPerPixelX
            intScaleHeight = .ScaleHeight \ Screen.TwipsPerPixelY
            
        Case vbPixels
            intScaleWidth = .ScaleWidth
            intScaleHeight = .ScaleHeight
        End Select
        
        'get the size and position of the form in pixels
        intFormLeft = .Left \ Screen.TwipsPerPixelX
        intFormTop = .Top \ Screen.TwipsPerPixelY
        intFormWidth = .Width \ Screen.TwipsPerPixelX
        intFormHeight = .Height \ Screen.TwipsPerPixelY
        
        'get the size of the left border
        intLeftOffset = (intFormWidth - intScaleWidth) / 2
        intTopOffset = (intFormHeight - intScaleHeight - (intLeftOffset + 2))
        
        'update the screen position of the class
        mintScreenX = intFormLeft + intLeftOffset + mintLeft
        mintScreenY = intFormTop + intTopOffset + mintTop
    End With
    
    'update the display
    Call mudtBackPic.GetWallpaper(mintScreenX + mintBorderWidth, _
                                  mintScreenY + mintBorderWidth)
    Call Refresh
End Sub
