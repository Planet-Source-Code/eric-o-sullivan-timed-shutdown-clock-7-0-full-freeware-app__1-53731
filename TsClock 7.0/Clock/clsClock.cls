VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This clock can display an analog or digital clock, the time, day or date, and can also notify when a particular time has occurred."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsTest"
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     2 August 2002
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    Clock Class
' -----------------------------------------------
'COMMENTS :
'Provides the basic functionality of a clock,
'capable of displaying in four styles; Digital,
'Analog, Day and Date
'=================================================

'require variable declatation
Option Explicit

'------------------------------------------------
'                   ENUMERATORS
'------------------------------------------------
'the different background styles
Public Enum EnmCBackgroundStyle
    clkNone = 0
    clkNormal = 1
    clkStretch = 2
    clkTile = 3
    clkCentre = 4
    clkWallpaper = 5
End Enum

'the kind of clock to display
Public Enum EnmCDisplayStyle
    clkAnalog = 0
    clkDigital = 1
    clkDate = 2
    clkDay = 3
End Enum

'------------------------------------------------
'                   CLASS EVENTS
'------------------------------------------------

'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent AlarmActivate[(arg1, arg2, ... , argn)]

'Please note that the events are fired in the order
'they appear here;

'This event fires when the time has changed
Public Event NewTime(ByVal NewTime As Date)

'This event fires when a new day has been found
Public Event NewDay(ByVal NewDate As Date)

'This event triggers then the time matches
'the alarm time. The event only triggers
'ONCE per second.
Public Event AlarmActivate()
Attribute AlarmActivate.VB_Description = "Triggers when the alarm time matches the time properties"

'------------------------------------------------
'               MODULE-LEVEL VARIABLES
'------------------------------------------------

'local variable(s) to hold property value(s)
Private mlngBackgroundDc        As Long                 'a handle to the background picture
Private mstrTime                As String               'the given time
Private mstrClockDate           As String               'the given date
Private mblnAutoUpdateData      As Boolean              'automatically gets the date and time from the system, but does not update the display
Private mblnAutoDisplay         As Boolean              'automatically displays the current background when the Refresh method is called
Private mintClockWidth          As Integer              'the width of the clock
Private mintClockHeight         As Integer              'the height of the clock
Private menmBackgroundStyle     As EnmCBackgroundStyle  'the background picture style
Private mlngBackgroundColour    As Long                 'the background colour
Private menmDisplayType         As EnmCDisplayStyle     'the display type, eg analog, digital etc
Private mblnDisplayBackground   As Boolean              'whether or not to display the background picture or the background colour
Private mintClockTop            As Integer              'the top position of the clock on the surphase
Private mintClockLeft           As Integer              'the left position of the clock on the surphase
Private mlngHandHourColour      As Long                 'the colour of the hour hand
Private mlngHandMinuteColour    As Long                 'the colour of the minute hand
Private mlngHandSecondColour    As Long                 'the colour of the second hand
Private mlngForeColour          As Long                 'the colour of any text displayed
Private mintBackgroundX         As Integer              'the left position of where to copy the background from
Private mintBackgroundY         As Integer              'the top position of where to copy the background from
Private mstrAlarmTime           As String               'the time to trigger the alarm (only triggers once)
Private mlngDotColour           As Long                 'the colour of the dots on the analog display
Private mblnVisible             As Boolean              'whether or not to display any graphical data (eg if you only want an alarm)
Private mfntFont                As StdFont              'the font settings for the text
Private mlngSurphaseDc          As Long                 'the surphase to draw the clock on
Private mblnTime24Hour          As Boolean              'whether or not to display the time in 24 hour format
Private mbmpTemp                As clsBitmap            'temperorily holds the background picture before blitting to the screen

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Public Property Let Time24Hour(ByVal vData As Boolean)
Attribute Time24Hour.VB_Description = "Whether or not to display the time in 24 hour format or 12 hour format"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.Time24Hour = 5
    mblnTime24Hour = vData
    
    Call RebuildClock(True)
End Property

Public Property Get Time24Hour() As Boolean
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.Time24Hour
    Time24Hour = mblnTime24Hour
End Property

Public Property Let SurphaseDC(ByVal vData As Long)
Attribute SurphaseDC.VB_Description = "This is the surphase that we are supposed to copy to"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.SurphaseDC = 5
    mlngSurphaseDc = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get SurphaseDC() As Long
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.SurphaseDC
    SurphaseDC = mlngSurphaseDc
End Property

Public Property Set Font(ByVal fntNew As StdFont)
Attribute Font.VB_Description = "The font settings of any text displayed"
'used when assigning an Object to the property,
'on the left side of a Set statement.
'Syntax: Set x.Font = Form1
    
    'make sure the font object exists
    If mfntFont Is Nothing Then
        Set mfntFont = New StdFont
    End If
    
    'copy the font details - we don't want to create
    'a reference to the object being passed
    With mfntFont
        .Bold = fntNew.Bold
        .Charset = fntNew.Charset
        .Italic = fntNew.Italic
        .Name = fntNew.Name
        .Size = fntNew.Size
        .Strikethrough = fntNew.Strikethrough
        .Underline = fntNew.Underline
        .Weight = fntNew.Weight
    End With
    
    'rebuild the clock
    Call AdjustFontSize
    Call RebuildClock(True)
End Property

Public Property Get Font() As StdFont
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.Font
    Set Font = mfntFont
End Property

Public Property Let Visible(ByVal blnShow As Boolean)
Attribute Visible.VB_Description = "Returns or sets whether or not the class has a visual interface."
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.Visible = 5
    mblnVisible = blnShow
    
    'rebuild the clock
    Call AdjustFontSize
    Call RebuildClock(True)
End Property

Public Property Get Visible() As Boolean
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.Visible
    Visible = mblnVisible
End Property

Public Sub Cls()
Attribute Cls.VB_Description = "This wipes the clock area on the Background with the Background colour"
    'This will wipe the surphase with the
    'background.
    
    Call DrawBackground
    Call PaintClock
End Sub

Public Property Let DotColour(ByVal vData As Long)
Attribute DotColour.VB_Description = "The colour of the dots displayed when the clock style is set to Analog"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.DotColour = 5
    mlngDotColour = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get DotColour() As Long
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.DotColour
    DotColour = mlngDotColour
End Property

Public Property Let AlarmTime(ByVal vData As String)
Attribute AlarmTime.VB_Description = "This will trigger the Alarm event when the specified time matches the AlarmTime. Only active when the AlarmTime is a valid time format, otherwise the Alarm event does NOT trigger."
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.AlarmTime = 5
    mstrAlarmTime = vData
    
    'check the alarm time
    Call CheckAlarm
End Property

Public Property Get AlarmTime() As String
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.AlarmTime
    AlarmTime = mstrAlarmTime
End Property

Public Property Let BackgroundY(ByVal vData As Integer)
Attribute BackgroundY.VB_Description = "The Y position of where on the Background hDC we should display the background picture from"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.BackgroundY = 5
    mintBackgroundY = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get BackgroundY() As Integer
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.BackgroundY
    BackgroundY = mintBackgroundY
End Property

Public Property Let BackgroundX(ByVal vData As Integer)
Attribute BackgroundX.VB_Description = "The X position of where on the Background hDC we should display the background picture from"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.BackgroundX = 5
    mintBackgroundX = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get BackgroundX() As Integer
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.BackgroundX
    BackgroundX = mintBackgroundX
End Property

Public Property Let ForeColour(ByVal vData As Long)
Attribute ForeColour.VB_Description = "The colour of the text being displayed"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.ForeColour = 5
    mlngForeColour = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get ForeColour() As Long
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.ForeColour
    ForeColour = mlngForeColour
End Property

Public Property Let HandSecondColour(ByVal vData As Long)
Attribute HandSecondColour.VB_Description = "The colour of the second hand on the clock"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.HandSecondColour = 5
    mlngHandSecondColour = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get HandSecondColour() As Long
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.HandSecondColour
    HandSecondColour = mlngHandSecondColour
End Property

Public Property Let HandMinuteColour(ByVal vData As Long)
Attribute HandMinuteColour.VB_Description = "The colour of the minute hand on the clock"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.HandMinuteColour = 5
    mlngHandMinuteColour = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get HandMinuteColour() As Long
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.HandMinuteColour
    HandMinuteColour = mlngHandMinuteColour
End Property

Public Property Let HandHourColour(ByVal vData As Long)
Attribute HandHourColour.VB_Description = "The colour of the hour hand on the analog clock"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.HandHourColour = 5
    mlngHandHourColour = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get HandHourColour() As Long
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.HandHourColour
    HandHourColour = mlngHandHourColour
End Property

Public Property Let ClockLeft(ByVal vData As Integer)
Attribute ClockLeft.VB_Description = "Returns or sets the horizontal position of the left side of the clock"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.ClockLeft = 5
    mintClockLeft = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get ClockLeft() As Integer
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.ClockLeft
    ClockLeft = mintClockLeft
End Property

Public Property Let ClockTop(ByVal vData As Integer)
Attribute ClockTop.VB_Description = "Returns or sets the vertical position of the top of the class"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.ClockTop = 5
    mintClockTop = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get ClockTop() As Integer
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.ClockTop
    ClockTop = mintClockTop
End Property

Public Property Let DisplayBackground(ByVal vData As Boolean)
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.DisplayBackground = 5
    mblnDisplayBackground = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get DisplayBackground() As Boolean
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.DisplayBackground
    DisplayBackground = mblnDisplayBackground
End Property

Public Property Let DisplayType(ByVal vData As EnmCDisplayStyle)
Attribute DisplayType.VB_Description = "Returns or sets what is displayed on the Background hDC\r\n\r\nStyle 0 = Analog      [Default]\r\nStyle 1 = Digital\r\nStyle 2 = Date\r\nStyle 3 = Day"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.DisplayType = 5
    menmDisplayType = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get DisplayType() As EnmCDisplayStyle
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.DisplayType
    DisplayType = menmDisplayType
End Property

Public Property Let BackgroundColour(ByVal vData As Long)
Attribute BackgroundColour.VB_Description = "Returns or sets the background colour of the class"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.BackgroundColour = 5
    mlngBackgroundColour = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get BackgroundColour() As Long
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.BackgroundColour
    BackgroundColour = mlngBackgroundColour
End Property

Public Property Let BackgroundStyle(ByVal vData As EnmCBackgroundStyle)
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.BackgroundStyle = 5
    menmBackgroundStyle = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get BackgroundStyle() As EnmCBackgroundStyle
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.BackgroundStyle
    BackgroundStyle = menmBackgroundStyle
End Property

Public Property Let ClockHeight(ByVal intNewHeight As Integer)
Attribute ClockHeight.VB_Description = "The height of the area to display"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.ClockHeight = 5
    mintClockHeight = intNewHeight
    
    'rebuild the clock
    Call AdjustFontSize
    Call RebuildClock(True)
End Property

Public Property Get ClockHeight() As Integer
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.ClockHeight
    ClockHeight = mintClockHeight
End Property

Public Property Let ClockWidth(ByVal intNewWidth As Integer)
Attribute ClockWidth.VB_Description = "The width of the area to display"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.ClockWidth = 5
    mintClockWidth = intNewWidth
    
    'rebuild the clock
    Call AdjustFontSize
    Call RebuildClock(True)
End Property

Public Property Get ClockWidth() As Integer
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.ClockWidth
    ClockWidth = mintClockWidth
End Property

Public Property Let AutoDisplay(ByVal vData As Boolean)
Attribute AutoDisplay.VB_Description = "Returns or set whether or not the class is to automatically update the display when the time changes"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.AutoDisplay = 5
    mblnAutoDisplay = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get AutoDisplay() As Boolean
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.AutoDisplay
    AutoDisplay = mblnAutoDisplay
End Property

Public Property Let AutoUpdateData(ByVal vData As Boolean)
Attribute AutoUpdateData.VB_Description = "Returns or sets whether or not to get the settings from the system clock and to update the properties when they change"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.AutoUpdateData = 5
    mblnAutoUpdateData = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get AutoUpdateData() As Boolean
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.AutoUpdateData
    AutoUpdateData = mblnAutoUpdateData
End Property

Public Property Let ClockDate(ByVal vData As String)
Attribute ClockDate.VB_Description = "Returns or sets the date"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.ClockDate = 5
    mstrClockDate = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get ClockDate() As String
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.ClockDate
    ClockDate = mstrClockDate
End Property

Public Property Let ClockTime(ByVal vData As String)
Attribute ClockTime.VB_Description = "Returns or sets the time"
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.Time = 5
    mstrTime = vData
    
    'check for the alarm
    Call CheckAlarm
End Property

Public Property Get ClockTime() As String
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.Time
    Time = mstrTime
End Property

Public Property Let BackgroundDc(ByVal vData As Long)
'used when assigning a value to the property,
'on the left side of an assignment.
'Syntax: X.BackgroundDc = 5
    mlngBackgroundDc = vData
    
    'rebuild the clock
    Call RebuildClock(True)
End Property

Public Property Get BackgroundDc() As Long
'used when retrieving value of a property,
'on the right side of an assignment.
'Syntax: Debug.Print X.BackgroundDc
    BackgroundDc = mlngBackgroundDc
End Property

Public Sub Refresh()
    'This method is "reserved" as has strange effects
    'when called internally, hence the code is located
    'in the procedure RebuildClock()
    Call RebuildClock
End Sub

Private Sub RebuildClock(Optional ByVal blnForce As Boolean = False)
    'This will rebuild the clock and all its'
    'settings
    
    Static strDate As String    'remembers the last date. This is checked to see if we need to fire the NewDay event
    Static strTime As String    'remembers the last time
    
    'there is no point in updating any data unless the
    'time has changed, requiring the graphic to be
    'rebuilt. However, the clock can be "forced" to
    'redraw itself
    If (strTime <> Time) Or (blnForce) Then
        'do we trigger the NewTime event
        If strTime <> Time Then
            RaiseEvent NewTime(Time)
        End If
        
        'get automatic settings if necessary
        If mblnAutoUpdateData Then
            mstrTime = Time
            mstrClockDate = Date
        End If
        
        'check for a new day
        If strDate <> mstrClockDate Then
            If strDate <> "" Then
                'a new day
                RaiseEvent NewDay(mstrClockDate)
            End If
        End If
            
        'remember the current date
        strDate = mstrClockDate
        
        'check the time for the alarm time
        'and raise event if necessary
        Call CheckAlarm
        
        If mblnVisible Then
            Call DrawGraphic
        End If
        
        'update the new time
        strTime = Time
    End If
    
    'display the current graphic
    If mblnAutoDisplay Then
        'display the clock
        Call PaintClock
    End If
End Sub

Private Sub CheckAlarm()
    
    Static blnActive As Boolean     'only trigger the alarm once
    
    'check to see if the alarm needs to
    'be triggered
    If mstrAlarmTime <> "" Then
        If mstrTime <> "" Then
            If (DateDiff("s", mstrTime, mstrAlarmTime) = 0) Then
                'if the alarm has not already been
                'triggered this second, then
                If Not blnActive Then
                    'raise the alarm event
                    blnActive = True
                    RaiseEvent AlarmActivate
                End If
            Else
                'reset the alarm
                blnActive = False
            End If
        End If
    End If
End Sub

Public Sub PaintClock()
    'this will copy the clock onto the
    'background
    
    'check if we need to display graphical
    'information
    If (Not mblnVisible) Or (mlngSurphaseDc = 0) Then
        Exit Sub
    End If
    
    Call mbmpTemp.Paint(mlngSurphaseDc, _
                        mintClockLeft, _
                        mintClockTop)
End Sub

Private Sub DrawGraphic()
    'This will either draw the appropiate
    'text, or it will draw the analog
    'graphic onto the bitmap
    
    'get any bitmap dimensions
    'Call SetClockSize
    
    'draw a background (if any)
    Call DrawBackground
    
    Select Case menmDisplayType
    Case clkAnalog
        Call DrawAnalog
    
    Case Else
        'draw the text onto the clock
        Call DrawText(GetClockText)
    End Select
End Sub

Private Function GetClockText() As String
    'This will return the text being displayed in the clock
    
    Select Case menmDisplayType
    Case clkDate
        GetClockText = mstrClockDate
        
    Case clkDay
        GetClockText = WeekdayName(Weekday(mstrClockDate, vbUseSystemDayOfWeek))
    
    Case clkDigital
        If mblnTime24Hour Then
            '24 hour time
            GetClockText = Format(mstrTime, "h:nn:ss")
        Else
            '12 hour time
            GetClockText = Format(mstrTime, "h:nn:ss am/pm")
        End If
    End Select
End Function

Private Sub DrawAnalog()
    'This will draw the analog graphic onto
    'the background
    
    Dim intCounter  As Integer  'used to display the hands and dots
    Dim intRadius   As Integer  'the radius of the clock. This is 2 pixels shorter than the smallest height or width
    Dim intX        As Integer  'the center of the clock
    Dim intY        As Integer  'the center of the clock
    Dim intPosX     As Integer  'temperory position calculated
    Dim intPosY     As Integer  'temperory position calculated
    Dim intTailX    As Integer  'temperory position of the hand "tail' (the bit sticking out from the center in the opposite direction the hand is pointing to)
    Dim intTailY    As Integer  'temperory position of the hand "tail' (the bit sticking out from the center in the opposite direction the hand is pointing to)
    Dim intSize     As Integer  'temperory size value for a graphic
    Dim intHandX(3) As Integer  'holds the X polygon points for drawing the hands
    Dim intHandY(3) As Integer  'holds the Y polygon points for drawing the hands
    Dim sngAngle    As Single   'holds the angle for the specified position
    
    'don't show the graphic if text is supposed
    'to be displayed
    If menmDisplayType <> clkAnalog Then
        Exit Sub
    End If
    
    'get the center co-ordinates of the clock
    intX = (mintClockWidth / 2) - 1
    intY = (mintClockHeight / 2) - 1
    
    'get the dot radius
    If intX < intY Then
        'the horizontal is smaller than
        'the vertical
        intRadius = intX - 2
    Else
        'the vertical is smaller than
        'the horizontal
        intRadius = intY - 2
    End If
    
    'Call mbmpTemp.Paint(frmClock.picTest.hDc)
    'draw the dots onto the bitmap
    For intCounter = 0 To 360 Step 6
        'calculate the position of the dot
        '[position] = [center point] + ( Sin|Cos ( [angle] ) * [radius] )
        '< [angle] * PI / 180 > = conversion to radians for triginometric function
        '180 degrees = PI radians       90 degrees = (PI / 2) radians
        intPosX = intX + (Sin(intCounter * Pi / 180) * intRadius)
        intPosY = intY + (Cos(intCounter * Pi / 180) * intRadius)
        
        'draw the dots larger if they are
        'at 5 minute intervals (30 degrees)
        If (intCounter Mod 30) = 0 Then
            '5 minute interval
            intSize = 2
        Else
            'normal minute (6 degrees)
            intSize = 1
        End If
        
        'draw the dot
        Call mbmpTemp.DrawRect(mlngDotColour, _
                               intPosY, _
                               intPosX, _
                               intSize, _
                               intSize)
    Next intCounter
    
    'largest hand first
    intRadius = intRadius - 3
    intSize = 1
    
    'draw the second hand
    '(6 degrees per second/minute)
    sngAngle = ((Second(mstrTime) * -6) + 180) * Pi / 180
    intPosX = intX + (Sin(sngAngle) * intRadius)
    intPosY = intY + (Cos(sngAngle) * intRadius)
    intTailX = intX + (Sin(sngAngle) * (intRadius / 3))
    intTailY = intY + (Cos(sngAngle) * (intRadius / 3))
    Call mbmpTemp.DrawLine(intTailX, _
                           intTailY, _
                           intPosX, _
                           intPosY, _
                           mlngHandSecondColour, _
                           intSize)
    
    
    'shrink for the minute hand
    intRadius = intRadius - (intRadius / 9)
    If intRadius < 1 Then
        intRadius = 1
    End If
    
    'draw the Minute hand
    '(6 degrees per second/minute)
    sngAngle = ((Minute(mstrTime) * 6) + 180) * -Pi / 180
    intHandX(0) = intX + (Sin(sngAngle) * intRadius)
    intHandY(0) = intY + (Cos(sngAngle) * intRadius)
    intHandX(1) = intX + (Sin(sngAngle + (Pi / 2)) * (intRadius / 20))
    intHandY(1) = intY + (Cos(sngAngle + (Pi / 2)) * (intRadius / 20))
    intHandX(2) = intX + (Sin(sngAngle) * -(intRadius / 7))
    intHandY(2) = intY + (Cos(sngAngle) * -(intRadius / 7))
    intHandX(3) = intX + (Sin(sngAngle + (Pi / 2)) * -(intRadius / 20))
    intHandY(3) = intY + (Cos(sngAngle + (Pi / 2)) * -(intRadius / 20))
    Call mbmpTemp.DrawPoly(intHandX(), _
                           intHandY(), _
                           mlngHandMinuteColour, _
                           mlngHandMinuteColour)
    
    'shrink for the hour hand
    intRadius = intRadius - (intRadius / 4)
    If intRadius < 1 Then
        intRadius = 1
    End If
    
    'draw the hour hand
    '(1/2 degree per minute)
    sngAngle = (((Hour(mstrTime) * 30) + (Minute(mstrTime) * 0.5)) + 180) * -Pi / 180
    intHandX(0) = intX + (Sin(sngAngle) * intRadius)
    intHandY(0) = intY + (Cos(sngAngle) * intRadius)
    intHandX(1) = intX + (Sin(sngAngle + (Pi / 2)) * (intRadius / 20))
    intHandY(1) = intY + (Cos(sngAngle + (Pi / 2)) * (intRadius / 20))
    intHandX(2) = intX + (Sin(sngAngle) * -(intRadius / 7))
    intHandY(2) = intY + (Cos(sngAngle) * -(intRadius / 7))
    intHandX(3) = intX + (Sin(sngAngle + (Pi / 2)) * -(intRadius / 20))
    intHandY(3) = intY + (Cos(sngAngle + (Pi / 2)) * -(intRadius / 20))
    Call mbmpTemp.DrawPoly(intHandX(), _
                           intHandY(), _
                           mlngHandHourColour, _
                           mlngHandHourColour)
End Sub

Private Sub AdjustFontSize()
    'This will adjust the font size to fill the clock
    
    'we only need to adjust this when the clock is visable
    If (mblnVisible) And (Not mfntFont Is Nothing) Then
        mfntFont.Size = mbmpTemp.GetFontPointSize(mfntFont, _
                                                  GetClockText, _
                                                  mintClockHeight, _
                                                  mintClockWidth)
    End If
End Sub

Private Sub DrawText(ByVal strText As String)
    'This procedure will only draw text
    'onto the background picture if the
    'background style is not analog.
    
    Dim intCentreHeight     As Integer      'holds the height to centre the text at
    
    'only display text if we are supposed to
    If menmDisplayType = clkAnalog Then
        Exit Sub
    End If
    
    'if there isn't any text to display, then exit
    If Trim(strText = "") Then
        Exit Sub
    End If
    
    With mbmpTemp
        'get the height to centre the text at
        intCentreHeight = (mintClockHeight - .GetTextHeight(mfntFont)) \ 2
        
        'draw the text onto the background
        Call .DrawString(strText, _
                         intCentreHeight, _
                         0, _
                         mintClockHeight, _
                         mintClockWidth, _
                         mfntFont, _
                         mlngForeColour)
    End With
End Sub

Private Sub DrawBackground()
    'This procedure will redraw the picture onto the
    'background bitmap.
    
    'check if we need to display graphical
    'information
    If Not mblnVisible Then
        Exit Sub
    End If
    
    'now draw the background if any
    'was specified
    If mlngSurphaseDc = 0 Then
        'no valid dc
        Exit Sub
    End If
    
    'redraw the basic bitmap
    Call SetClockSize
    
    If Not mblnDisplayBackground Then
        'set the background style to None
        menmBackgroundStyle = clkNone
    End If
    
    Select Case menmBackgroundStyle
    Case clkNone
        'do nothing. bitmap was already
        'created using default colour
        
    Case clkNormal
        'copy the bitmap from the specified
        'DC onto the clock background
        With mbmpTemp
            Call .PaintFrom(mlngBackgroundDc, _
                            mintClockWidth, _
                            mintClockHeight, _
                            mintBackgroundX, _
                            mintBackgroundY)
        End With
        
    Case clkStretch
        'stretch the bitmap onto the background
    Case clkTile
        'tile the bitmap onto the background
    Case Else
        menmBackgroundStyle = clkNone
        Exit Sub
    End Select
End Sub

Private Sub Class_Initialize()
    'create the background bitmap and
    'set any settings if necessary
    
    'create the background
    Set mbmpTemp = New clsBitmap
    
    'create the font
    Set mfntFont = New StdFont
    
    'create the default settings
    With Me
        .AlarmTime = ""
        .AutoDisplay = True
        .AutoUpdateData = True
        .BackgroundColour = vbBlack
        .BackgroundStyle = clkNone
        .ClockHeight = 50   'pixels
        .ClockWidth = 50    'pixels
        .DisplayBackground = False
        .DisplayType = clkDigital
        Set mfntFont = New StdFont
        With mfntFont
            .Name = "System"
            .Size = 12
        End With
        .ForeColour = vbYellow
        .HandHourColour = vbRed
        .HandMinuteColour = vbGreen
        .HandSecondColour = vbYellow
        .DotColour = vbBlack
        .Time24Hour = True
        .Visible = False
    End With
    
    'create any background settings as necessary
    Call DrawBackground
End Sub

Private Sub SetClockSize()
    'This will resize the bitmap to it's new size
    
    'if there is no surphase object
    'specified, then there is nothing
    'to draw on
    If mlngSurphaseDc = 0 Then
        Exit Sub
    End If
    
    'should we display any graphical data?
    If Not mblnVisible Then
        Exit Sub
    End If
    
    'test the new size
    If (mintClockWidth = 0) Or (mintClockHeight = 0) Then
        'no point in creating a bitmap
        Exit Sub
    End If
    
    'create the new background
    Call mbmpTemp.SetBitmap(mintClockWidth, _
                            mintClockHeight, _
                            mlngBackgroundColour)
End Sub

Private Sub Class_Terminate()
    'make sure that all data is removed from memory
    Set mbmpTemp = Nothing
End Sub
