VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timed Shutdown Options"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   FillColor       =   &H80000006&
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraShut 
      Caption         =   "Shut Down Settings"
      Height          =   2415
      Left            =   3240
      TabIndex        =   17
      Top             =   120
      Width           =   3855
      Begin VB.CheckBox chkPrev 
         Alignment       =   1  'Right Justify
         Caption         =   "Prevent Other Apps From Closing Windows"
         Height          =   400
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   3375
      End
      Begin VB.ComboBox cboMethod 
         Height          =   315
         ItemData        =   "frmOptions.frx":030A
         Left            =   2160
         List            =   "frmOptions.frx":030C
         TabIndex        =   21
         Text            =   "Shut Down"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkIdle 
         Alignment       =   1  'Right Justify
         Caption         =   "Idle Shutdown On"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   3375
      End
      Begin VB.ComboBox cboMin 
         Height          =   315
         ItemData        =   "frmOptions.frx":030E
         Left            =   3000
         List            =   "frmOptions.frx":03C6
         TabIndex        =   19
         Text            =   "00"
         Top             =   1920
         Width           =   615
      End
      Begin VB.ComboBox cboHour 
         Height          =   315
         ItemData        =   "frmOptions.frx":04BA
         Left            =   2160
         List            =   "frmOptions.frx":0506
         TabIndex        =   18
         Text            =   "01"
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblMethod 
         BackStyle       =   0  'Transparent
         Caption         =   "Shut Down Method"
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label lblIdle 
         BackStyle       =   0  'Transparent
         Caption         =   "Shutdown If Idle For"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   1920
         Width           =   1770
      End
      Begin VB.Label lblMin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2865
         TabIndex        =   24
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label lblIdleHour 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   23
         Top             =   1920
         Width           =   120
      End
   End
   Begin VB.Frame fraDaily 
      Caption         =   "Daily Settings"
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txtDelay 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "15"
         ToolTipText     =   "Delay time in seconds"
         Top             =   1920
         Width           =   375
      End
      Begin VB.PictureBox picTime 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   1320
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   11
         Top             =   840
         Width           =   975
         Begin VB.TextBox txtSec 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   690
            MaxLength       =   2
            TabIndex        =   3
            Text            =   "00"
            ToolTipText     =   "Seconds"
            Top             =   0
            Width           =   255
         End
         Begin VB.TextBox txtMin 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            MaxLength       =   2
            TabIndex        =   2
            Text            =   "00"
            ToolTipText     =   "Minutes"
            Top             =   0
            Width           =   255
         End
         Begin VB.TextBox txtHour 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   30
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "00"
            ToolTipText     =   "Hours"
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lblMinSecSep 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   13
            Top             =   0
            Width           =   135
         End
         Begin VB.Label lblHourMinSep 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.ComboBox cboDays 
         Height          =   315
         ItemData        =   "frmOptions.frx":056A
         Left            =   1320
         List            =   "frmOptions.frx":056C
         TabIndex        =   0
         Text            =   "Weekday"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkShut 
         Alignment       =   1  'Right Justify
         Caption         =   "Shut Down On/Off"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox chkConfirm 
         Alignment       =   1  'Right Justify
         Caption         =   "Confirm Shut Down"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblSeconds 
         BackStyle       =   0  'Transparent
         Caption         =   "Wait For"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblSec 
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds"
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Top             =   1980
         Width           =   900
      End
      Begin VB.Label lblDay 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Day"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Time 24H"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Time in 24 hours"
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "frmOptions"
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
'TITLE :    Alarm Options Screen
' -----------------------------------------------
'COMMENTS :
'This screen is used to set the various alarm
'options for the form frmClock and is not intended
'for use outside the Timed Shutdown Clock project.
'=================================================

'require variable decaration
Option Explicit

'------------------------------------------------
'               MODULE-LEVEL VARIABLES
'------------------------------------------------
Private mudtWeek(6) As TypeShutdown     'temperorily holds the daily settings
Private mblnLoading As Boolean          'a flag that is set only when the form is loading

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Private Sub cboDays_Change()
    'display the setting for the current day
    If (cboDays.ListIndex <= UBound(gudtDay)) Then
        Call DisplaySettingsForDay(cboDays.ListIndex)
    End If
End Sub

Private Sub cboDays_Click()
    'trigger the change event
    Call cboDays_Change
End Sub

Private Sub cboDays_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    'trigger the change event
    Call cboDays_Change
End Sub

Private Sub cboHour_Change()
    'update the appropiate settings
    If mblnLoading Then
        'don't save any changes - the program is only
        'entering initial values
        Exit Sub
    End If
    
    'make sure there is a minimum limit of 1 minute
    'for the idle time
    If (cboHour.ListIndex = 0) And _
       (cboMin.ListIndex = 0) Then
        'adjust the minute box
        cboMin.ListIndex = 1
    End If
    
    glngIdleTime = (cboHour.ListIndex * CLng(3600)) + _
                   (cboMin.ListIndex * CLng(60))
    frmClock.mtscAnalog.AlarmTime = GetShutdownTime
End Sub

Private Sub cboHour_KeyPress(KeyAscii As Integer)
    'prevents the user changing text data
    Call cboHour_Change
    KeyAscii = 0
End Sub

Private Sub cboHour_Click()
    Call cboHour_Change
End Sub

Private Sub cboMethod_Change()
    'update the settings
    If mblnLoading Then
        'don't save any changes - the program is only
        'entering initial values
        Exit Sub
    End If
    
    genmShutdownMethod = cboMethod.ListIndex
End Sub

Private Sub cboMethod_KeyPress(KeyAscii As Integer)
    'prevents the user changing text data
    KeyAscii = 0
End Sub

Private Sub cboMethod_Click()
    'trigger a change
    Call cboMethod_Change
End Sub

Private Sub cboMin_Change()
    'update the appropiate settings
    If mblnLoading Then
        'don't save any changes - the program is only
        'entering initial values
        Exit Sub
    End If
    
    'make sure that the minimum limit of 1 minute is
    'not compromised
    If (cboHour.ListIndex = 0) And _
       (cboMin.ListIndex = 0) Then
        'set the minimum limit
        cboMin.ListIndex = 1
    End If
    
    'update the number of minutes
    'Formula: Current_Total_Time - _
              Current_Minutes + _
              New_Amount_Minutes
    glngIdleTime = glngIdleTime - _
                   (GetIdleMinutes * CLng(60)) + _
                   (cboMin.ListIndex * CLng(60))
    
    'update the alarm
    frmClock.mtscAnalog.AlarmTime = GetShutdownTime
End Sub

Private Sub cboMin_KeyPress(KeyAscii As Integer)
    'prevents the user changing text data
    KeyAscii = 0
    Call cboMin_Change
End Sub

Private Sub cboMin_Click()
    Call cboMin_Change
End Sub

Private Sub chkConfirm_Click()
    'update the appropiate settings
    
    Dim intCounter As Integer   'used to cycle through all the setting if selected
    
    If cboDays.ListIndex > UBound(mudtWeek) Then
        'update all days
        For intCounter = LBound(mudtWeek) To UBound(mudtWeek)
            mudtWeek(intCounter).blnDoDelay = chkConfirm.Value
        Next intCounter
    Else
        'just update the selected day
        mudtWeek(cboDays.ListIndex).blnDoDelay = chkConfirm.Value
    End If
End Sub

Private Sub chkIdle_Click()
    'update the appropiate settings
    If mblnLoading Then
        'don't save any changes - the program is only
        'entering initial values
        Exit Sub
    End If
    
    gblnIdleShut = chkIdle.Value
    frmClock.mtscAnalog.AlarmTime = GetShutdownTime
End Sub

Private Sub chkPrev_Click()
    'update the appropiate settings
    If mblnLoading Then
        'don't save any changes - the program is only
        'entering initial values
        Exit Sub
    End If
    
    gblnExclusiveShut = chkPrev.Value
End Sub

Private Sub chkShut_Click()
    'update the appropiate settings
    
    Dim intCounter As Integer   'used to cycle through all the setting if selected
    
    If cboDays.ListIndex > UBound(mudtWeek) Then
        'update all days
        For intCounter = LBound(mudtWeek) To UBound(mudtWeek)
            mudtWeek(intCounter).blnDoShutdown = chkShut.Value
        Next intCounter
    Else
        'just update the selected day
        mudtWeek(cboDays.ListIndex).blnDoShutdown = chkShut.Value
    End If
End Sub

Private Sub cmdCancel_Click()
    'exit without saving
    Unload Me
End Sub

Private Sub cmdOk_Click()
    'update and save all appropiate settings before
    'exiting this screen
    
    Dim intCounter As Integer   'used to cycle through the daily settings
    
    'if [All] was selected, make sure the changes
    'have been made to all daily settings before
    'saving
    If cboDays.ListIndex > UBound(gudtDay) Then
        For intCounter = LBound(mudtWeek) To UBound(mudtWeek)
            With mudtWeek(intCounter)
                .blnDoShutdown = chkShut.Value
                .blnDoDelay = chkConfirm.Value
                .intDelay = Val(txtDelay.Text)
                .intHour = Val(txtHour.Text)
                .intMinute = Val(txtMin.Text)
                .intSecond = Val(txtSec.Text)
            End With
        Next intCounter
    End If
    
    'update the daily settings
    For intCounter = 0 To 6
        gudtDay(intCounter) = mudtWeek(intCounter)
    Next intCounter
    
    'save the settings before exiting
    Call SaveSettings
    
    'set the new time
    frmClock.mtscAnalog.AlarmTime = GetShutdownTime
    
    'exit
    Unload Me
End Sub

Private Sub Form_Load()
    'set up the initial data in the controls
    mblnLoading = True
    Call Populate
    mblnLoading = False
    
    'set the form fonts to the current system default
    Call SetFormFontsToSystem(Me, FNT_MESSAGE)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'flag to clear this form from memory
    gblnDoCleanUp = True
End Sub

Private Sub Populate()
    'This will fill in all current data into the
    'controls on the form.
    
    Dim intCounter As Integer   'used to enter specific data into the combo box
    
    'set up the Days combo box
    For intCounter = 0 To 7
        If intCounter < 7 Then
            'enter a day name
            Call cboDays.AddItem(WeekdayName(intCounter + 1), _
                                 intCounter)
        Else
            'the last item in the box is to indicate
            'ALL days
            Call cboDays.AddItem("[All]", _
                                 intCounter)
        End If
    Next intCounter
    'select today's settings
    cboDays.ListIndex = Today
    
    'set up the shutdown combo box
    For intCounter = shtLogOut To shtLockWorkstation
        'only add the Lock Workstation option if this is a Win 2000 or XP machine
        If (intCounter < shtLockWorkstation) Or ((intCounter = shtLockWorkstation) And (IsW2000)) Then
            Call cboMethod.AddItem(GetShutText(intCounter), _
                                   intCounter)
        End If
    Next intCounter
    'select the current shutdown method
    cboMethod.ListIndex = genmShutdownMethod
    
    'set to display the current hour of the shutdown
    'time
    cboHour.ListIndex = (glngIdleTime \ 3600)
    
    'set to display the current minute of the shutdown
    'time
    cboMin.ListIndex = GetIdleMinutes
    
    'if the idle time on or off
    chkIdle.Value = Abs(gblnIdleShut)
    
    'shut we prevent other apps from shutting down the
    'computer?
    chkPrev.Value = Abs(gblnExclusiveShut)
    
    'copy the array settings
    For intCounter = 0 To 6
        mudtWeek(intCounter) = gudtDay(intCounter)
    Next intCounter
    
    'set the daily settings for the shutdown for today
    Call DisplaySettingsForDay(Today)
End Sub

Private Sub DisplaySettingsForDay(ByVal intDay As Integer)
    'display the current settings for the specified day
    
    'make sure that the day is within valid ranges
    If (intDay < LBound(mudtWeek)) Or _
       (intDay > UBound(mudtWeek)) Then
        'invalid index, default to today
        intDay = Today
    End If
    
    'set the daily settings for the shutdown for today
    With mudtWeek(intDay)
        txtHour.Text = Format(.intHour, "00")
        txtMin.Text = Format(.intMinute, "00")
        txtSec.Text = Format(.intSecond, "00")
        
        chkShut.Value = Abs(.blnDoShutdown)
        chkConfirm.Value = Abs(.blnDoDelay)
        txtDelay.Text = .intDelay
    End With
End Sub

Private Sub txtDelay_Change()
    'update the appropiate settings
    
    Dim intCounter As Integer   'used to cycle through all the setting if selected
    
    If cboDays.ListIndex > UBound(mudtWeek) Then
        'update all days
        For intCounter = LBound(mudtWeek) To UBound(mudtWeek)
            mudtWeek(intCounter).intDelay = Val(txtDelay.Text)
        Next intCounter
    Else
        'just update the selected day
        mudtWeek(cboDays.ListIndex).intDelay = Val(txtDelay.Text)
    End If
End Sub

Private Sub txtDelay_GotFocus()
    'highlight the text in the text box
    Call HighLight(txtDelay)
End Sub

Private Sub txtDelay_Validate(Cancel As Boolean)
    'make sure the delay time is between the valid
    'ranges
    txtDelay.Text = modNumbers.LimitRange(Val(txtDelay.Text), 1, 60)
    
    'save the settings appropiatly
    Call txtDelay_Change
End Sub

Private Sub txtHour_Change()
    'update the changes
    
    Dim intCounter As Integer   'used to cycle through all the setting if selected
    
    If cboDays.ListIndex > UBound(mudtWeek) Then
        'update all days
        For intCounter = LBound(mudtWeek) To UBound(mudtWeek)
            mudtWeek(intCounter).intHour = Val(txtHour.Text)
        Next intCounter
    Else
        'just update the selected day
        mudtWeek(cboDays.ListIndex).intHour = Val(txtHour.Text)
    End If
End Sub

Private Sub txtHour_GotFocus()
    'highlight the text in the text box
    Call HighLight(txtHour)
End Sub

Private Sub txtHour_Validate(Cancel As Boolean)
    'make sure the hour is between the valid ranges
    txtHour.Text = Format(modNumbers.LimitRange(Val(txtHour.Text), 0, 23), "00")
    
    'save the settings appropiatly
    Call txtHour_Change
End Sub

Private Sub txtMin_Change()
    'update the changes
    
    Dim intCounter As Integer   'used to cycle through all the setting if selected
    
    If cboDays.ListIndex > UBound(mudtWeek) Then
        'update all days
        For intCounter = LBound(mudtWeek) To UBound(mudtWeek)
            mudtWeek(intCounter).intMinute = Val(txtMin.Text)
        Next intCounter
    Else
        'just update the selected day
        mudtWeek(cboDays.ListIndex).intMinute = Val(txtMin.Text)
    End If
End Sub

Private Sub txtMin_GotFocus()
    'highlight the text in the text box
    Call HighLight(txtMin)
End Sub

Private Sub txtMin_Validate(Cancel As Boolean)
    'make sure the minute is between the valid ranges
    txtMin.Text = Format(modNumbers.LimitRange(Val(txtMin.Text), 0, 59), "00")
    
    'save the settings appropiatly
    Call txtMin_Change
End Sub

Private Sub txtSec_Change()
    'update the changes
    
    Dim intCounter As Integer   'used to cycle through all the setting if selected
    
    If cboDays.ListIndex > UBound(mudtWeek) Then
        'update all days
        For intCounter = LBound(mudtWeek) To UBound(mudtWeek)
            mudtWeek(intCounter).intSecond = Val(txtSec.Text)
        Next intCounter
    Else
        'just update the selected day
        mudtWeek(cboDays.ListIndex).intSecond = Val(txtSec.Text)
    End If
End Sub

Private Sub txtSec_GotFocus()
    'highlight the text in the text box
    Call HighLight(txtSec)
End Sub

Private Sub txtSec_Validate(Cancel As Boolean)
    'make sure the second is between the valid ranges
    txtSec.Text = Format(modNumbers.LimitRange(Val(txtSec.Text), 0, 59), "00")
    
    'save the settings appropiatly
    Call txtSec_Change
End Sub

Private Function GetIdleMinutes() As Integer
    'returns the current number of minutes (excluding
    'hours and seconds) for the idle time
    'GetIdleMinutes = (((glngIdleTime - (glngIdleTime Mod 60)) \ 60) Mod 60)
    GetIdleMinutes = (glngIdleTime \ 60) Mod 60
End Function

