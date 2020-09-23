VERSION 5.00
Begin VB.Form frmShut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shut Down"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2790
   Icon            =   "frmShut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   186
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer timShut 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame fraAuto 
      Caption         =   "Automatic Shutdown In"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2535
      Begin TimedShutdownClock.ctlProgBar cpbShut 
         Height          =   372
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2292
         _extentx        =   4048
         _extenty        =   661
         value           =   50
         percentcaption  =   0
         caption         =   "0 Seconds"
         backcolour      =   -2147483643
         font            =   "frmShut.frx":030A
      End
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblShut 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shut down the computer ?"
      Height          =   192
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2772
   End
End
Attribute VB_Name = "frmShut"
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
'TITLE :    Shutdown Screen
' -----------------------------------------------
'COMMENTS :
'This form is used to ask the user to either
'confirm or cancel shutdown. It can be set to
'shutdown the computer automatically after a
'certain period of time (within 1 minute)
'=================================================

'require variable declaration
Option Explicit

'------------------------------------------------
'              MODULE-LEVEL VARIABLES
'------------------------------------------------
Private mmosMove    As clsMouse 'used to move the mouse to a particular point on the screen

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Private Sub MoveToButton()
    'This procedure will move the mouse to the Yes
    'button on the form
    
    Dim intX    As Integer  'the x position we want to move to
    Dim intY    As Integer  'the y position we want to move to
    
    'get the pixel co-ordinate of the centre of the
    'Yes button on the form
    With frmShut
        intX = ((.Left / Screen.TwipsPerPixelX) + _
                (((.Width / Screen.TwipsPerPixelX) - .ScaleWidth) / 2) + _
                cmdYes.Left + (cmdYes.Width / 2))
        intY = ((.Top / Screen.TwipsPerPixelY) + _
                ((.Height / Screen.TwipsPerPixelY) - .ScaleHeight) - _
                (((.Width / Screen.TwipsPerPixelX) - .ScaleWidth) / 2) + _
                cmdYes.Top + (cmdYes.Height / 2))
    End With
    
    'move the mouse to the Yes button
    'Call MoveMouseTo(intX, intY)
    Call mmosMove.MoveTo(intX, intY)
End Sub

Public Sub Start()
    'This procedure will set up everything needed to
    'display the form correctly and ask the user if
    'they want to shut down the computer. This is to
    'replace the need for the Form_Load event so that
    'it can be explicitly called from outside the form
    
    Const DELAY_HEIGHT = 2070
    Const NO_DELAY_HEIGHT = 1350
    
    Dim blnEnableTimer As Boolean   'do we enable the timer or not (this has to be synchronized with the form visibility)
    
    With gudtDay(Today)
        'check if we are supposed to shut down the
        'computer
        If Not .blnDoShutdown Then
            Exit Sub
        End If
        
        'resize the form appropiatly based on the current
        'settings
        If .blnDoDelay Then
            'display the progress bar
            Me.Height = DELAY_HEIGHT
            cpbShut.Max = .intDelay * CLng(1000)
            cpbShut.Value = 0
            cpbShut.Caption = .intDelay & " Seconds"
            blnEnableTimer = True
        Else
            'hide the progress bar
            Me.Height = NO_DELAY_HEIGHT
        End If
        
        'display the appropiate text in the label
        Me.Caption = GetShutText(genmShutdownMethod)
        lblShut.Caption = GetShutText(genmShutdownMethod) & _
                          " the computer?"
    End With
    
    'display the form and keep it the top-most window
    Me.Visible = True
    Call StayOnTop(Me)
    Me.Show
    timShut.Enabled = blnEnableTimer
    Beep
    
    'move the mouse to the Yes button
    DoEvents
    Call MoveToButton
End Sub

Private Sub cmdNo_Click()
    'just unload the form
    Unload Me
End Sub

Private Sub cmdYes_Click()
    'shut down the computer
    timShut.Enabled = False
    Me.Hide
    Call DoShutMethod
    Unload Me
End Sub

Private Sub Form_Load()
    Set mmosMove = New clsMouse
    
    'set the form fonts to the current system default
    Call SetFormFontsToSystem(Me, FNT_MESSAGE)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'unconditionally cose the form
    mmosMove.CancelMouseMove
    Call NotOnTop(Me)
    Unload Me
    
    'flag to clear this form from memory
    gblnDoCleanUp = True
End Sub

Private Sub timShut_Timer()
    'This timer is only active once per load, for the
    'duration of DelayTime (default 15 seconds). Once
    'complete it will try to do the shutdown method
    
    Static lngStartTick As Long     'this will store the tick that the timer control was enabled at
    
    Dim lngRemaining As Long        'the number of ticks remaining
    
    'reset the starting tick if the timer is active for
    'the first time or if the time elapsed since the
    'previous Timer_Tick is greater than the amount of
    'time we are meant to be active for
    If (lngStartTick = 0) Or _
       ((GetTickCount - lngStartTick) > _
        (gudtDay(Today).intDelay * CLng(5000)) _
       ) Then
        
        'reset
        lngStartTick = GetTickCount
    End If
    
    'get the time remaining
    lngRemaining = cpbShut.Max - _
                   ((lngStartTick + _
                    (gudtDay(Today).intDelay * _
                     CLng(1000))) - _
                    GetTickCount)
    
    'update the progress bar
    cpbShut.Value = lngRemaining
    cpbShut.Caption = Abs(Int((lngRemaining / 1000) - _
                          gudtDay(Today).intDelay)) & _
                      " Seconds"
    
    If lngRemaining >= cpbShut.Max Then
        'disable the timer
        timShut.Enabled = False
        
        'reset if the timer is enabled again
        lngStartTick = 0
        
        'call the code to shut down the computer
        Call cmdYes_Click
    End If
End Sub
