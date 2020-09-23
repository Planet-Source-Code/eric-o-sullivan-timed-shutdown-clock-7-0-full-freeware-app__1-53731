VERSION 5.00
Begin VB.Form frmChangeTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Time"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2790
   Icon            =   "frmChangeTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set"
      Default         =   -1  'True
      Height          =   375
      Left            =   428
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1508
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   975
      TabIndex        =   2
      Top             =   480
      Width           =   1035
      Begin VB.TextBox txtHour 
         Alignment       =   2  'Center
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
         Height          =   240
         Left            =   0
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "0"
         ToolTipText     =   "Hours"
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtMin 
         Alignment       =   2  'Center
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
         Height          =   240
         Left            =   360
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "0"
         ToolTipText     =   "Minutes"
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtSec 
         Alignment       =   2  'Center
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
         Height          =   240
         Left            =   720
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Seconds"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblBreak1 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   7
         Top             =   0
         Width           =   135
      End
      Begin VB.Label lblBreak1 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.Timer timCurrent 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time (24H)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Time in 24 hours"
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "15:00:00   20 September 2002"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmChangeTime"
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
'TITLE :    Change Local Time Screen
' -----------------------------------------------
'COMMENTS :
'This screen is used to change the local time on a
'machine, provided that the user already has the
'privilages to do so
'=================================================

'all variables must be declared
Option Explicit

'------------------------------------------------
'               MODULE-LEVEL VARIABLES
'------------------------------------------------
Private mblnGotFocus    As Boolean  'used to stop updating the text boxes if the user has changed anything

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Private Sub cmdCancel_Click()
    'just unload the form - don't change anything
    Unload Me
End Sub

Private Sub cmdSet_Click()
    'set the new system time
    
    Dim intSetHour  As Integer
    Dim intSetMin   As Integer
    Dim intSetSec   As Integer
    
    'get each value from the text boxes
    intSetHour = Val(txtHour.Text)
    intSetMin = Val(txtMin.Text)
    intSetSec = Val(txtSec.Text)
    
    'set the local time
    Call SetNewTime(intSetHour, intSetMin, intSetSec)
End Sub

Private Sub Form_Activate()
    'highlight the text in the text box
    Call HighLight(txtHour)
End Sub

Private Sub Form_Load()
    'display the current time in the text boxes
    txtHour.Text = Hour(Time)
    txtMin.Text = Minute(Time)
    txtSec.Text = Second(Time)
    mblnGotFocus = False
    
    'make sure that whatever system setting, the time
    'is always centered in the label
    lblCurrent.Width = Me.ScaleWidth
    
    'set the form fonts to the current system default
    Call SetFormFontsToSystem(Me, FNT_MESSAGE)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'flag to clear this form from memory
    gblnDoCleanUp = True
End Sub

Private Sub timCurrent_Timer()
    'display the current time
    
    Static intOldSecond As Integer  'used to only update the label when necessary
    
    Dim strCurrTime     As String   'holds the date and time to be dispalyed
    
    'do we have to update the display
    If intOldSecond <> Val(Second(Time)) Then
        If Not mblnGotFocus Then
            'display the current time in the text boxes
            txtHour.Text = FormatTime(Hour(Time))
            txtMin.Text = FormatTime(Minute(Time))
            txtSec.Text = FormatTime(Second(Time))
            mblnGotFocus = False
        End If
        
        'display time
        intOldSecond = Second(Time)
        strCurrTime = Time & "   " & _
                      Format(Date, "Long date")
        lblCurrent.Caption = strCurrTime
    End If
End Sub

Private Sub txtHour_Change()
    'assume the user is going to try and change the time
    'so disable the update of the time in the text boxes
    mblnGotFocus = True
End Sub

Private Sub txtHour_GotFocus()
    'highlight the text in the text box
    Call HighLight(txtHour)
End Sub

Private Sub txtHour_Validate(Cancel As Boolean)
    'make sure that the user entered a valid time (0-23)
    With txtHour
        .Text = FormatTime(modNumbers.LimitRange(Val(.Text), 0, 23))
    End With
End Sub

Private Sub txtMin_Change()
    'assume the user is going to try and change the time
    'so disable the update of the time in the text boxes
    mblnGotFocus = True
End Sub

Private Sub txtMin_GotFocus()
    'highlight the text in the text box
    Call HighLight(txtMin)
End Sub

Private Sub txtMin_Validate(Cancel As Boolean)
    'make sure that the user entered a valid time (0-59)
    With txtMin
        .Text = FormatTime(modNumbers.LimitRange(Val(.Text), 0, 59))
    End With
End Sub

Private Sub txtSec_Change()
    'assume the user is going to try and change the time
    'so disable the update of the time in the text boxes
    mblnGotFocus = True
End Sub

Private Sub txtSec_GotFocus()
    'highlight the text in the text box
    Call HighLight(txtSec)
End Sub

Private Sub txtSec_Validate(Cancel As Boolean)
    'make sure that the user entered a valid time (0-59)
    With txtSec
        .Text = FormatTime(modNumbers.LimitRange(Val(.Text), 0, 59))
    End With
End Sub

Private Function FormatTime(ByVal intNum As Integer) _
                            As String
    'This will make sure that the time section will
    'always be two digits
    
    Dim strResult As String * 2     'we only return the first two characters
    
    strResult = Format(Trim(Str(intNum)), "00")
    FormatTime = strResult
End Function
