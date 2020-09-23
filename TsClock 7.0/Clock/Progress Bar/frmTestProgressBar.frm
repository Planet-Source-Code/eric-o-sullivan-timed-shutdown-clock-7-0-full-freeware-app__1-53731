VERSION 5.00
Object = "{ED415606-CCE4-11D6-B787-DED17DD29476}#1.0#0"; "Clock Progress Bar.ocx"
Begin VB.Form frmTestProgressBar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Testing Progress Bar"
   ClientHeight    =   972
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6852
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   972
   ScaleWidth      =   6852
   StartUpPosition =   3  'Windows Default
   Begin ctlClockProgressBar.ctlProgBar ctlProgBar2 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6855
      _ExtentX        =   12086
      _ExtentY        =   868
      Value           =   50
      Caption         =   "This is just testing the damn control"
      PercentCaption  =   0   'False
      BackColour      =   -2147483636
      FillColour      =   -2147483647
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ctlClockProgressBar.ctlProgBar ctlProgBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6160
      _ExtentY        =   868
      Value           =   50
      Appearance      =   0
      BackColour      =   -2147483643
      TextColour      =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   0
   End
End
Attribute VB_Name = "frmTestProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
    Static sngCounter As Single
    
    If sngCounter > ctlProgBar2.Max Then
        sngCounter = ctlProgBar2.Min
    End If
    ctlProgBar1.Value = sngCounter
    ctlProgBar2.Value = sngCounter
    sngCounter = sngCounter + 0.25
End Sub
