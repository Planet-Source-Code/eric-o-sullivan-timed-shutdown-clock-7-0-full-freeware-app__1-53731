VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIdleDll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     16 December 2002
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    Idle Time Class (C++ IdleTime.dll)
' -----------------------------------------------
'COMMENTS :
'This will calculate the amount of idle time
'currently on the system. This works a bit like a
'screensaver.
'=================================================

'all variables must be declared
Option Explicit

'text comparisons are not case sensitive
Option Compare Text

'------------------------------------------------
'               API DECLARATIONS
'------------------------------------------------
'This will start the idle tracking
Private Declare Function TrackIdleTime _
        Lib "IdleTime" _
            () As Integer

'This will finish the idle tracking
Private Declare Sub EndTracking _
        Lib "IdleTime" ()

'this will get the current idle time
Private Declare Function GetIdleTime _
        Lib "IdleTime" _
            () As Long

'------------------------------------------------
'                    METHODS
'------------------------------------------------
Public Sub Start()
    'make sure that the idle time is not active
    
    Dim lngResult       As Long     'holds any returned error value from an api call
    
    Call EndTracking
    lngResult = TrackIdleTime
End Sub

Public Sub Finish()
    'make sure that we are not tracking the idle time any more
    Call EndTracking
End Sub

Public Function IdleTime() As Long
    'This will return the current idle time in milliseconds
    IdleTime = GetIdleTime
End Function

'------------------------------------------------
'               CLASS PROCEDURES
'------------------------------------------------
Private Sub Class_Initialize()
    'start tracking the idle time
    Call Start
End Sub

Private Sub Class_Terminate()
    'make sure that we are not tracking the idle time any more
    Call Finish
End Sub
