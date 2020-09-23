VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIdleSystem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub StartTracking()
    'starts tracking
    Call SysStartTracking
End Sub

Public Sub StopTracking()
    'stops tracking
    Call SysStopTracking
End Sub

Public Function CurrentIdleTime() As Long
    'returnes the idle time in ticks
    CurrentIdleTime = SysCurrentIdleTime
End Function

Private Sub Class_Initialize()
    Call StartTracking
End Sub

Private Sub Class_Terminate()
    Call StopTracking
End Sub
