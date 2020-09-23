Attribute VB_Name = "modSystemTime"
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     2 September 1999
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    System Time Module
' -----------------------------------------------
'COMMENTS :
'This module will let the programmer set the time
'on the local machine.
'=================================================

'all variables must be declared
Option Explicit

'------------------------------------------------
'               API DECLARATIONS
'------------------------------------------------

'changes the local time to the specified time as long
'as the user has the privilages to do so
Private Declare Function SetLocalTime _
        Lib "kernel32.dll" _
            (lpSystemTime As SYSTEMTIME) _
             As Long

'------------------------------------------------
'               USER-DEFINED TYPES
'------------------------------------------------
Private Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Public Sub SetNewTime(ByVal intNewHour As Integer, _
                      ByVal intNewMinute As Integer, _
                      ByVal intNewSecond As Integer)
    ' Set the system time to the time specified
    
    Dim SetTime As SYSTEMTIME
    Dim RetVal  As Long
    
    SetTime.wHour = intNewHour
    SetTime.wMinute = intNewMinute
    SetTime.wSecond = intNewSecond
    SetTime.wMilliseconds = 0
    SetTime.wDay = Day(Date)
    SetTime.wMonth = Month(Date)
    SetTime.wYear = Year(Date)
    
    ' Set time and date.
    RetVal = SetLocalTime(SetTime)
End Sub
