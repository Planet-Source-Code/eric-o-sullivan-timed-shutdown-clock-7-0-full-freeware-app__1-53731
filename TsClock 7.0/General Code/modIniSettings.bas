Attribute VB_Name = "modIniAccess"
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     20 Feburary 2002
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    Ini File Access Module
' -----------------------------------------------
'COMMENTS :
'This module is used to access an ini file, both
'to write data to one and to read data from one.
'=================================================

'require variable declaration
Option Explicit

'this module cannot be accessed outside this project
Option Private Module

'------------------------------------------------
'               API DECLARATIONS
'------------------------------------------------
'write a record to an ini file
Private Declare Function WritePrivateProfileString _
        Lib "kernel32" _
        Alias "WritePrivateProfileStringA" _
            (ByVal lpApplicationName As String, _
             ByVal lpKeyName As Any, _
             ByVal lpString As Any, _
             ByVal lpFileName As String) _
             As Long

'read a record from an ini file
Private Declare Function GetPrivateProfileString _
        Lib "kernel32" _
        Alias "GetPrivateProfileStringA" _
            (ByVal lpApplicationName As String, _
             ByVal lpKeyName As Any, _
             ByVal lpDefault As String, _
             ByVal lpReturnedString As String, _
             ByVal nSize As Long, _
             ByVal lpFileName As String) _
             As Long

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Public Sub WriteToIni(ByVal strFilePath As String, _
                      ByVal strHeading As String, _
                      ByVal strSetting As String, _
                      ByVal strValue As String)
    
    'This will write a string into the specified
    'ini file, under the specified heading. If the
    'ini file does not exist, then it is
    'automatically created.
    
    Dim lngResult As Long       'holds the return value of the WritePrivateProfileString function
    
    'insert a null character at the end of the
    'filepath before calling the api function
    strFilePath = strFilePath & vbNullChar
    
    'headings are automatically converted to
    'uppercase to conform with coding conventions
    strHeading = UCase(strHeading)
    
    'write the value to the ini file
    lngResult = WritePrivateProfileString(strHeading, _
                                          strSetting, _
                                          strValue, _
                                          strFilePath)
End Sub

Public Function GetFromIni(ByVal strFilePath As String, _
                           ByVal strHeading As String, _
                           ByVal strSetting As String) _
                           As String

    'This will read a specified value from under
    'the specified heading, in the specified ini
    'file.
    
    Const BUFFER_SIZE = 255
    
    Dim strBuffer As String * BUFFER_SIZE   'set up the buffer that holds the value returned
    Dim lngResult As Long                   'holds the return value of the GetPrivateProfileString fuction
    
    'initialise the buffer to remove unwanted
    'characters.
    strBuffer = String(BUFFER_SIZE, vbNullChar)
    
    'insert a null character at the end of the file
    'path before calling the api function.
    strFilePath = strFilePath & vbNullChar
    
    lngResult = GetPrivateProfileString(strHeading, _
                                        strSetting, _
                                        vbNullString, _
                                        strBuffer, _
                                        BUFFER_SIZE, _
                                        strFilePath)
    
    'return the result and exit
    If Left(strBuffer, 1) <> vbNullChar Then
        'there was something to return
        GetFromIni = Left(strBuffer, _
                          InStr(strBuffer, _
                                vbNullChar) _
                          - 1)
    End If
End Function

Public Function AddFile(ByVal strPath As String, _
                        ByVal strFileName As String) _
                        As String
    'This function takes a file name and a path and will
    'put the two together to form a filepath. This is useful
    'for when the applications' path happens to be the root
    'directory.
    
    'check the last character for a backslash
    If Left(strPath, 1) = "\" Then
        'don't insert a backslash
        AddFile = strPath & strFileName
    Else
        'insert a backslash
        AddFile = strPath & "\" & strFileName
    End If
End Function

Public Sub GetFileList(ByRef strFiles() As String, _
                       Optional ByVal strPath As String, _
                       Optional ByVal strExtention As String = "*.*", _
                       Optional ByVal lngAttributes As Long = vbNormal, _
                       Optional ByVal intNumFiles As Integer)
   'This procedure will get a list of files
    'available in the specified directory. If
    'no directory is specified, then the
    'applications directory is taken to be
    'the default.
    
    Dim intCounter  As Integer      'used to reference new elements in the array
    Dim strTempName As String       'temperorily holds a file name
    
    'validate the parameters for correct values
    If (Trim(strPath = "")) _
       Or (Dir(strPath, vbDirectory) = "") Then
        
        'invalid path, assume applications
        'directory
        strPath = App.Path
    End If
    
    'reset the array before entering new data
    ReDim strFiles(0)
    
    'resize the array to nothing if the
    'number of files specified is less
    'than can be returned
    If intNumFiles < 1 Then
        'return the maximum number of files (if possible)
        intNumFiles = 32767
    End If
    
    'include a wild card if the user only
    'specified the extention
    If Left(strExtention, 1) = "." Then
        strExtention = "*" & strExtention
    ElseIf InStr(strExtention, ".") = 0 Then
        strExtention = "*." & strExtention
    End If
    
    'get the first file name to start
    'the file search for this directory
    strTempName = Dir(AddFile(strPath, _
                              strExtention), _
                      lngAttributes)
    
    'keep getting new files until there are
    'no more to return
    Do While (strTempName <> "") _
       And (intCounter <= intNumFiles)
        
        'enter the element into the array
        ReDim Preserve strFiles(intCounter)
        strFiles(intCounter) = strTempName
        intCounter = intCounter + 1
        
        'get a new file
        strTempName = Dir
    Loop
End Sub

Public Function LimitRange(ByVal lngCheck As Long, _
                           Optional ByVal lngMin As Long = 0, _
                           Optional ByVal lngMax As Long = 100) _
                           As Long
    'This will make sure the value is between the valid
    'ranges. If the value is below, then the value is
    'changed to be Min. If the value is above, then the
    'value is changed to be Max.
    
    Select Case lngCheck
    Case Is > lngMax
        'value is above the maximum
        lngCheck = lngMax
        
    Case Is < lngMin
        'value is below the minimum
        lngCheck = lngMin
    End Select
    
    'return the result
    LimitRange = lngCheck
End Function
                           
Public Function GetFilePath(ByVal strFilePath As String, _
                            Optional ByVal blnReturnPath As Boolean = True) _
                            As String
    'This will return the path part of a filepath by default, but can be
    'set to return the file section of the path
    
    Dim intSlashPos     As Integer  'holds the position of the last backslash in the file path
    
    'ake sure we were passed a correct parameter
    If Trim(strFilePath) = "" Then
        GetFilePath = ""
        Exit Function
    End If
    
    'is the path specified already pointing to a directory
    If Dir(strFilePath, vbDirectory) <> "" Then
        If (GetAttr(strFilePath) And vbDirectory) And blnReturnPath Then
            'path is pointing to a directory, return full path
            GetFilePath = strFilePath
            Exit Function
        End If
    End If
    
    'return everything after the last backslash in the string to return
    'the path
    intSlashPos = InStrRev(strFilePath, "\")
    If intSlashPos = 0 Then
        'probably an invalid string, but could just be a drive letter, so
        'return full string
        If (Right(strFilePath, 1) = ":") And (Len(strFilePath) = 2) Then
            'assume a drive letter is referenced and add a backslash
            GetFilePath = strFilePath & "\"
        Else
            'unknown format - return whole string
            GetFilePath = strFilePath
        End If
        Exit Function
    End If
    
    'return everything before the last backslash
    If blnReturnPath Then
        'return the path section of the string
        Select Case intSlashPos
        Case Is > 3
            'return the path minus the backslash
            GetFilePath = Left(strFilePath, intSlashPos - 1)
        
        Case 2
            'only a drive letter in the string, specify the root directory
            'by leaving the backslash in
            GetFilePath = Left(strFilePath, intSlashPos)
            
        Case Else
            'there is something wrong
            GetFilePath = ""
        End Select
    Else
        'return the filename minus the backslash
        If intSlashPos = Len(strFilePath) Then
            'remove the blackslash at the end of the string
            GetFilePath = Left(strFilePath, intSlashPos - 1)
        Else
            'return everything after the backslash
            GetFilePath = Mid(strFilePath, intSlashPos + 1)
        End If
    End If
End Function
