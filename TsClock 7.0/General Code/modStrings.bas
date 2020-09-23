Attribute VB_Name = "modStrings"
'-------------------------------------------------------------------------------
'                                MODULE DETAILS
'-------------------------------------------------------------------------------
'   Program Name:   General Use
'  ---------------------------------------------------------------------------
'   Author:         Eric O'Sullivan
'  ---------------------------------------------------------------------------
'   Date:           07 July 2002
'  ---------------------------------------------------------------------------
'   Company:        CompApp Technologies
'  ---------------------------------------------------------------------------
'   Contact:        DiskJunky@hotmail.com
'  ---------------------------------------------------------------------------
'   Description:    This will perform various manipulations with strings.
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

'require variable declaration
Option Explicit

'-------------------------------------------------------------------------------
'                              API DECLARATIONS
'-------------------------------------------------------------------------------
'this can convert entire type structures
'to other types like a Long
Private Declare Sub RtlMoveMemory _
        Lib "kernel32.dll" _
            (Destination As Any, _
             Source As Any, _
             ByVal Length As Long)

'-------------------------------------------------------------------------------
'                                 PROCEDURES
'-------------------------------------------------------------------------------
Public Function AddFile(ByVal strDirectory As String, _
                        ByVal strFileName As String) _
                        As String
    'This will add a file name to a directory path to
    'create a full filepath.
    
    If Right(strDirectory, 1) <> "\" Then
        'insert a backslash
        strDirectory = strDirectory & "\"
    End If
    
    'append the file name to the directory path now
    AddFile = strDirectory & strFileName
End Function

Public Function BinarySearch(ByRef strOrdered() As String, _
                             ByVal strFind As String, _
                             Optional ByRef lngFoundAt As Long, _
                             Optional ByVal strMatchFullString As Boolean = True) _
                             As Boolean
    'This will perform a binary search on the ordered array of strings. The
    'function will return True if the specified string was found, and if
    'the lngFoundAt parameter was specified, it will hold the position at
    'which the string was found at. If the string was not found the parameter
    'lngFoundAt will point to the first element in the array
    
    Dim lngStart        As Long         'holds the lower bound of the search
    Dim lngFinish       As Long         'holds the upper bound of the search
    Dim blnFound        As Boolean      'holds whether or not the element was found
    
    'make sure that there is something to find
    If (strFind = "") Then
        Exit Function
    End If
    
    'initialise the variables for the search
    lngStart = LBound(strOrdered)
    lngFinish = UBound(strOrdered)
    
    Do While (lngStart <= lngFinish) And _
             (Not blnFound)
        'get the half way point
        lngFoundAt = ((lngFinish - lngStart) \ 2) + lngStart
        
        If ((lngFinish - lngStart) > 1) Then
            'does this match the string
            Select Case strOrdered(lngFoundAt)
            Case strFind
                blnFound = True
                
            Case Is > strFind
                'move the upper bound to this point
                lngFinish = lngFoundAt - 1
                
            Case Is < strFind
                'move the lower bound to this point
                lngStart = lngFoundAt + 1
            End Select
        
        Else
            'there are only two elements left to compare
            Select Case strOrdered(lngFoundAt)
            Case strFind
                blnFound = True
                
            Case Else
                'check the other element
                If (strOrdered(lngFinish) = strFind) Then
                    blnFound = True
                    lngFoundAt = lngFinish
                End If
            End Select
            Exit Do
        End If
    Loop
    
    'Debug.Assert (strFind <> "A")
    
    'return the result
    If Not blnFound Then
        'make sure that this points at nothing
        lngFoundAt = LBound(strOrdered)
    End If
    BinarySearch = blnFound
End Function

Public Function CommaCount(ByVal strLine As String) _
                           As Integer
    'This will return the number of commas foun in the string. Mainly
    'use to find the number of variables declared on the same line
    
    CommaCount = CountString(strLine, ",")
End Function

Public Function CountString(ByVal strCheck As String, _
                            ByVal strFind As String, _
                            Optional ByVal enmCompare As VbCompareMethod = vbTextCompare) _
                            As Long
    'This will count how many times a string occurrs in another string
    
    Dim strRemoved      As String   'holds the string with the specified characters removed from it
    
    'make sure that the user passed valid strings
    If ((strCheck = "") Or (strFind = "")) Then
        Exit Function
    End If
    
    'remove all instances of the specified string
    strRemoved = Replace(strCheck, strFind, "", compare:=enmCompare)
    
    'return the number of times that the string was found
    CountString = (Len(strCheck) - Len(strRemoved)) / Len(strFind)
End Function

Public Function GetAfter(ByVal strSentence As String, _
                         Optional ByVal strCharacter As String = "=") _
                         As String
    'This procedure returns all the character of a
    'string after the "=" sign.
    
    Dim intCounter As Integer
    Dim strRest As String
    Dim strSign As String
    
    strSign = strCharacter
    
    'find the last position of the specified sign
    intCounter = InStrRev(strSentence, strSign)
    
    If intCounter <> Len(strSentence) Then
        strRest = Right(strSentence, (Len(strSentence) - (intCounter + Len(strCharacter) - 1)))
    Else
        strRest = ""
    End If
    
    GetAfter = strRest
End Function

Public Function GetBefore(ByVal strSentence As String, _
                          Optional ByVal strSign As String = "=") _
                          As String
    'This procedure returns all the character of a
    'string before the "=" sign.
    
    Dim intCounter As Integer
    Dim strBefore As String
    
    'find the position of the equals sign
    intCounter = InStr(1, strSentence, strSign)
    
    If (intCounter <> Len(strSentence)) And (intCounter <> 0) Then
        strBefore = Left(strSentence, (intCounter - 1))
    Else
        strBefore = strSentence
    End If
    
    GetBefore = strBefore
End Function

Public Sub GetFileList(ByRef strFiles() As String, _
                       Optional ByVal strPath As String, _
                       Optional ByVal strExtention As String = "*.*", _
                       Optional ByVal enmAttributes As VbFileAttribute = vbNormal, _
                       Optional ByVal lngNumFiles As Long, _
                       Optional ByVal blnSearchSubDir As Boolean = False, _
                       Optional ByVal blnAddToExistingList As Boolean = False)
    'This procedure will get a list of files available in the specified
    'directory. If no directory is specified, then the applications directory
    'is taken to be the default.
    'The parameter blnAddToExistingList will not reset the strFiles array, but
    'will instead add the file list to the array
    
    Dim lngCounter      As Long         'used to reference new elements in the array
    Dim strTempName     As String       'temperorily holds a file name
    Dim strSubDirs()    As String       'holds a list of sub directories to scan
    Dim blnIsDir        As Boolean      'flags if the current file is a directory
    
    'validate the parameters for correct values
    If (Trim(strPath = "")) _
       Or (Dir(strPath, vbDirectory) = "") Then
        
        'invalid path, assume applications
        'directory
        strPath = App.Path
    End If
    
    'make sure that the specified path was a directory
    If ((GetAttr(strPath) And vbDirectory) <> vbDirectory) Then
        Exit Sub
    End If
    
    If Not blnAddToExistingList Then
        'reset the array before entering new data
        ReDim strFiles(0)
    Else
        'set the counter
        If (strFiles(0) <> "") Then
            lngCounter = UBound(strFiles) + 1
        End If
    End If
    
    'if no number of files was specified, then assume maximum
    If (lngNumFiles = 0) Then
        lngNumFiles = 2147483647    '2,147,483,647
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
                      enmAttributes)
    
    'keep getting new files until there are
    'no more to return
    Do While (strTempName <> "") _
       And (lngCounter <= lngNumFiles)
        
        'if we are scanning directories, then ignore "." and "..",
        'otherwise assume that the file exists and has at least
        'one matching attribute
        blnIsDir = ((enmAttributes And vbDirectory) = vbDirectory)
        If (blnIsDir And _
            (Trim(strTempName) <> ".") And _
            (Trim(strTempName) <> "..")) Or _
           (((enmAttributes Or enmAttributes) >= 0) And _
            (Not blnIsDir)) Then
            
            'make sure that the file has the attributes we are looking for
            If (((GetAttr(AddFile(strPath, strTempName)) Or enmAttributes) > 0) Or _
                 (enmAttributes = 0)) Or _
               ((Not blnIsDir) And _
                (strTempName Like strExtention)) Then
                'enter the file into the array
                
                ReDim Preserve strFiles(lngCounter)
                strFiles(lngCounter) = AddFile(strPath, strTempName)
                lngCounter = lngCounter + 1
            End If
        End If
        
        'get a new file
        strTempName = Dir
        DoEvents
    Loop
    
    'are we meant to search sub directories
    If blnSearchSubDir Then
        Call GetFileList(strSubDirs(), _
                         strPath, _
                         , _
                         vbDirectory)
        
        'get a list of files for each sub directory
        For lngCounter = 0 To (UBound(strSubDirs))
            If (Trim(strSubDirs(lngCounter)) <> "") Then
                Call GetFileList(strFiles(), _
                                 strSubDirs(lngCounter), _
                                 strExtention, _
                                 enmAttributes, _
                                 lngNumFiles, _
                                 True, _
                                 True)
            End If
        Next lngCounter
    End If  'search sub directories
End Sub

Public Function GetFilePath(ByVal strFilePath As String, _
                            Optional ByVal blnReturnPath As Boolean = True) _
                            As String
    'This will return the path part of a filepath by default, but can be
    'set to return the file section of the path
    
    Dim intSlashPos        As Integer  'holds the position of the last backslash in the file path
    
    'make sure we were passed a correct parameter
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
            GetFilePath = strFilePath + "\"
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
        
        Case 3
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

Public Function GetLine(ByVal strText As String, _
                        ByVal lngFromPoint As Long)
    'This will return a line of text from multiple lines from a specified
    'character position
    
    Dim lngStart        As Long         'holds the start of the line
    Dim lngEnd          As Long         'holds the end of the line
    Dim strBuffer       As String       'holds the text returned from the function
    
    If (lngFromPoint < 1) Or (strText = "") Then
        'invalid parameters
        Exit Function
    End If
    
    'get the start of the line
    lngStart = InStrRev(strText, vbCrLf, lngFromPoint)
    If (lngStart = 0) Then
        lngStart = 1
    End If
    
    'get the end of the line
    lngEnd = InStr(lngFromPoint, strText, vbCrLf)
    If (lngEnd = 0) Then
        lngEnd = Len(strText) + 1
    End If
    
    'return the line minus the line feed and carrage returns
    strBuffer = Mid(strText, lngStart, (lngEnd - lngStart))
    strBuffer = Replace(strBuffer, Chr(10), "")
    strBuffer = Replace(strBuffer, Chr(13), "")
    GetLine = strBuffer
End Function

Public Function GetMode(ByVal strText As String) As String
    'This will return the most occuring character
    
    Dim intChars()      As Integer      'holds the string in integer form
    Dim lngTextLen      As Long         'holds the number of characters in the string
    Dim intUnique()     As Integer      'holds each unique character in the string
    Dim lngUnFound      As Long         'holds the number of unique characters found
    Dim lngCounterUn    As Long         'used for cycling through the intUnique array
    Dim intCount()      As Integer      'holds the number of times each character occurs in the string (elements directly relate to elements in intUnique)
    Dim lngCounter      As Long         'used for cycling through the intChars() array
    Dim intMode         As Integer      'holds the array element of the most occuring character
    Dim intMax          As Integer      'holds the maximum count of the specified character
    Dim intCharTest     As Integer      'holds a single character to test
    Dim blnFound        As Boolean      'holds if the character to test alreay exists in the array
    
    'make sure that something was passed
    lngTextLen = Len(strText)
    If (lngTextLen = 0) Then
        GetMode = ""
        Exit Function
    End If
    
    'convert the string to an integer array (the upperbound of the array will
    'be the number of characters - 1)
    intChars() = StringToInt(strText)
    
    'resize the arrays to match the string (potentially all characters are unique
    ReDim intUnique(lngTextLen - 1)
    ReDim intCount(lngTextLen - 1)
    lngUnFound = 0
    
    'search through the text
    For lngCounter = 0 To (lngTextLen - 1)
        'get the value to test
        intCharTest = intChars(lngCounter)
        
        'does this value exist in the array
        blnFound = False
        For lngCounterUn = 0 To (lngUnFound - 1)
            If (intUnique(lngCounterUn) = intCharTest) Then
                'this character already exists in the array, increase the
                'character count
                intCount(lngCounterUn) = intCount(lngCounterUn) + 1
                blnFound = True
                Exit For
            End If
        Next lngCounterUn
        
        If Not blnFound Then
            'add the character to the array
            intUnique(lngUnFound) = intCharTest
            intCount(lngUnFound) = 1
            lngUnFound = lngUnFound + 1
        End If
    Next lngCounter
    
    'find the most occuring character by checking the maximum element in the
    'lngCount array
    intMax = 0      'assume "lowest" character as a starting point
    For lngCounter = 0 To (lngUnFound - 1)
        If (intCount(lngCounter) > intMax) Then
            'new highest count
            intMax = intCount(lngCounter)
            
            'remember the position of this character so that we know which on
            'it is
            intMode = lngCounter
        End If
    Next lngCounter
    
    'return the result
    GetMode = Chr$(intUnique(intMode))
End Function

Public Function GetPath(ByVal strAddress As String) _
                        As String
    'This function returns the path from a string containing the full
    'path and filename of a file.
    
    Dim intLastPos As Integer
    
    'find the position of the last "\" mark in the string
    intLastPos = InStrRev(strAddress, "\")
    
    'if no \ found in the string, then
    If intLastPos = 0 Then
        'return the whole string
        intLastPos = Len(strAddress) + 1
    End If
    
    'return everything before the last "\" mark
    GetPath = Left(strAddress, (intLastPos - 1))
End Function

Public Function IsNotInQuote(ByVal strText As String, _
                             ByVal strWords As String, _
                             Optional ByVal lngWordPos As Long = 0) _
                             As Boolean
    'This function will tell you if the specified text is in quotes within
    'the second string. It does this by counting the number of quotation
    'marks before the specified strWords. If the number is even, then the
    'strWords are not in qototes, otherwise they are.
    
    'the quotation mark, " , is ASCII character 34
    
    Dim lngGotPos As Long
    Dim lngCounter As Long
    Dim lngNextPos As Long
    
    'was the position of the work specified
    If (lngWordPos > 0) Then
        'where does the word in the string occur
        lngGotPos = lngWordPos
        
    Else
        'find where the position of strWords in strText
        lngGotPos = InStr(1, strText, strWords)
        If lngGotPos = 0 Then
            IsNotInQuote = True
            Exit Function
        End If
    End If
    
    'start counting the number of quotation marks
    lngNextPos = 0
    Do
        lngNextPos = InStr(lngNextPos + 1, strText, Chr(34))
        
        If (lngNextPos <> 0) And (lngNextPos < lngGotPos) Then
            'quote found, add to total
            lngCounter = lngCounter + 1
        End If
    Loop Until (lngNextPos = 0) Or (lngNextPos >= lngGotPos)
    
    'no quotes at all found
    If lngCounter = 0 Then
        IsNotInQuote = True
        Exit Function
    End If
    
    'if the number of quotes is even, then return true, else return false
    If lngCounter Mod 2 = 0 Then
        IsNotInQuote = True
    End If
End Function

Public Function IsWord(ByVal strLine As String, _
                       ByVal strWord As String, _
                       Optional ByVal lngWordAtPos As Long = 0) _
                       As Boolean
    'This function will return True if the
    'specified word is not part of another
    'word
    
    Dim blnLeftOk As Boolean    'the left side of the word is valid
    Dim blnRightOk As Boolean   'the right side of the word is valid
    Dim lngWordPos As Long      'the position of the specified word in the string
    
    If (Len(strWord) > Len(strLine)) _
       Or (strLine = "") _
       Or (strWord = "") Then
        'invalid parameters
        IsWord = False
        Exit Function
    End If
    
    'remove leading/trailing spaces
    strLine = strLine
    strWord = strWord
    
    'get the position of the word in the line
    If (lngWordAtPos < 1) Or (lngWordAtPos >= Len(strLine)) Then
        lngWordPos = InStr(UCase(strLine), UCase(strWord))
    Else
        lngWordPos = lngWordAtPos
    End If
    
    If lngWordPos = 0 Then
        'word not found
        IsWord = False
        Exit Function
    End If
    
    'check the left side of the word
    If lngWordPos = 1 Then
        'word is on the left side of the string
        blnLeftOk = True
    Else
        'check the character to the left of the word
        Select Case UCase(Mid(strLine, lngWordPos - 1, 1))
        Case "A" To "Z", "0" To "9", "_"
        Case Else
            blnLeftOk = True
        End Select
    End If
    
    'check the right side of the word
    If ((lngWordPos + Len(strWord)) - 1) = Len(strLine) Then
        'word is on the left side of the string
        blnRightOk = True
    Else
        'check the character to the left of the word
        'Debug.Print strWord, UCase(Mid(strLine, lngWordPos + Len(strWord), 1))
        Select Case UCase(Mid(strLine, lngWordPos + Len(strWord), 1))
        Case "A" To "Z", "0" To "9", "_"
            'Stop
        Case Else
            blnRightOk = True
        End Select
    End If
    
    'if both sides are OK, then return True
    If blnLeftOk And blnRightOk Then
        IsWord = True
    End If
End Function

Public Function LineCount(ByVal strLines As String) _
                          As Long
    'This will return the number of lines of text in the string
    
    If (strLines <> "") Then
        LineCount = CountString(strLines, vbCrLf) + 1
    End If
End Function

Public Function MulString(ByVal lngNumber As Long, _
                          ByVal strText As String) _
                          As String
    'This will add the passed string onto itself for the specified number of
    'times
    
    Dim strResult   As String       'holds the final string to be returned
    Dim lngCounter  As Long         'used to cycle through the number of times to add onto the return string
    
    'add onto the string for the specified number of times and return the result
    For lngCounter = 1 To lngNumber
        strResult = strResult + strText
    Next lngCounter
    MulString = strResult
End Function

Public Function PaddString(ByVal strText As String, _
                           ByVal lngTotalLen As Long, _
                           Optional enmAlign As AlignmentConstants = vbLeftJustify) _
                           As String
    'This will padd a string with trailing spaces so that the returned string
    'matches the total number of characters specified. If the string passed is
    'bigger than the total number of characters, then the string is truncated
    'and then returned.
    
    Dim lngTextLen As Long  'the length of the text string passed
    
    'if the number of characters is zero, then
    'return nothing
    If lngTotalLen = 0 Then
        PaddString = ""
        Exit Function
    End If
    
    'remove null characters
    strText = Replace(strText, vbNullChar, " ")
    
    'get the length of the string
    lngTextLen = Len(strText)
    
    If lngTextLen >= lngTotalLen Then
        'return a trucated string
        PaddString = Left(strText, lngTotalLen)
    Else
        Select Case enmAlign
        Case vbLeftJustify
            'padd the string out with spaces on the right side of the string
            PaddString = strText + Space(lngTotalLen - lngTextLen)
            
        Case vbCenter
            'padd only half the number of spaces on the left, and half on
            'the right
            PaddString = Space((lngTotalLen - lngTextLen) \ 2) + strText
            PaddString = PaddString + Space(lngTotalLen - Len(PaddString))
        
        Case vbRightJustify
            'padd spaces on the left side of the string
            PaddString = Space(lngTotalLen - lngTextLen) + strText
        End Select
    End If
End Function

Public Function ParseWordByCaps(ByVal strLine As String, _
                                Optional ByRef lngWordsFound As Long = 0) _
                                As String()
    'This function will parse a given peice of text into several words by
    'assuming that each word in the text begins with a capital letter. Any
    'non-text is also assumed to seperate words. Eg, "HelloThere" becomes
    '"Hello" and "There".
    
    Dim strWords()      As String       'holds the list of words extracted from the line
    Dim lngNumWords     As Long         'holds the number of words found
    Dim lngCounter      As Long         'used to cycle through the text in the line
    Dim strChar         As String * 1   'holds a single character to check
    Dim lngTextLen      As Long         'holds the length of the text passed after it was validate
    Dim lngStartPos     As Long         'holds the starting position of the word in the string
    
    'initialise the variables
    lngNumWords = 0
    lngStartPos = 1
    ReDim strWords(lngNumWords)
    
    'validate the text passed
    strLine = Trim(strLine)
    If (Len(strLine) = 0) Then
        'no string was passed
        ParseWordByCaps = strWords()
        Exit Function
    End If
    
    'get the length of the text
    lngTextLen = Len(strLine)
    
    'cycle through the text
    For lngCounter = 1 To lngTextLen
        'get a character from the string
        strChar = Mid(strLine, lngCounter, 1)
        
        'what kind of character is this
        Select Case strChar
        Case "a" To "z"     'make sure that the previous character is not a space
            If (lngCounter > 1) Then
                If (Mid(strLine, lngCounter - 1, 1) = " ") Then
                    'start a new word here
                    Mid(strLine, lngCounter, 1) = UCase(strChar)
                    lngCounter = lngCounter - 1
                End If
            End If
        
        Case "A" To "Z"     'enter the new word
            'enter a new word into the array
            If (lngCounter > 1) Then
                ReDim Preserve strWords(lngNumWords)
                strWords(lngNumWords) = Trim(Mid(strLine, _
                                                 lngStartPos, _
                                                 (lngCounter - lngStartPos)))
                
                lngNumWords = lngNumWords + 1
                lngStartPos = lngCounter
            End If
            
        Case Else           'replace this with a space to ignore this character
            Mid(strLine, lngCounter, 1) = " "
        End Select
    Next lngCounter
    
    'enter the last word in the string
    ReDim Preserve strWords(lngNumWords)
    strWords(lngNumWords) = Trim(Mid(strLine, _
                                     lngStartPos, _
                                     (lngCounter - lngStartPos)))
    
    lngNumWords = lngNumWords + 1
    
    'return the list of words and the number found
    lngWordsFound = lngNumWords
    ParseWordByCaps = strWords()
End Function

Public Sub QSortStrings(ByRef strArray() As String, _
                        Optional ByVal lngStart As Long = -1, _
                        Optional ByVal lngFinish As Long = -1)
    'This will sort the string array within the bounds specified using the
    'QuickSort method.
    
    Dim lngUBound       As Long         'holds the upperbound of the array
    Dim lngLBound       As Long         'holds the lower bound of the array
    Dim lngTempHi       As Long         'temperorily holds the location of a string that should be moved up
    Dim lngTempLo       As Long         'temperorily holds the lcoation of a string that should be moved down
    Dim strTemp         As String       'holds a string as it is being swapped between two array elements
    Dim strPivot        As String       'holds a string to compare the current array elemnts with. This pivot should always be between to array positions
    
    'get the size of the array
    lngUBound = UBound(strArray)
    lngLBound = LBound(strArray)
    
    'set the defaults
    If (lngStart = -1) Then
        'default to the start of the array
        lngStart = lngLBound
    End If
    If (lngFinish = -1) Then
        'default to the end of the array
        lngFinish = lngUBound
    End If
    
    'initialise the start of the sort
    lngTempLo = lngStart
    lngTempHi = lngFinish
    strPivot = strArray((lngTempLo + lngTempHi) \ 2)
    
    'make a single pass of the array
    Do Until (lngTempLo > lngTempHi)
        'find the first element that is greater than the pivot string from
        'the current "low" position in the array
        Do While (strArray(lngTempLo) < strPivot)
            'check the next element
            lngTempLo = lngTempLo + 1
        Loop
        
        'find the first element that is less than the pivot string from the
        'current "hi" position in the array
        Do While (strArray(lngTempHi) > strPivot)
            lngTempHi = lngTempHi - 1
        Loop
        
        'did we find any elements that needed sorting
        If (lngTempLo <= lngTempHi) Then
            'swap the two values
            strTemp = strArray(lngTempLo)
            strArray(lngTempLo) = strArray(lngTempHi)
            strArray(lngTempHi) = strTemp
            
            'check the rest of the elements
            lngTempLo = lngTempLo + 1
            lngTempHi = lngTempHi - 1
        End If
    Loop
    
    'the array has only been sorted if the "Hi" "Lo" are at the Start and
    'Finish positions respectively
    If (lngTempHi > lngStart) Then
        Call QSortStrings(strArray(), lngStart, lngTempHi)
    End If
    If (lngTempLo < lngFinish) Then
        Call QSortStrings(strArray(), lngTempLo, lngFinish)
    End If
End Sub

Public Function RemoveDuplicates(ByVal strText As String, _
                                 Optional ByVal strRemove As String = " ") _
                                 As String
    'This will remove all double instances of the specified string and
    'will return the result
    
    Dim strDuplicate    As String       'holds a duplicate instance of the string to remove
    
    strDuplicate = MulString(2, strRemove)
    Do While (InStr(1, strText, strDuplicate) > 0)
        'remove double instances of the text
        strText = Replace(strText, strDuplicate, strRemove)
    Loop
    
    'return the filtered string
    RemoveDuplicates = strText
End Function

Public Sub RemoveDuplicateStrings(ByRef strArray() As String, _
                                  Optional ByVal blnRemoveBlanks As Boolean = True)
    'This procedure will remove any duplicate entries from the specified array.
    'The array WILL be sorted to remove duplicates in one array pass.
    
    Dim lngCounter      As Long         'used to cycle through the array
    Dim lngMin          As Long         'holds the lower bound of the array
    Dim lngMax          As Long         'holds the upper bound of the array
    Dim lngNumDel       As Long         'holds the number of elements deleted from the array
    
    'get the size of the array
    lngMin = LBound(strArray)
    lngMax = UBound(strArray)
    
    'make sure that there is more than one element in the array
    If ((lngMax - lngMin) < 0) Then
        Exit Sub
    End If
    
    'make sure that the array is sorted
    Call QSortStrings(strArray)
    
    'scan through the array
    For lngCounter = (lngMin + 1) To lngMax
        'have we scanned through all the elements (not including the deleted
        'ones)
        If (lngCounter > (lngMax - lngNumDel)) Then
            Exit For
        End If
        
        'copy the next element down skipping deleted elements
        strArray(lngCounter) = strArray(lngCounter + lngNumDel)
        
        'do we remove this element
        If (strArray(lngCounter) = strArray(lngCounter - 1)) Or _
           ((Len(strArray(lngCounter)) = 0) And _
            (blnRemoveBlanks)) Then
            
            'remove this element by scanning it again. The code above will
            'copy the next element to check above it
            lngNumDel = lngNumDel + 1
            lngCounter = lngCounter - 1
        End If  'do we remove this element
    Next lngCounter
    
    'resize the array to remove the elements
    If (lngNumDel > 0) Then
        'make sure that we don't resize the array smaller than possible
        If ((lngMax - lngNumDel) < 0) Then
            'wipe the array
            ReDim strArray(lngMin To lngMin)
            strArray(lngMin) = ""
        Else
            'remove all deleted elements
            ReDim Preserve strArray(lngMin To (lngMax - lngNumDel))
        End If
    End If
End Sub

Public Function SameChar(ByVal strText As String, _
                         ByVal strChar As String, _
                         Optional ByVal lngStart As Long = 1, _
                         Optional ByVal enmCompare As VbCompareMethod = vbBinaryCompare) _
                         As Boolean
    'This will test if the string is completely made up of the specified
    'characters.
    
    Dim lngTextLen      As Long         'holds the length of the original text
    
    'make sure that the parameters are correct
    lngTextLen = Len(strText)
    If (lngTextLen = 0) Or _
       (Len(strChar) = 0) Or _
       ((lngStart < 1) Or _
        (lngStart > lngTextLen)) Then
        
        'invalid parameters
        SameChar = False
        Exit Function
    End If
    
    'if this produces an empty string, then the pattern matched completely
    strText = Replace(strText, strChar, "", lngStart, , enmCompare)
    
    'returns True or False
    SameChar = (Len(strText) = 0)
End Function

Public Function ScrollString(ByVal strText As String, _
                             ByVal lngScrollFrom As Long, _
                             ByVal lngDisplayLength As Long, _
                             Optional ByVal lngCycleGap As Long = 0) _
                             As String
    'this will return a string that is part of the original string based
    'on the scroll values. Eg, if the original text is "Hello There" and
    'the display length is 5, and we are currently scrolling from the
    '3rd character (the default cycle gap, ie the gap between the end of
    'the original string and the start of a new scroll is the display
    'length), then the result is "llo t". From the 4th scroll position,
    'the result is "lo th" and so on.
    
    Dim strResult       As String       'holds the string returned from the function
    Dim lngLenText      As Long         'holds the length of the original text
    Dim lngLenResult    As Long         'holds the length of the text to be returned
    Dim lngCharsLeft    As Long         'holds the number of characters left to padd out before returning the result to the user
    
    'are we able to return anything
    If (lngDisplayLength < 1) Then
        Exit Function
    End If
    
    'make sure that a string was passed
    lngLenText = Len(strText)
    If (lngLenText = 0) Then
        'padd out to the return length
        ScrollString = Space$(lngDisplayLength)
        Exit Function
    End If
    
    'was a cycle gap specified
    If (lngCycleGap < 1) Then
        'use default
        lngCycleGap = lngDisplayLength
    End If
    
    'make sure we don't scroll past the end of the string
    lngScrollFrom = lngScrollFrom Mod (lngLenText + lngCycleGap + 1)
    
    'build the return string
    strResult = Mid$(strText, lngScrollFrom, lngDisplayLength)
    
    'do we need to padd out the result
    lngLenResult = Len(strResult)
    If (lngLenResult < lngDisplayLength) Then
        'we need to padd the gap between the end of the original text and the
        'start of the next scroll of it.
        lngCharsLeft = (lngDisplayLength - lngLenResult)
        If (lngCharsLeft < lngCycleGap) Then
            'padd out the result with spaces
            strResult = strResult + Space$(lngCharsLeft)
            
        Else
            'fill out as much as we can with spaces
            strResult = strResult + Space$(lngCycleGap - (lngScrollFrom - lngLenText))
            lngCharsLeft = lngDisplayLength - Len(strResult)
            strResult = strResult + Left(strText, lngCharsLeft)
        End If
    End If  'do we need to padd out the reuslt
    
    'return the string
    ScrollString = strResult
End Function

Public Function StringToInt(ByVal strText As String) _
                            As Integer()
    'convert a string to an array of integers. This is easier to manage
    'sometimes than picking out individual character values. The values
    'in the array are usually ascii values.
    
    Dim intText()   As Integer      'holds the integer version of the text
    Dim bytText()   As Byte         'holds the byte version of the text
    
    'vb byte conversion
    bytText() = strText
    
    'resize integer array
    ReDim intText(UBound(bytText) \ 2)
    
    'copy data
    Call RtlMoveMemory(intText(0), bytText(0), UBound(bytText))
    
    'return the array
    StringToInt = intText()
End Function

Public Function IntToString(ByRef intText() As Integer) As String
    'This is the reverse of the StringToInt function in that it takes an
    'integer array and converts it into a string.
    
    Dim bytText()       As Byte     'holds the byte form of the text
    Dim lngSize         As Long     'holds the size in bytes to be copied
    
    'copy the array to bytes
    lngSize = ((UBound(intText) + 1) * 2)
    ReDim bytText(lngSize - 1)
    
    'copy the integer array values to byte values
    Call RtlMoveMemory(bytText(0), intText(0), lngSize)
    
    'convert bytes to string
    IntToString = bytText()
End Function

Public Function StripBetween(ByVal strText As String, _
                             Optional ByVal strStartSign As String = """", _
                             Optional ByVal strEndSign As String = """") _
                             As String
    'This function will remove all text between the
    'specified marks (start and end.. by default; " )
    
    Dim lngQuoteStart As Long       'the position of the first quotation mark found in the string
    Dim lngQuoteFinish As Long      'the position of the quote mark after the first position
    
    'get the position of a quotation mark
    lngQuoteStart = InStr(strText, strStartSign)
    
    Do While (lngQuoteStart > 0)
        'find the next quote mark after the found position
        lngQuoteFinish = InStr(lngQuoteStart + Len(strStartSign), _
                               strText, _
                               strEndSign)
        
        'if a second quotation mark was found, remove
        'all text between
        If lngQuoteFinish > 0 Then
            strText = Left(strText, _
                           lngQuoteStart - Len(strStartSign)) & _
                      Right(strText, _
                            Len(strText) - lngQuoteFinish)
        Else
            'there is only one quotation mark. Remove it
            strText = Left(strText, _
                           lngQuoteStart - Len(strStartSign)) & _
                      Right(strText, _
                            Len(strText) - lngQuoteStart)
        End If
        
        'get the next occurance of a quotation mark
        lngQuoteStart = InStr(lngQuoteStart, _
                              strText, _
                              strStartSign)
    Loop
    
    'return the stripped text
    StripBetween = strText
End Function

Public Function StripNonText(ByVal strText As String) _
                             As String
    'This will strip all non-letters from the string, except spaces
    
    Dim intCounter      As Integer      'used to cycle through the character codes
    
    For intCounter = 0 To 255
        Select Case intCounter
        Case 32, 65 To 90, 97 To 122
            'valid character
            
        Case Else
            'strip character from string
            strText = Replace(strText, Chr(intCounter), " ")
        End Select
    Next intCounter
    
    'return the text
    StripNonText = RemoveDuplicates(strText)
End Function

Public Function WrapText(ByVal strText As String, _
                         ByVal lngWrapToLength As Long) _
                         As String
    'This will wrap a single line of text into seperate lines, of maximum
    'length specified by lngWrapToLength
    
    Dim lngCounter  As Long     'used to cycle through the text
    Dim strWords()  As String   'holds each seperate word in the text
    Dim strWrapped  As String   'holds the wrapped text
    Dim strLine     As String   'holds a single line of text
    
    'parse the string into seperate words
    strWords = Split(strText, " ")
    
    For lngCounter = (LBound(strWords)) To (UBound(strWords))
        Select Case (Len(strWords(lngCounter)) + Len(strLine) + 1)
        Case Is > lngWrapToLength
            'break down the word into line sized lengths
            Do
                If strWrapped <> "" Then
                    'add line
                    strWrapped = strWrapped & vbCrLf & strLine
                Else
                    'this is the first line
                    strWrapped = strLine
                End If
                If Len(strWords(lngCounter)) > lngWrapToLength Then
                    'add a section of the word
                    strLine = Left(strWords(lngCounter), lngWrapToLength)
                    strWords(lngCounter) = Mid(strWords(lngCounter), lngWrapToLength)
                Else
                    'add the whole word
                    strLine = strWords(lngCounter)
                    strWords(lngCounter) = ""
                End If
            Loop Until (strWords(lngCounter) = "")
            
        Case lngWrapToLength
            'enter the word and complete the line
            strLine = strLine & " " & strWords(lngCounter)
            strWrapped = strWrapped & vbCrLf & strLine
            strLine = ""
            
        Case Is < lngWrapToLength
            'add the word to the current line
            strLine = Trim(strLine & " " & strWords(lngCounter))
        End Select
    Next lngCounter
    
    'if there is still something in the strLine buffer, then add it to the
    'wrapped text
    If strLine <> "" Then
        strWrapped = strWrapped & vbCrLf & strLine
    End If
    
    'if the first two characters are vbCrLf then, remove them
    If Left(strWrapped, 2) = vbCrLf Then
        strWrapped = Mid(strWrapped, 3)
    End If
    
    'return the wrapped text
    WrapText = strWrapped
End Function
