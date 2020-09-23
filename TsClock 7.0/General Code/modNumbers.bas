Attribute VB_Name = "modNumbers"
'-------------------------------------------------------------------------------
'                                MODULE DETAILS
'-------------------------------------------------------------------------------
'   Program Name:   General Use
'  ---------------------------------------------------------------------------
'   Author:         Eric O'Sullivan
'  ---------------------------------------------------------------------------
'   Date:           02 August 2003
'  ---------------------------------------------------------------------------
'   Company:        CompApp Technologies
'  ---------------------------------------------------------------------------
'   Contact:        DiskJunky@hotmail.com
'  ---------------------------------------------------------------------------
'   Description:    This manages various manipulations of numbers
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

'require variable declaration
Option Explicit

'-------------------------------------------------------------------------------
'                                 API DECLARATIONS
'-------------------------------------------------------------------------------
'this will copy memory from one point in ram to another
Private Declare Sub CopyMemory _
        Lib "kernel32" _
        Alias "RtlMoveMemory" _
            (pDst As Any, _
             pSrc As Any, _
             ByVal ByteLen As Long)

'-------------------------------------------------------------------------------
'                                 PROCEDURES
'-------------------------------------------------------------------------------
Public Function BaseToDecimal(ByVal intBase As Integer, _
                              ByVal strBase As String) _
                              As Variant
    'This will return the decimal value of the Bases number. The base conversion is
    'limited from base 2 to 36.
    
    Dim strChar         As String * 1       'holds a single character
    Dim lngCounter      As Long             'used to cycle through the characters in the Base string
    Dim varTemp         As Variant          'holds the decimal value of a single digit from the Base string
    Dim varTotal        As Variant          'holds the total decimal value as it's being calculated
    Dim lngLen          As Long             'holds the length of the Base string
    Dim intSign         As Integer          'holds what the sign of the Base value is
    
    'make sure that the specified base is valid
    If (intBase < 2) Or (intBase > 36) Then
        'invalid base
        Exit Function
    End If
    
    lngLen = Len(strBase)
    If (lngLen = 0) Then
        'there is nothing to process
        Exit Function
    End If
    strBase = UCase(strBase)
    
    'check the first character
    If (Left(strBase, 1) = "-") Then
        intSign = -1    'negative number
        strBase = Mid(strBase, 2)
        lngLen = lngLen - 1
    Else
        intSign = 1     'positive number
    End If
    
    'use CDec to convert the variant to Decimal data type
    varTotal = CDec(0)
    varTemp = CDec(0)
    
    'cycle through the characters
    For lngCounter = lngLen To 1 Step -1
        'get the character to process
        strChar = Mid(strBase, lngCounter, 1)
        
        Select Case strChar
        Case "0" To "9"
            varTemp = Val(strChar)
            
        Case "A" To "Z"     'values 10 to 36
            varTemp = Asc(strChar) - 55
            
        Case Else
            'invalid string
            Exit For
        End Select
        
        'make sure that the character is not outside this bases range
        If (varTemp >= intBase) Then
            'invalid character
            Exit For
        End If
        
        'get the Base value for this digit
        varTemp = varTemp * (intBase ^ (lngLen - lngCounter))
        
        'add to total
        varTotal = varTotal + varTemp
    Next lngCounter
    
    'return the value
    BaseToDecimal = varTotal * intSign
End Function

Public Function DecimalToBase(ByVal intBase As Integer, _
                              ByVal strNumber As String) _
                              As String
    'This will convert the decimal number to the specified base. This function is limited
    'to a base range of 2 to 36. The data type should be passed into the procedure as string
    'as precision is lost on large numbers when vb converts to scientific notation. This
    'string is then converted to the Decimal data type (not intrinsic and can only be used as
    'a sub type of Variant).
    
    Dim strBase         As String           'holds the string to return to the user
    Dim decResult       As Variant          'holds the decimal number as we divide it by the base
    Dim decIntResult    As Variant          'holds the decimal number rounded down
    Dim intDigit        As Integer          'holds a single Base digit in decimal form
    Dim strDigit        As String           'holds the string form of the Base digit
    Dim intSign         As Integer          'holds what the sign of the decimal value is
    Dim decNumber       As Variant          'holds the Decimal form of the number (14 byte variable)
    
    'trim any decimals
    decNumber = CDec(strNumber)
    decNumber = Int(decNumber)
    intSign = Sgn(decNumber)
    decNumber = Abs(decNumber)      'we will account for negatives later
    
    
    'account for zero
    If (decNumber = 0) Then
        DecimalToBase = "0"
        Exit Function
    End If
    
    'initialise the Base variables
    decResult = decNumber
    strBase = ""
    
    'keep dividing until we reach zero
    Do While decResult > 0
        'get the next digit
        decIntResult = CDec(Int(decResult / intBase))
        intDigit = decResult - (decIntResult * intBase) ' Mod intBase
        
        'convert the digit to it's string form
        Select Case intDigit
        Case 0 To 9
            'normal numeric digit - no conversion necessary
            strDigit = intDigit
            
        Case 10 To 35
            'convert to a letter
            strDigit = Chr(intDigit + 55)
            
        Case Else
            'invalid digit (this shouldn't occur)
            Exit Do
        End Select
        
        'add the digit to the final Base value
        strBase = strDigit + strBase
        
        'reduce the number
        decResult = decIntResult
    Loop
    
    'do we account for negatives
    If (intSign < 0) Then
        strBase = "-" + strBase
    End If
    
    'return the result
    DecimalToBase = strBase
End Function

Public Function GetRndNum(ByVal dblMin As Double, _
                          ByVal dblMax As Double) _
                          As Double
    'This function will produce a random number between the specified
    'values
    
    Dim dblTemp As Double       'used for swapping two values
    
    'if the two values are equal then
    If dblMin = dblMax Then
        'return same number and exit
        GetRndNum = dblMin
        Exit Function
    End If
    
    'if dblMin is bigger than dblMax then swap values
    If dblMin > dblMax Then
        dblTemp = dblMin
        dblMin = dblMax
        dblMax = dblMin
    End If
    
    'Randomize and make sure that the resulting number is exactly between
    'the ranges specified
    GetRndNum = CheckRange(dblMin, dblMax, ((dblMax - dblMin + 1) * Rnd + dblMin))
End Function

Public Function CheckRange(ByVal dblMin As Double, _
                           ByVal dblMax As Double, _
                           ByVal dblValue As Double) _
                           As Double
    'This function will check to see if dblValue is between dblMin and dblMax.
    'If not then it will divide dblValue until it fits between dblMin and
    'dblMax.
    
    Dim dblOffset   As Double       'holds the difference between min and max
    
    'exit program is max is less than or equal to min
    If (dblMax <= dblMin) Then
        CheckRange = dblMin
        Exit Function
    End If
    
    'make sure dblValue does not exceed bounds under any initial conditions
    dblOffset = dblMax - dblMin
    dblValue = (dblValue Mod (dblOffset + 1)) + dblMin
    If (dblValue < dblMin) Then
        dblValue = dblMax - (dblMin - dblValue)
    End If
    
    'return the result
    CheckRange = dblValue
End Function

Public Function LimitRange(ByVal dblMin As Double, _
                           ByVal dblMax As Double, _
                           ByVal dblValue As Double) _
                           As Integer
    'This function takes three value, dblMin, dblMax and dblValue. If
    'dblValue is below dblMin, then value = dblMin. If dblValue is greater
    'than dblMax, then dblValue is equal to dblMax.
    
    If (dblValue < dblMin) Then
        dblValue = dblMin
    ElseIf (dblValue > dblMax) Then
        dblValue = dblMax
    End If
    
    'return result
    LimitRange = dblValue
End Function

Public Sub FlipDbl(ByRef dblVal1 As Double, _
                   ByRef dblVal2 As Double)
    'This procedure will swap the two values
    
    Dim dblTemp     As Double       'holds the first value during the swap
    
    'swap the values
    dblTemp = dblVal1
    dblVal1 = dblVal2
    dblVal2 = dblTemp
End Sub

Public Function BinToLong(ByVal strBinary As String) _
                          As Long
    'This converts a binary string to decimal. It will stop at the first
    'non-binary character in the string
    
    Dim lngCounter  As Long         'used to cycle through the string
    Dim lngResult   As Long         'holds the value of the binary string in decimal
    Dim lngBinLen   As Long         'holds the length of the string
    Dim strChar     As String * 1   'holds a single character from the string
    
    'get the length of the string passed
    lngBinLen = Len(strBinary)
    If (lngBinLen = 0) Then
        'the string is empty - we cannot do anything with it
        BinToLong = 0
        Exit Function
    End If
    
    For lngCounter = 1 To Len(strBinary)
        'get the character from the string and make sure that it is numeric
        strChar = Mid(strBinary, lngCounter, 1)
        If Not IsNumeric(strChar) Then
            'we can do nothing else with the string - return any exitsing
            'result
            Exit For
        End If
        
        '(2 ^ x) * digit
        lngResult = lngResult + ((2 ^ (lngBinLen - lngCounter)) * Val(strChar))
    Next lngCounter
    
    'return the result
    BinToLong = lngResult
End Function

Public Function HexToDouble(ByVal strHex As String) As Double
    'This will return the decimal value of the Hex number
    
    Dim strChar         As String * 1       'holds a single character
    Dim lngCounter      As Long             'used to cycle through the characters in the Hex string
    Dim dblTemp         As Double           'holds the decimal value of a single digit from the Hex string
    Dim dblTotal        As Double           'holds the total decimal value as it's being calculated
    Dim lngLen          As Long             'holds the length of the hex string
    Dim intSign         As Integer          'holds what the sign of the Hex value is
    
    lngLen = Len(strHex)
    If (lngLen = 0) Then
        'there is nothing to process
        Exit Function
    End If
    strHex = UCase(strHex)
    
    'check the first character
    If (Left(strHex, 1) = "-") Then
        intSign = -1    'negative number
        strHex = Mid(strHex, 2)
        lngLen = lngLen - 1
    Else
        intSign = 1     'positive number
    End If
    
    dblTotal = 0
    dblTemp = 0
    
    'cycle through the characters
    For lngCounter = lngLen To 1 Step -1
        'get the character to process
        strChar = Mid(strHex, lngCounter, 1)
        
        Select Case strChar
        Case "0" To "9"
            dblTemp = Val(strChar)
            
        Case "A" To "F"     'values 10 to 15
            dblTemp = Asc(strChar) - 55
            
        Case Else
            'invalid string
            Exit For
        End Select
        
        'get the hex value for this digit
        dblTemp = dblTemp * (16 ^ (lngLen - lngCounter))
        
        'add to total
        dblTotal = dblTotal + dblTemp
    Next lngCounter
    
    'return the value
    HexToDouble = dblTotal * intSign
End Function

Public Function BinToDouble(ByVal strBin As String) As Double
    'This will return the decimal value of the Bin number
    
    Dim strChar         As String * 1       'holds a single character
    Dim lngCounter      As Long             'used to cycle through the characters in the Bin string
    Dim dblTemp         As Double           'holds the decimal value of a single digit from the Bin string
    Dim dblTotal        As Double           'holds the total decimal value as it's being calculated
    Dim lngLen          As Long             'holds the length of the Bin string
    Dim intSign         As Integer          'holds what the sign of the Bin value is
    
    lngLen = Len(strBin)
    If (lngLen = 0) Then
        'there is nothing to process
        Exit Function
    End If
    strBin = UCase(strBin)
    
    'check the first character
    If (Left(strBin, 1) = "-") Then
        intSign = -1    'negative number
        strBin = Mid(strBin, 2)
        lngLen = lngLen - 1
    Else
        intSign = 1     'positive number
    End If
    
    dblTotal = 0
    dblTemp = 0
    
    'cycle through the characters
    For lngCounter = lngLen To 1 Step -1
        'get the character to process
        strChar = Mid(strBin, lngCounter, 1)
        
        Select Case strChar
        Case "0" To "1"
            dblTemp = Val(strChar)
            
        Case Else
            'invalid string
            Exit For
        End Select
        
        'get the Bin value for this digit
        dblTemp = dblTemp * (2 ^ (lngLen - lngCounter))
        
        'add to total
        dblTotal = dblTotal + dblTemp
    Next lngCounter
    
    'return the value
    BinToDouble = dblTotal * intSign
End Function

Public Function LongToBin(ByVal lngNumber As Long) _
                          As String
    'This will convert a decimal long to binary.
    
    Dim lngResult   As Long         'holds the decreasing value of the original number as it is converted
    Dim strBinary   As String       'holds the number in binary
    
    'convert the long to a binary string
    lngResult = lngNumber
    Do While lngResult > 0
        strBinary = strBinary & (lngResult Mod 2)
        lngResult = lngResult \ 2
    Loop
    
    'return the result
    LongToBin = StrReverse(strBinary)
End Function

Public Function SecToTime(Seconds As Long) _
                          As String
    'This function will convert the time in seconds to a time format
    '(hh:mm:ss)
    
    SecToTime = Format((Seconds \ 3600), "00") & ":" & _
                Format((Seconds \ 60) Mod 60, "00") & ":" & _
                Format(Seconds Mod 60, "00")
End Function

Public Function TimeToSec(ByVal strTime As String) _
                          As Long
    'This function will convert the time into the number of seconds in
    'that time.
    
    If (strTime <> "") Then
        TimeToSec = (Hour(strTime) * 3600) + _
                    (Minute(strTime) * 60) + _
                    (Second(strTime))
    End If
End Function

Public Function RetProbability(ByVal sngChance As Single) _
                               As Boolean
    'This function takes a parameter containing the percent change of
    'True happening
    
    Dim sngRndResult    As Single
    
    'make sure that the parameter passed is valid
    sngChance = CheckRange(0, 100, sngChance)
    
    sngRndResult = GetRndNum(0, 100)
    
    'if the result returned is less than the percentage chance of True
    'happening, then return True, else return false
    If (sngRndResult <= sngChance) Then
        RetProbability = True
    Else
        RetProbability = False
    End If
End Function

Public Function QSString(ByVal strText As String) _
                         As String
    'This function uses the Quick Sort method to sort a string.
    
    'after the first pass of the string, the character with the lowest
    'ASCII number has been moved to the start of the string.
    'Since the lowest is found, then search the rest of the string
    'for the next lowest.
    
    Dim lngCounter  As Long     'used to cycle through the string
    Dim lngPivot    As Long     'holds a reference to a particular character in the string
    Dim strTemp     As String   'holds a single character in the string
    
    'don't run function if passed empty string
    If strText = "" Then
        Exit Function
    End If
    
    For lngPivot = 1 To Len(strText)
        For lngCounter = lngPivot To Len(strText)
            'If GetAsc(strText, lngCounter) < GetAsc(strText, lngPivot) Then
            If (Mid(strText, lngCounter, 1) < Mid(strText, lngPivot, 1)) Then
                'switch values and set counter to new position
                
                strTemp = Left(strText, lngPivot - 1) 'get everything before the pivot position
                strTemp = strTemp & Mid(strText, lngCounter, 1) 'store the value found that was lower to new position
                strTemp = strTemp & Mid(strText, lngPivot + 1, (lngCounter - lngPivot) - 1) 'get everything between lngCounter and the pivot point
                strTemp = strTemp & Mid(strText, lngPivot, 1) 'store the pivot to its new forward position
                strTemp = strTemp & Mid(strText, lngCounter + 1, (Len(strText) - lngCounter)) 'store everything after the pivot point
                
                'The new counter value is where the pivot is
                lngCounter = lngPivot
                
                'keep the new string
                strText = strTemp
            End If
        Next lngCounter
    Next lngPivot
    
    QSString = strText
End Function

Public Function GetTimeLeft(ByVal lngTicksElapsed As Long, _
                            ByVal lngTotal As Long, _
                            ByVal sngCurrentPoint As Single, _
                            Optional ByVal lngMin As Long = 0) _
                            As String
    'This will return the expected time left to finish (to lngTotal) given the current
    'time elapsed in Ticks (Second/1000), and the current point (a number from lngMin to
    'lngTotal). The time returned is formatted hh:mm:ss.
    
    Dim strTimeLeft     As String           'holds the actual time left to finish
    Dim intSeconds      As Integer          'holds the seconds left until finish
    Dim intMinutes      As Integer          'holds the minutes left until finish
    Dim lngHours        As Long             'holds the hours left until finish
    Dim lngTicksLeft    As Long             'holds the number of tick left until the finish
    Dim lngDiff         As Long             'holds the actual "distance" between lngMin and lngTotal
    
    'validate the parameters
    If (lngTotal <= lngMin) Then
        'there is no time to calculate
        GetTimeLeft = "00:00:00"
        Exit Function
    End If
    Select Case sngCurrentPoint
    Case Is < lngMin
        sngCurrentPoint = lngMin
    
    Case Is > lngTotal
        sngCurrentPoint = lngTotal
    End Select
    
    'get the actual "distance" that the "CurrentPoint" has to "travel" until finish
    lngDiff = lngTotal - lngMin
    
    'get the time left in ticks
    lngTicksLeft = (((lngTicksElapsed / (sngCurrentPoint + lngMin)) * lngDiff) - lngTicksElapsed)
    
    'convert the ticks to time
    strTimeLeft = SecToTime((lngTicksLeft / 1000))
    
    'return the result
    GetTimeLeft = strTimeLeft
End Function

Public Sub OnlyNumeric(ByRef intKeyAscii As Integer, _
                       ByRef txtBox As TextBox, _
                       Optional ByVal intMaxDecimals As Integer = -1)
    'This will ensure that the new character is numeric and is valid for the specified number of decimals.
    'If no decimals were specified (-1), then any number of decimals are allowed. Please note that if
    'the user presses [RETURN] and character 13 is passed to this function, intKeyAscii will then be 0. The
    '[RETURN] must be processed before this procedure is called. This should be called from the KeyPress event
    'of the text box in question.
    
    Dim strCurrText     As String       'holds the current text that the user has entered
    Dim lngCharPos      As Long         'holds the position of a character in the text
    
    If (txtBox Is Nothing) Then
        'there is no text box to work with
        Exit Sub
    End If
    
    'get the current text in the box
    strCurrText = Trim(txtBox.Text)
    
    'how many decimals is the user allowed enter
    If (intMaxDecimals < 0) Then
        intMaxDecimals = -1
    End If
    
    'is there any decimal place in the existing text
    lngCharPos = InStr(1, strCurrText, ".")
    
    Select Case intKeyAscii
    Case 48 To 57   'numeric value - further processing will be done after
    
    Case Asc(".")
        'only allow one decimal character (if decimals are allowed)
        If (lngCharPos > 0) Or (intMaxDecimals = 0) Then
            'there is already a decimal
            intKeyAscii = 0
        End If
        Exit Sub
    
    Case 8  'backspace - allow, but don't process anything
        Exit Sub
    
    Case Else
        'invalid character
        intKeyAscii = 0
    End Select
    
    'was the character allowed
    If (intKeyAscii = 0) Then
        Exit Sub
    End If
    
    'check the number of decimals
    If ((Len(strCurrText) - lngCharPos) >= intMaxDecimals) And (intMaxDecimals > 0) Then
        'the user has already entered the maximum number of decimals but check if they are entering this
        'new number before the decimal or after
        
        If ((txtBox.SelStart + 1) > lngCharPos) Then
            'user is entering the character after the decimal
            intKeyAscii = 0
        End If
    End If
End Sub
