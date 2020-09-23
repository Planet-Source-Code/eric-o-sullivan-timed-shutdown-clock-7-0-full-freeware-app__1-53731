Attribute VB_Name = "modEncryption"
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     20 Feburary 2002
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    Encryption Module
' -----------------------------------------------
'COMMENTS :
'This encryption algorithim first gets the tick
'of the current second (including the character
'"0" if applicable in two digits), converts both
'characters into ascii values, adds them to a
'prime number and uses that as the encryption
'key. The value is encrypted under a simple
'ascii value addition and can be deduced from
'the first few characters of the encrypted
'string. The first character of the encrypted
'data is ALWAYS the ascii value of how many
'characters after it is the decrypt key, ie the
'length of the decrypt key is the ascii value of
'the first character.
'=================================================

'require variable declaration
Option Explicit

'this module cannot be accessed outside this project
Option Private Module

'------------------------------------------------
'               API DECLARATIONS
'------------------------------------------------
'used to help get a "unique" key value
Private Declare Function GetTickCount _
        Lib "kernel32" () _
                        As Long

'------------------------------------------------
'               MODULE-LEVEL CONSTANTS
'------------------------------------------------
'These numbers are to be unique to this program only.
'Data can only be decrypted using these numbers
Private Const BASE_KEY      As Integer = 71 'used to encrypt the main key
Private Const ADD_TO_KEY    As Integer = 19 'added to help form the main key

'some numbers to help get only alphanumeric characters
Private Const NUM_OFFSET    As Integer = 48 'the difference between the ascii value and the character index
Private Const CAP_OFFSET    As Integer = 55 'the difference between the ascii value and the character index
Private Const LOW_OFFSET    As Integer = 61 'the difference between the ascii value and the character index

'the amount of characters encrypted during file operations.
Private Const FILE_IN_TAKE  As Integer = 30720

'a list of the base 64 characters
Private Const BASE_64_CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Private Function GenerateKey() As Integer
    'generates the main key use for encryption.
    
    Dim intMilliSecond As Integer
    
    'I changed the daynum value to hold a second value
    'instead of a day value for more variances.
    'Changed again to an even shorter time value.
    intMilliSecond = (GetTickCount Mod 100)
    GenerateKey = Val(Trim(Str(Format(intMilliSecond, "00")))) + ADD_TO_KEY
End Function

Public Function EncryptData(ByVal strText As String) _
                            As String
    'This will encrypt a string and return the encryped
    'information. See also DecryptData()
    
    Dim lngCounter  As Long          'used to go through each character in the string
    Dim intDayKey   As Integer       'the key used to encrypt the main string
    Dim strRetData  As String        'the encrypted text
    Dim strEncrypt  As String * 1    'an encrypted character
    Dim lngTextLen  As Long          'holds the length of the string (used to improve performance over "Len(strText)" )
    
    'if strText is empty, return empty
    If strText = "" Then
        EncryptData = ""
        Exit Function
    End If
    
    intDayKey = GenerateKey
    
    'store the amount of digits intDayKey is, in the first
    'character.
    strRetData = Chr(Len(Trim(Str(intDayKey))))
    strRetData = strRetData & EncryptKey(Trim(Str(intDayKey)))
    
    'strEncrypt the rest of the data
    
    lngTextLen = Len(strText)
    lngCounter = 1
    For lngCounter = 1 To lngTextLen
        'encrypt each character by adding the key value to the ascii value
        'of each character
        strEncrypt = Chr((Asc(Mid(strText, _
                                  lngCounter, _
                                  1)) _
                          + intDayKey) _
                         Mod 256)
        
        'save the encrypted character
        strRetData = strRetData & strEncrypt
    Next lngCounter
    
    'return the encrypted text
    EncryptData = strRetData
End Function

Public Function DecryptData(ByVal strText As String) _
                            As String
    'This will strDecrypt information encrypted using the
    'EncryptData function.
    
    Dim intCounter      As Integer
    Dim strDayNum       As String
    Dim intDayKey       As Integer
    Dim strRetData      As String
    Dim strDecrypt      As String
    Dim intDecryptNum   As Integer
    Dim lngTextLen      As Long
    
    'get the amount of digits the key is and strDecrypt the
    'key
    If (strText = "") Or (strText = vbNullString) Then
        Exit Function
    End If
    
    'get the amount of digits the key is
    strDayNum = GetKeyLength(strText)
    
    'strDecrypt the key from the encrypted strText
    intDayKey = Val(DecryptKey(Mid(strText, 2, Val(strDayNum))))
    
    'strDecrypt the rest of the strText
    lngTextLen = Len(strText)
    intCounter = (Val(strDayNum) + 2)
    
    While intCounter <= lngTextLen
        DoEvents
        
        'subtract the key value from the ascii value of each character
        'and account for a negative lngResult
        intDecryptNum = (Asc(Mid(strText, intCounter, 1)) - intDayKey) Mod 256
        If intDecryptNum < 0 Then
            intDecryptNum = 255 + intDecryptNum
        Else
            intDecryptNum = intDecryptNum Mod 256
        End If
        
        'the character has been strDecrypted, save the lngResult
        strDecrypt = Right(Chr(intDecryptNum), 1)
        strRetData = strRetData & strDecrypt
    
        'next character
        intCounter = intCounter + 1
    Wend
    
    'return the strDecrypted data
    DecryptData = strRetData
End Function

Public Function GetKeyLength(ByVal strText As String) _
                             As String
    'get the length of digits the key is and returns the result
    
    Dim intKeyLength As Integer
    
    'get the amount of digits the key is and decrypt the
    'key
    If strText = "" Then
        Exit Function
    End If
        
    intKeyLength = Asc(Left(strText, 1))
    
    GetKeyLength = intKeyLength
End Function

Private Function EncryptKey(ByVal strKey As String) _
                            As String
    'adds the encryption strKey to the ASCII value of each
    'character.
    
    Dim intCounter  As Integer
    Dim strNewKey   As String
    
    On Error Resume Next
    
    For intCounter = 1 To Len(strKey)
        strNewKey = strNewKey & Right(Chr(Asc(Mid(strKey, intCounter, 1)) + BASE_KEY), 1)
    Next intCounter
    
    EncryptKey = strNewKey
End Function

Private Function DecryptKey(ByVal strKey As String) _
                            As String
    'subtracts the encryption strKey from the ASCII value
    'of each character.
    
    Dim intCounter  As Integer
    Dim strNewKey   As String
    
    'decrypt the key
    For intCounter = 1 To Len(strKey)
        strNewKey = strNewKey & Right(Chr(Asc(Mid(strKey, intCounter, 1)) - BASE_KEY), 1)
    Next intCounter
    
    If strKey = "" Then strNewKey = ""
    
    DecryptKey = strNewKey
End Function

Public Sub FileEncrypt(ByVal strSourcePath As String, _
                       ByVal strDestPath As String)
    'This procedure takes two arguments. The file you want encrypted
    'and the name and destination of the file you want the encrypted
    'data in.
    
    Dim strBuffer       As String    '30K strBuffer
    Dim intErrnum       As Integer
    Dim intFileNumOut   As Integer
    Dim intFileNumIn    As Integer
    
    'check for errors accessing the files
    '-------------------
    
    'if the source and destination files are the same, the exit
    If LCase(strSourcePath) = LCase(strDestPath) Then
        Exit Sub
    End If
    
    'check for access errors reading the source file
    On Error Resume Next
    intFileNumIn = FreeFile

    Open strSourcePath For Input As #intFileNumIn
        intErrnum = Err
    Close #intFileNumIn
    
    'if an error occured, the exit
    If intErrnum <> 0 Then
        Exit Sub
    End If
    
    'check for access error for writing the file
    Open strDestPath For Output As #intFileNumIn
        intErrnum = Err
    Close #intFileNumIn
    
    'if and error occurred, the exit
    If intErrnum <> 0 Then
        Exit Sub
    End If
    On Error GoTo 0
    '-------------------
    
    'encrypt the file
    intFileNumIn = FreeFile
    Open strSourcePath For Binary As #intFileNumIn
        intFileNumOut = FreeFile
        
        Open strDestPath For Binary As #intFileNumOut
            While Not EOF(intFileNumIn)
                'input a 30K chunk of the file to be encrypted
                strBuffer = Input((FILE_IN_TAKE), #intFileNumIn)
                
                'encrypt the information
                strBuffer = EncryptData(strBuffer)
                
                'save the encrypted information
                Put #intFileNumOut, , strBuffer
            Wend
        Close #intFileNumOut
    Close #intFileNumIn
End Sub

Public Sub FileDecrypt(ByVal strSoucePath As String, _
                       ByVal strDestPath As String)
    'This procedure takes two arguments. The file you want decrypted
    'and the name and destination of the file you want the decrypted
    'data in.
    
    Const FILE_IN_TAKE = 30720 'file buffer
    
    Dim strKeyLenChar       As String
    Dim intKeyLen           As Integer
    Dim strEncryptedData    As String
    Dim strDecryptedData    As String
    Dim intErrorNum         As Integer
    Dim intFileNumOut       As Integer
    Dim intFileNumIn        As Integer
    
    'check for errors accessing the files
    '-------------------
    
    'if the source and destination are the same then exit
    If LCase(strSoucePath) = LCase(strDestPath) Then
        Exit Sub
    End If
    
    'check for access errors for reading the source file
    On Error Resume Next
    intFileNumIn = FreeFile
    
    Open strSoucePath For Input As #intFileNumIn
        intErrorNum = Err
    Close #intFileNumIn
    
    'if an error occurred then exit
    If intErrorNum <> 0 Then
        Exit Sub
    End If
    
    'check for access errors when writing the file
    Open strDestPath For Output As #intFileNumIn
        intErrorNum = Err
    Close #intFileNumIn
    
    'if an error occurred the exit
    If intErrorNum <> 0 Then
        Exit Sub
    End If
    On Error GoTo 0
    '-------------------
    
    'This decryption works by the following steps;
    '1) Input the first character of the file. This will contain the length
    '    of the decyption key.
    '2) Input the decryption Key using the key length
    '3) Input the next 30720 characters - this is the encrypted data
    '4) Repeat steps 1, 2 and three until the entire file is read.
    
    intFileNumIn = FreeFile
    Open strSoucePath For Binary As #intFileNumIn
        
        intFileNumOut = FreeFile
        Open strDestPath For Binary As #intFileNumOut
            While Not EOF(intFileNumIn)
                'input the character with the keylength
                strKeyLenChar = Input(1, #intFileNumIn)
                
                'save the keylength
                intKeyLen = GetKeyLength(strKeyLenChar)
                
                'get the decryption key and the encrypted data
                strEncryptedData = Input(intKeyLen + FILE_IN_TAKE, #intFileNumIn)
                
                'decrypt the info recovered
                strDecryptedData = DecryptData(strKeyLenChar & strEncryptedData)
                
                'save decrypted data
                Put #intFileNumOut, , strDecryptedData
            Wend
        Close #intFileNumOut
    Close #intFileNumIn
End Sub

Public Function EncryptText(ByVal strText As String) _
                            As String
    'This will encrypt any valid text (only containing
    'valid alphanumeric characters only), and return
    'the encrypted text. The ascii number of the first
    'letter in the text returned is the key used to
    'encrypt the rest of the text. This function is
    'ideal for passwords.
    
    'if no text was passed, then return a blank string
    If Trim(strText) = "" Then
        EncryptText = ""
        Exit Function
    End If
    
    'encypt data as normal
    strText = EncryptData(strText)
    
    'encode in base 64 format so that the password
    'is not compatable with any text-based storage
    'method
    strText = Base64Encode(strText)
    
    'return the encrypted text
    EncryptText = strText
End Function

Public Function DecryptText(ByVal strText As String) _
                            As String
    'This function will decrypt any text encrypted with
    'the EncryptText and Base64Decode function
    
    'the string passed must be base 64 encrypted so that it can
    'be stored in text format, so we must decode this first
    strText = Base64Decode(strText)
    
    'decrypt as normal
    strText = DecryptData(strText)
    
    'return the decrypted text
    DecryptText = strText
End Function

Private Function WrapAlphaNumeric(ByVal intAscii As Integer, _
                                  ByVal intAdd As Integer) _
                                  As Integer
    'This procedure will add a numeric value to the
    'ascii value of a VALID alphanumeric character
    'making sure that the result is also an alphanumeric
    'character, wrapping the value if necessary
    '
    'The valid ascii ranges are; 48-57, 65-90, 97-122
    '                             0-9 ,  A-Z ,  a-z
    
    Const MAX_VALID As Integer = 62    'there are 62 valid characters
    
    'make sure that the values passed are valid
    intAscii = AsciiToIndex(intAscii)
    
    'wrap the value to be added
    intAdd = intAdd Mod MAX_VALID
    
    'add the value to create to create a new character
    intAscii = intAscii + intAdd
    
    'wrap the value accounting for negatives
    If intAscii < 0 Then
        'value is negative, wrap from top ("z")
        intAscii = MAX_VALID + intAscii
    Else
        'value is postive - just wrap
        intAscii = intAscii Mod MAX_VALID
    End If
    
    'convert the number back into a valid ascii number
    intAscii = IndexToAscii(intAscii)
    
    'return the new character
    WrapAlphaNumeric = intAscii
End Function

Private Function AsciiToIndex(ByVal intAscii As Integer) _
                              As Integer
    'convert a valid ascii number to the alphanumeric
    'index
    
    Select Case intAscii
    Case 48 To 57   'numeric character
        'these characters will be numbered index 0-9
        intAscii = intAscii - NUM_OFFSET
    
    Case 65 To 90   'capital letters
        'these character will be numbered index 10-35
        intAscii = intAscii - CAP_OFFSET
    
    Case 97 To 122  'lower case letters
        'these characters will be numbered index 36-62
        intAscii = intAscii - LOW_OFFSET
    
    Case Else
        'invalid character value
        intAscii = 0
        Exit Function
    End Select
    
    'return the index value
    AsciiToIndex = intAscii
End Function

Private Function IndexToAscii(ByVal intAscii As Integer) _
                              As Integer
    'convert a valid index to it's alphanumeric ascii
    'value
    
    'make sure the index is valid
    intAscii = intAscii Mod 63
    
    Select Case intAscii
    Case 0 To 9     'numeric characters
        intAscii = intAscii + NUM_OFFSET
    
    Case 10 To 35   'capital letters
        intAscii = intAscii + CAP_OFFSET
        
    Case 36 To 62   'lower case letters
        intAscii = intAscii + LOW_OFFSET
    End Select
    
    'return the new ascii character
    IndexToAscii = intAscii
End Function

Public Function Base64EncodeFile(ByVal FilePath As String) _
                             As String
    'This will encode the specified file into Base64 format for emailing. There
    'is a one megabytes file size limit on the attachment. Should the file exceed
    'this limit, only the first megabyte is encoded and returned
    
    Const SIZE_LIMIT    As Long = 1048576   'the maximum file size that this can handle (1 Mb)
    
    Dim FileNum         As Integer  'holds a handle to the file
    Dim FileLeft        As Long     'the number of bytes left to read from the file
    Dim InBuffer        As String   'holds a section of the file
    
    'make sure the file exists
    If (FilePath = "") Or (Dir(FilePath) = "") Then
        Exit Function
    End If
    
    FileNum = FreeFile
    Open FilePath For Binary As #FileNum
        'get the number of bytes left to read from the file
        FileLeft = LOF(FileNum) - Loc(FileNum)
        
        'make sure we don't read more than one megabyte
        If (FileLeft > SIZE_LIMIT) Then
            'read one megabyte
            InBuffer = Input(SIZE_LIMIT, FileNum)
        Else
            'read entire file
            InBuffer = Input(FileLeft, FileNum)
        End If
    Close #FileNum
    
    'encode the data and return
    Base64EncodeFile = Base64Encode(InBuffer)
End Function

Public Function Base64Decode(ByVal Data As String) _
                             As String
    'make sure that we can retrieve the string in 4 character chunks. If
    '4 does not divide evenly by four, then add the necessary number of
    'spaces
    
    Dim Byte1       As Integer      'holds the first byte in the 4 character set
    Dim Byte2       As Integer      'holds the second byte in the 4 character set
    Dim Byte3       As Integer      'holds the third byte in the 4 character set
    Dim Byte4       As Integer      'holds the last byte in the 4 character set
    Dim TempData    As String       'holds four bytes for processing
    Dim Counter     As Long         'used to cycle through the string
    Dim AllDecoded  As String       'the total decode text
    Dim Decoded     As String       'holds 4 bytes of decoded text
    Dim Track       As Integer      'used to track how many chunks have been processed
    
    If Len(Data) Mod 4 > 0 Then
        'padd with the necessary spaces
        Data = Data + Space(4 - (Len(Data) Mod 4))
    End If

    Track = 0
    AllDecoded = ""
    For Counter = 1 To Len(Data) Step 4
        Decoded = "   "
        
        'get four bytes from the string
        TempData = Mid$(Data, Counter, 4)
        
        'get the string positions of each of those characters
        Byte1 = InStr(BASE_64_CHARS, Mid(TempData, 1, 1)) - 1
        Byte2 = InStr(BASE_64_CHARS, Mid(TempData, 2, 1)) - 1
        Byte3 = InStr(BASE_64_CHARS, Mid(TempData, 3, 1)) - 1
        Byte4 = InStr(BASE_64_CHARS, Mid(TempData, 4, 1)) - 1
        
        'convert to normal text
        Mid(Decoded, 1, 1) = Chr(((Byte2 And 48) \ 16) Or (Byte1 * 4) And &HFF)
        Mid(Decoded, 2, 1) = Chr(((Byte3 And 60) \ 4) Or (Byte2 * 16) And &HFF)
        Mid(Decoded, 3, 1) = Chr((((Byte3 And 3) * 64) And &HFF) Or (Byte4 And 63))
        
        'add to already decoded text
        AllDecoded = AllDecoded + Decoded
        
        'do we need to remove the carrage return and line feed (every
        '60 characters)
        Track = Track + 1
        If Track >= 19 Then
            Track = 0
            Data = Mid(Data, 3)
        End If
    Next Counter
    
    'return the decoded text
    Base64Decode = RTrim(AllDecoded)
End Function

Public Function Base64Encode(ByVal Data As String) _
                             As String
    'Base64Encode a string into base64 format
    
    Dim Track       As Integer  'used to track how much of the string has been processed. Every 60 characters, a carrage return and line feed are added (vbCrlf)
    Dim Counter     As Long     'used to scan through the Data string
    Dim AllEncoded  As String   'the complete encoded text
    Dim Encoded     As String   'an encoded section of the string. This is added into AllEncoded
    Dim Char1       As Integer  'holds the encoded first character
    Dim Char2       As Integer  'holds the encoded second character
    Dim Char3       As Integer  'holds the encoded second character
    Dim SaveBits1   As Integer  'holds any saved bits from the first character
    Dim SaveBits2   As Integer  'holds any saved bits from the second character
    Dim TempChars   As String   'holds a string section to work on
    
    If Len(Data) Mod 3 > 0 Then
        Data = Data + Space(3 - (Len(Data) Mod 3))
    End If
    
    Track = 0
    AllEncoded = ""
    For Counter = 1 To Len(Data) Step 3
        Encoded = "    "
        TempChars = Mid(Data, Counter, 3)
        
        'sort the bits of each character
        Char1 = Asc(Mid(TempChars, 1, 1))
        SaveBits1 = Char1 And 3
        
        Char2 = Asc(Mid(TempChars, 2, 1))
        SaveBits2 = Char2 And 15
        
        Char3 = Asc(Mid(TempChars, 3, 1))
        
        'encode each character
        Mid(Encoded, 1) = Mid(BASE_64_CHARS, ((Char1 And 252) \ 4) + 1, 1)
        Mid(Encoded, 2) = Mid(BASE_64_CHARS, (((Char2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
        Mid(Encoded, 3) = Mid(BASE_64_CHARS, (((Char3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
        Mid(Encoded, 4) = Mid(BASE_64_CHARS, (Char3 And 63) + 1, 1)
        
        'store the encode text
        AllEncoded = AllEncoded + Encoded
        
        'add a carrage return and line feed every 60 characters
        Track = Track + 1
        If Track >= 19 Then
            Track = 0
            AllEncoded = AllEncoded & vbCrLf
        End If
    Next Counter
    
    'return the encoded text
    Base64Encode = AllEncoded
End Function

Public Sub Base64DecodeFile(ByVal SourceFile As String, _
                            ByVal DestFile As String)
    'This will assume all data in the source file is in base 64 format and
    'will output the result to the destination file
    
    Dim InputNum    As Integer      'holds the file handle for the source file
    Dim OutputNum   As Integer      'holds the file handle for the destination file
    Dim Buffer      As String       'holds a section of the file to decode
    
    If (SourceFile = "") Or (Dir(SourceFile) = "") Then
        'invalid source file
        Exit Sub
    End If
    
    'decode the file
    InputNum = FreeFile
    Open SourceFile For Input As #InputNum
        OutputNum = FreeFile
        Open DestFile For Binary As #OutputNum
            Do While Not EOF(InputNum)
                'get the encode data from the source file
                Line Input #InputNum, Buffer
                
                'decode the line
                Buffer = Base64Decode(Buffer)
                
                'write to the output file
                Put #OutputNum, , Buffer
            Loop
        Close #OutputNum
    Close #InputNum
End Sub
