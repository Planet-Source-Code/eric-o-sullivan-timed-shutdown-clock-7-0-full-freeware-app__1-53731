VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEggRun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run Program"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "frmEggRun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboOpen 
      Height          =   315
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   120
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblOpen 
      BackStyle       =   0  'Transparent
      Caption         =   "Open:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.Image imgPicture 
      Height          =   480
      Left            =   240
      Picture         =   "frmEggRun.frx":0442
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Type the name of a program, folder or document and Windows will open it for you."
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmEggRun"
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
'TITLE :    Run Dialog Screen (Egg)
' -----------------------------------------------
'COMMENTS :
'This screen is used in exactly the same way as
'the Run Dialog screen you find on your Task bar.
'It requires the module modRegistry to open the
'files correctly.
'This particular version is only intended as a
'hidden program "egg".
'=================================================

'require variable declaration
Option Explicit

'------------------------------------------------
'               MODULE-LEVEL CONSTANTS
'------------------------------------------------

'the registry sub-key and data name where we can find
'the recently run programs from the Run dialog box from
'the start bar. The registry entry is under CURRENT_USER
Private Const RUN_LIST_DATA     As String = "MRUList"
Private Const RUN_LIST_SUBKEY   As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Private Sub cboOpen_GotFocus()
    'hoghlight the text
    With cboOpen
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cmdBrowse_Click()
    'set up the Open dialog box
    
    Const FLAGS = cdlOFNExplorer Or _
                  cdlOFNNoLongNames Or _
                  cdlOFNPathMustExist Or _
                  cdlOFNFileMustExist
    Const FILTER = "Program Files (*.exe, *.bat, *.com)|*.exe;*.bat;*.com|" & _
                   "All Files (*.*)|*.*"
    
    With dlgBrowse
        'set the initial directory to the windows
        'directory
        .InitDir = GetWinDirectories(WindowsDir)
        .FLAGS = FLAGS
        .FILTER = FILTER
        .FilterIndex = 0
        
        'try and open a file
        .ShowOpen
        
        If .FileName <> "" Then
            'update the display
            cboOpen.Text = .FileName
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    'exit
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    'try and run the specified file
    
    Dim lngAppId As Long    'the process id for the application started (if any)
    Dim lngResponse As Long 'the users response to the message box
    
    'try to open the file
    lngAppId = ShellFile(cboOpen.Text)
    
    'could we open the file successfully?
    If lngAppId <> 0 Then
        'file was opened successfully - exit
        Unload Me
    Else
        'error - display a message
        lngResponse = MsgBox("Unable to open file. This may have the following causes;" & _
                             vbCrLf & _
                             vbCrLf & _
                             "1) File type does not have a registered program to open it" & _
                             vbCrLf & _
                             "2) Cannot read registered information" & _
                             vbCrLf & _
                             "3) You tried to open a shortcut" & _
                             vbCrLf & _
                             "4) File does not exist" & _
                             vbCrLf & _
                             "5) Invalid file name", _
                             vbExclamation + vbOKOnly, _
                             "Error")
    End If
End Sub

Private Sub Form_Activate()
    'highlight the text
    cboOpen.SetFocus
End Sub

Private Sub Form_Load()
    'set the text box to show the windows directory
    Call GetRecentList
    
    'set the form fonts to the current system default
    Call SetFormFontsToSystem(Me, FNT_MESSAGE)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'flag to clear this form from memory
    gblnDoCleanUp = True
End Sub

Private Sub GetRecentList()
    'This will get a list of recently run programs
    'from the registry and select the most recent
    'in the combo box
    
    Dim strFileOrder As String      'holds the order in which to display the file list and also gives me the key names in the registry
    Dim lngCounter As Long          'used to cycle through each file in the list
    Dim strBuffer As String         'temporarily holds the string that was used to run the file
    
    'get the file list order from the registry
    strFileOrder = ReadRegString(HKEY_CURRENT_USER, _
                                 RUN_LIST_SUBKEY, _
                                 RUN_LIST_DATA)
    
    'if we were unable to retrieve any data, then exit
    If LCase(Left(strFileOrder, 5)) = "error" Then
        Exit Sub
    End If
    
    'enter each file string into the combo box
    cboOpen.Clear
    For lngCounter = 1 To (Len(strFileOrder))
        'get the file string
        strBuffer = ReadRegString(HKEY_CURRENT_USER, _
                                  RUN_LIST_SUBKEY, _
                                  Mid(strFileOrder, _
                                      lngCounter, _
                                      1))
        
        'remove the trailing "\1" from the buffer text
        strBuffer = Left(strBuffer, Len(strBuffer) - 2)
        
        'enter the file string into the combo box
        Call cboOpen.AddItem(strBuffer)
    Next lngCounter
    
    'select the first item
    If cboOpen.ListCount > 0 Then
        cboOpen.ListIndex = 0
    Else
        'display the windows directory path
        cboOpen.Text = GetWinDirectories(WindowsDir)
    End If
End Sub
