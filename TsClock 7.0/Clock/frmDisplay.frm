VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDisplay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Settings"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmDisplay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboBorder 
      Height          =   315
      ItemData        =   "frmDisplay.frx":0442
      Left            =   3960
      List            =   "frmDisplay.frx":0464
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Text            =   "C:\Windows"
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox cboStyle 
      Height          =   315
      ItemData        =   "frmDisplay.frx":0486
      Left            =   3960
      List            =   "frmDisplay.frx":049C
      TabIndex        =   2
      Text            =   "None"
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdDelScheme 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdSaveScheme 
      Caption         =   "&Save"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox cboScheme 
      Height          =   315
      ItemData        =   "frmDisplay.frx":04D0
      Left            =   3960
      List            =   "frmDisplay.frx":04D2
      TabIndex        =   6
      Text            =   "[Default]"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdSchemeSet 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdColourSet 
      Caption         =   "&Colour..."
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox cboColour 
      Height          =   315
      ItemData        =   "frmDisplay.frx":04D4
      Left            =   3960
      List            =   "frmDisplay.frx":04FC
      TabIndex        =   4
      Text            =   "Analog Background"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Frame fraPreview 
      Caption         =   "Preview"
      Height          =   3615
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2385
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   2665
         Left            =   285
         ScaleHeight     =   178
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   675
         Width           =   1815
         Begin VB.Timer timPreview 
            Interval        =   200
            Left            =   0
            Top             =   0
         End
      End
      Begin VB.PictureBox picTitleBar 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   270
         Left            =   285
         Picture         =   "frmDisplay.frx":059E
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   390
         Width           =   1815
      End
      Begin VB.PictureBox picFrame 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3040
         Left            =   240
         ScaleHeight     =   203
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   127
         TabIndex        =   19
         Top             =   345
         Width           =   1905
      End
   End
   Begin MSComDlg.CommonDialog dlgDisplay 
      Left            =   3600
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Bitmaps (*.bmp)|*.bmp"
   End
   Begin VB.Label lblBorderWidth 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Border Width"
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblBackground 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Picture"
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblStyle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Style"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblScheme 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Scheme"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblColour 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Colour"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "frmDisplay"
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
'TITLE :    Change Clock Display Settings
' -----------------------------------------------
'COMMENTS :
'This form is used to change the display settings
'for the form frmClock and generally manage any
'graphical data for the program and save any
'changed settings
'=================================================

'require variable declaration
Option Explicit

'------------------------------------------------
'               MODULE-LEVEL VARIABLES
'------------------------------------------------
Private WithEvents mclkPreview  As clsTimedClock    'displays preview information
Attribute mclkPreview.VB_VarHelpID = -1

'the clock display settings
Private mlngAnaBack         As Long                 'the analog background colour
Private mlngDots            As Long                 'the colour of the dots
Private mlngHour            As Long                 'the colour of the hour hand
Private mlngMinute          As Long                 'the colour of the minute hand
Private mlngSecond          As Long                 'the colour of the second hand
Private mlngTimeBack        As Long                 'the background colour of the Time panel
Private mlngTimeFont        As Long                 'the font colour of the Time panel
Private mlngDayBack         As Long                 'the background colour of the Day panel
Private mlngDayFont         As Long                 'the font colour of the Day panel
Private mlngDateBack        As Long                 'the background colour of the Date panel
Private mlngDateFont        As Long                 'the font colour of the Date panel
Private mlngBorderColour    As Long                 'the border colour around the clock (if any)
Private mstrBackPath        As String               'the path of the background picture
Private menmBackStyle       As EnmCBackgroundStyle  'do we display the background, and how do we display it (eg, tile, stretch etc)
Private mintBorderWidth     As Integer              'the width of the border around the clock

'the colour schemes
Private mudtScheme()        As TypeSchemes          'holds a list of all the current colour schemes
Private mblnNewScheme       As Boolean              'a flag that is set if the user has tried to enter a new scheme

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Private Sub cboBorder_Change()
    'set the border width for the clock
    mintBorderWidth = cboBorder.ListIndex
    mclkPreview.BorderWidth = cboBorder.ListIndex
    Call mclkPreview.PaintClock
End Sub

Private Sub cboBorder_Click()
    'set the border width for the clock
    Call cboBorder_Change
End Sub

Private Sub cboColour_KeyDown(KeyCode As Integer, Shift As Integer)
    'don't let the user enter any data
    KeyCode = 0
End Sub

Private Sub cboScheme_Change()
    'don't allow the user to enter more than 50
    'characters
    If Len(cboScheme.Text) > 50 Then
        'truncate any extra characters
        cboScheme.Text = Left(cboScheme.Text, 50)
    Else
        'assume the user is typing a new scheme name to
        'save
        mblnNewScheme = True
    End If
End Sub

Private Sub cboStyle_Change()
    'update to the currently selected style
    menmBackStyle = cboStyle.ListIndex
    
    Call UpdatePreview
End Sub

Private Sub cboStyle_Click()
    Call cboStyle_Change
End Sub

Private Sub cboStyle_KeyPress(KeyAscii As Integer)
    'allow use of the arrow keys, but not characters
    Call cboStyle_Change
    KeyAscii = 0
End Sub

Private Sub cmdBrowse_Click()
    'ask the user for a new picture
    
    Const BROWSE_FLAGS  As String = cdlOFNFileMustExist + _
                                    cdlOFNHideReadOnly + _
                                    cdlOFNNoChangeDir
    
    With dlgDisplay
        .FLAGS = BROWSE_FLAGS
        .InitDir = modStrings.GetFilePath(txtPath.Text)
        .FileName = modStrings.GetFilePath(txtPath.Text, False)
        .ShowOpen
        If Dir(.FileName) = "" Then
            'use old path (eg, if the cancel button was pressed)
            txtPath.Text = mstrBackPath
        Else
            'set the full path specified
            txtPath.Text = .FileName
        
            'display the new picture
            mstrBackPath = .FileName
            
            Call UpdatePreview
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    'hide this form
    Unload Me
End Sub

Private Sub cmdColourSet_Click()
    'ask the user for a new colour
    With dlgDisplay
        .Color = cboColour.ItemData(cboColour.ListIndex)
        .ShowColor
        
        cboColour.ItemData(cboColour.ListIndex) = .Color
    End With
    
    'update the appropiate colour
    Call UpdateCurrentColour
    Call UpdatePreview
End Sub

Private Sub cmdDelScheme_Click()
    'we cannot delete the default scheme (scheme #0)
    If cboScheme.ListIndex > 0 Then
        'delete the selected scheme
        If DeleteScheme(cboScheme.ListIndex) Then
            'update the data
            Call PopulateSchemes
        End If
    End If
End Sub

Private Sub cmdOk_Click()
    'update all data
    
    'hide the form while working
    Me.Visible = False
    DoEvents
    
    'update the clock
    Call UpdateClock
    
    'exit this form
    Unload Me
End Sub

Private Sub cmdSaveScheme_Click()
    'create a new scheme if the user has changed the
    'name of the scheme, otherwise update the current
    'scheme and save it
    
    Dim udtNew      As TypeSchemes      'holds new scheme information
    Dim msgResult   As VbMsgBoxResult   'holds the users response to the message box
    
    'we cannot overwrite the default scheme
    If (Not mblnNewScheme) And _
       (cboScheme.ListIndex = 0) Then
        Exit Sub
    End If
    
    'store current scheme data a new scheme
    With udtNew
        .strName = cboScheme.Text
        .lngAnalogBack = mlngAnaBack
        .lngDots = mlngDots
        .lngHourHand = mlngHour
        .lngMinuteHand = mlngMinute
        .lngSecondHand = mlngSecond
        .lngTimeFont = mlngTimeFont
        .lngTimeBack = mlngTimeBack
        .lngDayFont = mlngDayFont
        .lngDayBack = mlngDayBack
        .lngDateFont = mlngDateFont
        .lngDateBack = mlngDateBack
        .lngBorderColour = mlngBorderColour
    End With
        
    If mblnNewScheme Then
        'add the scheme to the file
        Call AddRecord(gstrSchemePath, _
                       udtNew)
        
        'the scheme has been saved enter it into the
        'array and flag the scheme ok
        ReDim Preserve mudtScheme(UBound(mudtScheme) + 1)
        mudtScheme(UBound(mudtScheme)) = udtNew
        
        'enter the new value into the list
        mblnNewScheme = False
        Call cboScheme.AddItem(Trim(mudtScheme(UBound(mudtScheme)).strName))
        cboScheme.ListIndex = UBound(mudtScheme)
    Else
        'update current scheme
        'add the scheme to the file
        Call AddRecord(gstrSchemePath, _
                       udtNew, _
                       cboScheme.ListIndex)
    End If
    
    'display a message to the user that the scheme was saved
    msgResult = MsgBox("Colour scheme saved")
End Sub

Private Sub cmdSchemeSet_Click()
    'apply the currently selected scheme to the preview
    'box
    
    Dim intCounter As Integer   'used to cycle through the colours so that they can be updated
    
    If cboScheme.ListIndex < 0 Then
        'no scheme selected
        Exit Sub
    End If
    
    With mudtScheme(cboScheme.ListIndex)
        'update the colours combo box
        cboColour.ItemData(0) = .lngAnalogBack
        cboColour.ItemData(1) = .lngDots
        cboColour.ItemData(2) = .lngHourHand
        cboColour.ItemData(3) = .lngMinuteHand
        cboColour.ItemData(4) = .lngSecondHand
        cboColour.ItemData(5) = .lngTimeFont
        cboColour.ItemData(6) = .lngDayFont
        cboColour.ItemData(7) = .lngDateFont
        cboColour.ItemData(8) = .lngTimeBack
        cboColour.ItemData(9) = .lngDayBack
        cboColour.ItemData(10) = .lngDateBack
        cboColour.ItemData(11) = .lngBorderColour
        
        'update the variables
        For intCounter = 0 To 10
            Call UpdateCurrentColour(intCounter)
        Next intCounter
        
        'apply the colours to the clock
        Call UpdatePreview
    End With
End Sub

Private Sub Form_Activate()
    'display the clock
    mclkPreview.Visible = True
End Sub

Private Sub Form_Load()
    'create the intial settings
    
    Dim bmpInfo         As clsBitmap    'just used to get supported file types by this object as it is used by the Clock objects for graphical manipulation
    Dim intOldWidth     As Integer      'holds the size of the picture box before it's picture is resized
    Dim intOldHeight    As Integer      'holds the size of the picture box before it's picture is resized
    
    'make sure that the picture of the title bar is sized correctly. For some reason
    'this is sized smaller on some xp screens.
    With picTitleBar
        'get the size of the picture
        intOldWidth = .Width
        intOldHeight = .Height
        
        'change to the size that we want
        .Height = 270
        .Width = 1815
        Call .PaintPicture(.Picture, 0, 0, .Width, .Height, _
                                     0, 0, intOldWidth, intOldHeight)
    End With    'picTitleBar
    
    'create the preview clock
    Set mclkPreview = New clsTimedClock
    
    'create a bitmap object to get some info
    Set bmpInfo = New clsBitmap
    
    'set the border style on the background picture box to "raised". There is no property setting for
    'this so it'll have to be done via api. This is to simulate the standard window border for the
    'preview pane.
    With picFrame
        Call bmpInfo.DrawBorder(, , .ScaleWidth, .ScaleHeight, BDR_OUTLINE, , .hdc)
    End With    'picFrame
    
    
    With frmClock.mtscAnalog
        'apply other settings so that the preview
        'initially appears identical to the actual clock
        Set mclkPreview.Font = .Font
        mclkPreview.PicturePath = .PicturePath
        mclkPreview.DisplayBackground = .DisplayBackground
        mclkPreview.ShowAnalog = True   'always display this
        mclkPreview.Time24Hour = .Time24Hour
        mclkPreview.Height = picPreview.ScaleHeight
        mclkPreview.Width = picPreview.ScaleWidth
        mclkPreview.BorderColour = .BorderColour
        mclkPreview.BorderWidth = .BorderWidth
    End With
    
    'set the common dialog control flags and settings
    With dlgDisplay
        .InitDir = gstrBackPath
        .FileName = gstrBackPath
        .DefaultExt = "*.bmp"
        .FILTER = "Supported Image Files (" & _
                  bmpInfo.SupportedFormats & ")|" & _
                  bmpInfo.SupportedFormats
        .FLAGS = cdlCCFullOpen Or _
                 cdlCCRGBInit Or _
                 cdlOFNExplorer Or _
                 cdlOFNFileMustExist Or _
                 cdlOFNPathMustExist
    End With    'dlgDisplay
    
    'if this is a windows 2000 or xp machine, then the Wallpaper mode will not function properly so we
    'need to hide this from the user
    If IsW2000 Then
    
        'is the last item in the list "Wallpaper"
        If (UCase(cboStyle.List(cboStyle.ListCount - 1)) = UCase("Wallpaper")) Then
            'remove the item
            Call cboStyle.RemoveItem(cboStyle.ListCount - 1)
        End If  'is the last item in the list "Wallpaper"
    End If  'is this a windows 2000 or xp machine
    
    'set the form fonts to the current system default
    Call SetFormFontsToSystem(Me, FNT_MESSAGE)
    
    'make sure that the text box is the same height as the combo box so that the text fits in it
    txtPath.Height = cboScheme.Height
    
    'get the colour schemes
    Call PopulateSchemes
    
    'get the current settings and preview them
    Call DisplayCurrentSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'flag to clear this form from memory
    Set mclkPreview = Nothing
    gblnDoCleanUp = True
End Sub

Private Sub mclkPreview_NewTime(ByVal NewTime As Date)
    'update the display for the clock
    
    'is the form minimized
    If (Me.WindowState <> vbMinimized) Then
        Call mclkPreview.PaintClock
    End If  'is the form minimized
End Sub

Private Sub picFrame_Paint()
    'redraw the border if the control needs repainting
    
    Dim bmpInfo     As clsBitmap        'holds a referece to the Bitmap object used to draw the frame
    
    'create the bitmap object
    Set bmpInfo = New clsBitmap
    
    'set the border style on the background picture box to "raised". There is no property setting for
    'this so it'll have to be done via api.
    With picFrame
        Call bmpInfo.DrawFrame(0, 0, .ScaleWidth, .ScaleHeight, DFC_BUTTON, DFCS_BUTTONPUSH, .hdc)
    End With    'picFrame
    
    'free up memory
    Set bmpInfo = Nothing
End Sub

Private Sub picPreview_Paint()
    'update the display if somthing is moved over the
    'picture box
    Call mclkPreview.PaintClock
End Sub

Private Sub timPreview_Timer()
    'rebuild the clock display
    mclkPreview.Refresh
End Sub

Private Sub txtPath_GotFocus()
    'highlight the text
    Call HighLight(txtPath)
End Sub

Private Sub txtPath_KeyDown(KeyCode As Integer, Shift As Integer)
    'don't let the user enter any data
    KeyCode = 0
End Sub

Private Sub DisplayCurrentSettings()
    'displays the current settings in all the appropiate
    'controls and updates the preview box accordingly

    'set the variables with the current settings
    mlngAnaBack = glngColAnaBack
    mlngDots = glngColDots
    mlngHour = glngColHour
    mlngMinute = glngColMinute
    mlngSecond = glngColSecond
    mlngTimeFont = glngColTimeFont
    mlngTimeBack = glngColTimeBack
    mlngDayFont = glngColDayFont
    mlngDayBack = glngColDayBack
    mlngDateFont = glngColDateFont
    mlngDateBack = glngColDateBack
    mlngBorderColour = glngColBorder
    menmBackStyle = genmBackStyle
    mstrBackPath = gstrBackPath
    
    'apply the current settings to the controls
    With cboColour
        .ItemData(0) = glngColAnaBack
        .ItemData(1) = glngColDots
        .ItemData(2) = glngColHour
        .ItemData(3) = glngColMinute
        .ItemData(4) = glngColSecond
        .ItemData(5) = glngColTimeFont
        .ItemData(6) = glngColDayFont
        .ItemData(7) = glngColDateFont
        .ItemData(8) = glngColTimeBack
        .ItemData(9) = glngColDayBack
        .ItemData(10) = glngColDateBack
        .ItemData(11) = glngColBorder
        
        'set to the display the first element
        .ListIndex = 0
    End With
    txtPath.Text = gstrBackPath
    
    'set the style if it exists in the combo box (it might not if this is a 2000 or xp machine - see Form_Load)
    If (genmBackStyle >= cboStyle.ListCount) Then
        'default to "none"
        genmBackStyle = clkNone
    End If
    
    'set the style
    cboStyle.ListIndex = genmBackStyle
    
    'set the border width
    mintBorderWidth = gintBorderWidth
    cboBorder.ListIndex = mintBorderWidth
    
    'update the preview box
    Call UpdatePreview
End Sub

Private Sub UpdatePreview()
    'update the preview box with the current settings
    'in the controls
    
    With frmClock.mtscAnalog
        mclkPreview.SurphaseDC = picPreview.hdc
        mclkPreview.AutoDisplay = True
        mclkPreview.AnalogBackColour = mlngAnaBack
        mclkPreview.DotColour = mlngDots
        mclkPreview.HandHourColour = mlngHour
        mclkPreview.HandMinuteColour = mlngMinute
        mclkPreview.HandSecondColour = mlngSecond
        mclkPreview.TimeFontColour = mlngTimeFont
        mclkPreview.TimeBackColour = mlngTimeBack
        mclkPreview.DayFontColour = mlngDayFont
        mclkPreview.DayBackColour = mlngDayBack
        mclkPreview.DateFontColour = mlngDateFont
        mclkPreview.DateBackColour = mlngDateBack
        mclkPreview.BorderColour = mlngBorderColour
        Call mclkPreview.GetScreenPos(Me)
        mclkPreview.BackgroundStyle = menmBackStyle
        mclkPreview.PicturePath = mstrBackPath
        mclkPreview.BorderWidth = mintBorderWidth
        
        'make sure the changes are reflected
        Call .Refresh
        Call .PaintClock
    End With
End Sub

Private Sub UpdateClock()
    'This will update the main clock variables with all
    'the current settings from this form.
    
    'update the background settings
    gstrBackPath = mstrBackPath
    genmBackStyle = menmBackStyle
    
    'update the colours
    glngColAnaBack = mlngAnaBack
    glngColDots = mlngDots
    glngColHour = mlngHour
    glngColMinute = mlngMinute
    glngColSecond = mlngSecond
    glngColTimeBack = mlngTimeBack
    glngColDayBack = mlngDayBack
    glngColDateBack = mlngDateBack
    glngColTimeFont = mlngTimeFont
    glngColDayFont = mlngDayFont
    glngColDateFont = mlngDateFont
    glngColBorder = mlngBorderColour
    gintBorderWidth = mintBorderWidth
    
    'set the display
    Call SaveSettings
    Call SetAllSettings
End Sub

Private Sub UpdateCurrentColour(Optional ByVal intIndex As Integer = -1)
    'update the appropiate variable depending on which
    'element is selected
    
    Dim intData As Integer  'the index number to update from
    
    'did the programmer specify an index?
    If intIndex >= 0 Then
        'use programmer defined index
        intData = intIndex
    Else
        'use currently selected index
        intData = cboColour.ListIndex
    End If
    
    With cboColour
        Select Case intData
        Case 0  'analog background
            mlngAnaBack = .ItemData(intData)
        Case 1  'dot colour
            mlngDots = .ItemData(intData)
        Case 2  'hour hand
            mlngHour = .ItemData(intData)
        Case 3  'minute hand
            mlngMinute = .ItemData(intData)
        Case 4  'second hand
            mlngSecond = .ItemData(intData)
        Case 5  'time font
            mlngTimeFont = .ItemData(intData)
        Case 6  'day font
            mlngDayFont = .ItemData(intData)
        Case 7  'date font
            mlngDateFont = .ItemData(intData)
        Case 8  'time background
            mlngTimeBack = .ItemData(intData)
        Case 9  'day background
            mlngDayBack = .ItemData(intData)
        Case 10 'date background
            mlngDateBack = .ItemData(intData)
        Case 11 'border colour
            mlngBorderColour = .ItemData(intData)
        End Select
    End With
End Sub

Private Sub PopulateSchemes()
    'This procedure will populate the schemes combo box
    'with a list of all the stored colour schemes and
    'also enter them into the array
    
    Dim intCounter  As Integer  'used to help populate the combo box
    Dim intMatch    As Integer  'holds the index of a matching colour scheme (if found)
    
    'get all the colour schemes
    If Not GetAllRecords(mudtScheme, _
                         AddToPath(App.Path, _
                                   SCHEME_NAME)) Then
        'could not read the colour schemes
        Exit Sub
    End If
    
    'if only no records exist, then create the default
    If mudtScheme(0).strName = String(50, vbNullChar) Then
        Call CreateScheme(mudtScheme(0), True)
    End If
    
    'enter the colour schemes into the combo box
    With cboScheme
        .Clear
        For intCounter = 0 To (UBound(mudtScheme))
            'add the name of the scheme
            Call .AddItem(Trim(mudtScheme(intCounter).strName))
        Next intCounter
    
        'set the focus to display the first scheme
        intMatch = MatchScheme
        If intMatch < 0 Then
            'set the index to the default
            .ListIndex = 0
        Else
            .ListIndex = intMatch
        End If
    End With
End Sub

Private Function MatchScheme() As Integer
    'This will return the combo box index of the current colour
    'scheme for the clock (if a colour scheme exists). If the
    'function cannot find a matching colour scheme, then it
    'returnes -1
    
    Dim intCounter  As Integer  'used to cycle through the colour schemes
    Dim intMatch    As Integer  'holds the colour scheme that matches the current colours
    
    'the default value is -1
    intMatch = -1

    'search through all existing colour schemes
    For intCounter = (LBound(mudtScheme)) To (UBound(mudtScheme))
        With mudtScheme(intCounter)
            If (.lngAnalogBack = glngColAnaBack) And _
               (.lngDots = glngColDots) And _
               (.lngHourHand = glngColHour) And _
               (.lngMinuteHand = glngColMinute) And _
               (.lngSecondHand = glngColSecond) And _
               (.lngTimeFont = glngColTimeFont) And _
               (.lngTimeBack = glngColTimeBack) And _
               (.lngDayFont = glngColDayFont) And _
               (.lngDayBack = glngColDayBack) And _
               (.lngDateFont = glngColDateFont) And _
               (.lngDateBack = glngColDateBack) And _
               (.lngBorderColour = glngColBorder) Then
                'match found
                intMatch = intCounter
                Exit For
            End If
        End With
    Next intCounter
    
    'return the result of the search
    MatchScheme = intMatch
End Function
