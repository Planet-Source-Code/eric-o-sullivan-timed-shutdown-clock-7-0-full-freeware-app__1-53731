VERSION 5.00
Begin VB.Form frmAboutScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Information"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAboutScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timText 
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1740
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox picText 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Line lnSpacer 
      X1              =   120
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "frmAboutScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'require variable declaration
Option Explicit

'This screen was first created on the 17/11/2001 and was intended
'for use in several future programs. The idea was that I should only
'have to create this screen once and be able to integrate it into any
'other project seemlessly. I wanted to do this instead of creating a
'new about screen for every project where I wanted one.
'
'A note on this About Screen :
'This screen requires the module APIGraphics (APIGraphics.bas) to
'operate the display.
'
'Eric O'Sullivan
'email DiskJunky@hotmail.com
'============================================================

'Used to keep track of how many milliseconds have
'elapsed. Typically used for controlling frame rates
Private Declare Function GetTickCount _
        Lib "kernel32" () _
                        As Long

Private mstrAllText As String   'the scrolling text displayed
Private mblnStart As Boolean    'is the text starting to scroll from the bottom

Private Sub cmdOk_Click()
    'exit screen
    Unload Me
End Sub

Private Sub Form_Load()
    Call SetText
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub timText_Timer()
    'This timer will scroll the animated text
    
    Const WAIT = 50 'wait 15 ticks before drawing the next frame
    
    Dim udtBmp As clsBitmap
    Dim intTextHeight As Integer
    
    Static lngStartingTick As Long
    Static udtSurphase As clsBitmap
    Static intScroll As Integer
    Static intLineNum As Integer
    Static udtMask As clsBitmap
    
    'instanciate the bitmaps
    Set udtBmp = New clsBitmap
    
    'find out how much time it takes to draw a frame
    If lngStartingTick = 0 Then
        lngStartingTick = GetTickCount
    End If
    
    'set the bitmap dimensions and create them
    Call udtBmp.SetBitmap(picText.ScaleHeight, _
                          picText.ScaleWidth, _
                          vbWhite)
    'test code - not currently used
    'Call MakeText(picText.hDc, "Hello World!", 0, 0, 40, 180, udtFont, InPixels)
    
    'get the number of lines of text to display
    If intLineNum = 0 Then
        intLineNum = LineCount(mstrAllText)
    End If
    
    intTextHeight = picText.TextHeight("I") * intLineNum
    
    intScroll = intScroll - 1
    If (intScroll < (-intTextHeight)) _
       Or (Not mblnStart) Then
        intScroll = picText.ScaleHeight
        mblnStart = True
    End If
    
    'only create the surphase if necessary
    If udtSurphase Is Nothing Then
        Set udtSurphase = New clsBitmap
        
        Call udtSurphase.SetBitmap(picText.ScaleHeight, _
                                   picText.ScaleWidth, _
                                   vbYellow)
        
        'create the surphase
        'text fade in
        Call udtSurphase.Gradient(picText.ForeColor, _
                                  picText.FillColor, _
                                  0, _
                                  (udtSurphase.Height - ((intTextHeight / intLineNum) * 2)), _
                                  udtSurphase.Width, _
                                  (intTextHeight / intLineNum * 2), _
                                  GradHorizontal)
        'text fade out
        Call udtSurphase.Gradient(picText.FillColor, _
                                  picText.ForeColor, _
                                  0, _
                                  0, _
                                  udtSurphase.Width, _
                                  (intTextHeight / intLineNum) * 2, _
                                  GradHorizontal)
    End If
    
    'we only need to create the text once
    If udtMask Is Nothing Then
        Set udtMask = New clsBitmap
        
        Call udtMask.SetBitmap(intTextHeight + (picText.ScaleHeight * 2), _
                               picText.ScaleWidth)
    
        'draw the white text onto the mask in black
        Call udtMask.DrawString(mstrAllText, _
                                picText.ScaleHeight, _
                                0, _
                                intTextHeight, _
                                udtBmp.Width, _
                                picText.Font, _
                                vbWhite)
    End If
    
    'copy the surphase onto the background
    Call udtSurphase.Paint(udtBmp.hDc)
    
    'place the mask onto the background
    Call udtMask.Paint(udtBmp.hDc, _
                       0, _
                       intScroll - picText.ScaleHeight, _
                       udtMask.Height, _
                       udtMask.Width, _
                       lngPaintMode:=P_AND)
    
    'copy the result to the screen
    Call udtBmp.Paint(Me.hDc)
    
    'wait X ticks minus the time it took to draw the frame
    Call udtSurphase.Pause(WAIT - (GetTickCount - lngStartingTick), True)
    lngStartingTick = GetTickCount
End Sub

Private Sub SetText()
    'This procedure is used to setting the text displayed in the picture box
    
    '" & vbCrLf & "
    
    'please note that ProductName can be set by going to
    'Project, Project Properties,Make tab. You should see a list box about
    'half way down on the left side. Scroll down until you come to
    'Product Name and enter some text into the text box on the right
    'side of the list box.
    mstrAllText = App.ProductName & vbCrLf & _
                  "Version " & App.Major & "." & _
                               App.Minor & "." & _
                               App.Revision & vbCrLf & _
                  "" & vbCrLf & _
                  "This program was made by" & vbCrLf & _
                  "Eric O'Sullivan." & vbCrLf & _
                  "" & vbCrLf & _
                  "Copyright 2002" & vbCrLf & _
                  "All rights reserved" & vbCrLf & _
                  "" & vbCrLf & _
                  "For more information, email" & vbCrLf & _
                  "DiskJunky@hotmail.com"
End Sub

Private Function LineCount(ByVal strText As String) _
                           As Integer
    'This function will return the number of lines
    'in the strText
    
    Dim intTemp As Integer
    Dim intCounter As Integer
    Dim intLastPos As Integer
    
    intLastPos = 1
    
    Do
        intTemp = intLastPos
        intLastPos = InStr(intLastPos + Len(vbCrLf), strText, vbCrLf)
        
        If intTemp <> intLastPos Then
            'a line was found
            intCounter = intCounter + 1
        End If
    Loop Until intLastPos = 0 'intLastPos will =0 when InStr cannot find any more occurances of vbCrlf
    
    LineCount = intCounter + 1
End Function
