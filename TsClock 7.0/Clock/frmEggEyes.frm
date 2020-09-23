VERSION 5.00
Begin VB.Form frmEggEyes 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Screen Eyes"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   FillColor       =   &H80000008&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timUpdate 
      Interval        =   50
      Left            =   720
      Top             =   1320
   End
End
Attribute VB_Name = "frmEggEyes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     5 November 2002
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    Moving Eyes Screen (Egg)
' -----------------------------------------------
'COMMENTS :
'This screen is only used as a fun extra feature
'for the program should anyone be able to find
'it.
'=================================================

'require variable declaration
Option Explicit

'------------------------------------------------
'               MODULE-LEVEL VARIABLES
'------------------------------------------------
Private mbmpEyes    As clsBitmap    'holds the surphase on which the eyes are drawn
Private mmosGrab    As clsMouse     'used to grab the window for moving

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Private Sub Form_DblClick()
    'exit the form
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'exit the user if they press [RETURN] or [ESC]
    Select Case KeyAscii
    Case vbKeyEscape, vbKeyReturn
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    'setup the eyes bitmap settings
    Set mbmpEyes = New clsBitmap
    Set mmosGrab = New clsMouse
    With mbmpEyes
        Call .SetBitmap(Me.Width / Screen.TwipsPerPixelX, _
                        Me.Height / Screen.TwipsPerPixelY, _
                        Me.BackColor)
    End With
    
    'make sure this form is visible
    Call StayOnTop(Me)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'grab the window for dragging
    'Call mmosGrab.GrabWindow(Me.hWnd)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'release window from dragging
    'Call mmosGrab.ReleaseWindow
End Sub

Private Sub Form_Paint()
    'draw the eyes on the form
    Call mbmpEyes.Paint(Me.hDc)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'make sure we clean up
    Call NotOnTop(Me)
    Unload Me
    
    'flag to clear this form from memory
    gblnDoCleanUp = True
End Sub

Private Sub DrawEyes(Optional ByVal intX As Integer = -1, _
                     Optional ByVal intY As Integer = -1)
    'This will draw eyes looking in the direction of the mouse on the screen
    
    Const PUPIL_COLOUR  As Long = &HFF8080    'light blue
    Const SCLERA_COLOUR As Long = vbWhite
    Const X             As Integer = 0
    Const Y             As Integer = 1
    Const PI            As Single = 3.14159265358979
    
    Dim intTemp         As Integer  'holds a temperory mouse position
    Dim intLeftX        As Integer  'holds the X position of the centre of the left eye
    Dim intRightX       As Integer  'holds the X position of the centre of the right eye
    Dim intScleraY      As Integer  'holds the Y position of the centre of the eyes
    Dim intLeftIris(1)  As Integer  'holds the position of the left iris
    Dim intRightIris(1) As Integer  'holds the position of the right iris
    Dim intLeftRad(1)   As Integer  'holds the radius for the left iris
    Dim intRightRad(1)  As Integer  'holds the radius for the right iris
    Dim intScleraHeight As Integer  'holds the total height of the eye
    Dim intScleraWidth  As Integer  'holds the total width of the eye
    Dim intIrisHeight   As Integer  'holds the height of the iris
    Dim intIrisWidth    As Integer  'holds the width of the iris
    Dim intLeftAngle    As Integer  'holds the angle of the mouse from the left eye
    Dim intRightAngle   As Integer  'holds the angle of the mouse from the right eye
    Dim intAbsX         As Integer  'holds the distance of the mouse from one of the eyes
    Dim intAbsY         As Integer  'holds the distance of the mouse from one of the eyes
    Dim intFormHeight   As Integer  'holds the height in pixels of the form
    Dim intFormWidth    As Integer  'holds the width in pixels of the form
    Dim intFormTop      As Integer  'holds the Top position of the form in pixels
    Dim intFormLeft     As Integer  'holds the Left position of the form in pixels
    
    'if no parameters were passed, then default to the screen mouse position
    With mbmpEyes
        If intX < 0 Then
            Call .MousePosition(intX, intTemp)
        End If
        If intY < 0 Then
            Call .MousePosition(intTemp, intY)
        End If
    End With
    
    'convert the forms twip metrics to pixels
    With Me
        intFormLeft = .Left \ Screen.TwipsPerPixelX
        intFormTop = .Top \ Screen.TwipsPerPixelY
        intFormWidth = .Width \ Screen.TwipsPerPixelX
        intFormHeight = .Height \ Screen.TwipsPerPixelY
    End With
    
    'the left eye is 1/3 the width from the left edge of the form
    intLeftX = intFormWidth \ 3
    
    'the right eye it 1/3 the width from the right edge
    intRightX = (intFormWidth \ 3) * 2
    
    'get the size of the eyes in relation to the size of the form
    intScleraY = (intFormHeight \ 2)
    intScleraWidth = intFormWidth \ 4
    intScleraHeight = (intFormHeight \ 3) * 2
    intIrisWidth = intScleraWidth \ 3
    intIrisHeight = intScleraHeight \ 3
    
    'get the mouse angle in relation to each eye
    With mbmpEyes
        intLeftAngle = .GetAngle(intFormLeft + intLeftX, _
                                 intFormTop + intScleraY, _
                                 intX, _
                                 intY)
        intRightAngle = .GetAngle(intFormLeft + intRightX, _
                                  intFormTop + intScleraY, _
                                  intX, _
                                  intY)
        
        'adjust to "real" angle
        intLeftAngle = (intLeftAngle + 270) Mod 360
        intRightAngle = (intRightAngle + 270) Mod 360
    End With
    
    'calculate the radius for each iris relative to the distance
    'of the mouse from the centre of the eye - this is what moves
    'the eyes to the edge of the eye.
    intAbsX = intFormWidth + ((intLeftX / 4) * 3)
    intAbsY = intFormTop + intScleraY
    If intAbsX <> 0 Then
        'the total distance the iris moves
        intLeftRad(X) = ((intScleraWidth - intIrisWidth) / 2)
        
        'in relation to the total horizontal distance the mouse moves
        intLeftRad(X) = (intLeftRad(X) * (Abs(intAbsX - intX) / intAbsX))
    End If
    If intAbsY <> 0 Then        'this is the same for both eyes
        'the total distance the iris moves
        intLeftRad(Y) = ((intScleraHeight - intIrisHeight) / 2)
        
        'in relation to the total vertical distance the mouse moves
        intLeftRad(Y) = (intLeftRad(Y) * (Abs(intAbsY - intY) / intAbsY))
        
        'this also applies to the right eye as they are on the same
        'horizontal plane
        intRightRad(Y) = intLeftRad(Y)
    End If
    intAbsX = intFormLeft + ((intRightX / 4) * 3)
    If intAbsX <> 0 Then
        'the total distance the iris moves
        intRightRad(X) = ((intScleraWidth - intIrisWidth) / 2)
        
        'in relation to the total horizontal dirance the mouse moves
        intRightRad(X) = (intRightRad(X) * (Abs(intX - intAbsX) / intAbsX))
    End If
    
    'calculate the position of each iris
    intLeftIris(X) = intLeftX - (Sin(intLeftAngle * PI / 180) * _
                                 intLeftRad(X))
    intLeftIris(Y) = intScleraY + (Cos(intLeftAngle * PI / 180) * _
                                   intLeftRad(Y))
    intRightIris(X) = intRightX - (Sin(intRightAngle * PI / 180) * _
                                   intRightRad(X))
    intRightIris(Y) = intScleraY + (Cos(intRightAngle * PI / 180) * _
                                    intRightRad(Y))
    
    'draw the eyes
    With mbmpEyes
        'clear the existing display
        .Cls
        
        'draw the sclera of the eyes (the "whites" of the eyes)
        Call .DrawEllipse(intLeftX, _
                          intScleraY, _
                          intScleraHeight, _
                          intScleraWidth, _
                          90, _
                          SCLERA_COLOUR, _
                          1, _
                          False)
        Call .DrawEllipse(intRightX, _
                          intScleraY, _
                          intScleraHeight, _
                          intScleraWidth, _
                          90, _
                          SCLERA_COLOUR, _
                          1, _
                          False)
        
        'draw the irises on top of the sclera
        Call .DrawEllipse(intLeftIris(X), _
                          intLeftIris(Y), _
                          intIrisHeight, _
                          intIrisWidth, _
                          90, _
                          PUPIL_COLOUR, _
                          1, _
                          False)
        Call .DrawEllipse(intRightIris(X), _
                          intRightIris(Y), _
                          intIrisHeight, _
                          intIrisWidth, _
                          90, _
                          PUPIL_COLOUR, _
                          1, _
                          False)
    End With
End Sub

Public Sub ReDraw()
    'This will update the display on the form
    Call DrawEyes
    Call Form_Paint
End Sub

Private Sub timUpdate_Timer()
    'update the display
    If Me.Visible Then
        Call ReDraw
    End If
End Sub
