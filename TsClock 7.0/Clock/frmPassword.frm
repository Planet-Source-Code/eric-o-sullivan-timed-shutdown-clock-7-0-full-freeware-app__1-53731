VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Screen"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraAsk 
      Caption         =   "Password"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   4935
      Begin VB.CommandButton cmdEnter 
         Caption         =   "&Enter"
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdCan 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   34
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "C&hange"
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblEnter 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame fraChange 
      Caption         =   "Change Password"
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "&Set"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtRetype 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   34
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtNew 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   34
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtOld 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   34
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblOld 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Old Password"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblRetype 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Type Password"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblNew 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter New Password"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmPassword"
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
'TITLE :    Password Screen (Enter/Change)
' -----------------------------------------------
'COMMENTS :
'This form can either be used to ask the user for
'a password, or ask the user to change their
'password.
'=================================================

'all variables must be declared
Option Explicit

'------------------------------------------------
'               MODULE-LEVEL CONSTANTS
'------------------------------------------------
'the different sizes for the two different functions
'of this screen
Private Const CHANGE_HEIGHT As Integer = 2430   'used when changing the password
Private Const CHANGE_WIDTH  As Integer = 4305   'used when changing the password
Private Const ASK_HEIGHT    As Integer = 1590   'used when entering the password
Private Const ASK_WIDTH     As Integer = 4305   'used when entering the password

'------------------------------------------------
'               MODULE-LEVEL VARIABLES
'------------------------------------------------
Private mintOrigHeight      As Integer          'holds the oridinal height of the form

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Private Sub cmdCan_Click()
    'the cancel button for entering the password
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    'the cancel button for changing the password
    Unload Me
End Sub

Private Sub cmdChange_Click()
    Call SetScreen("Change")
End Sub

Private Sub cmdSet_Click()
    'do we set a new password
    If (txtRetype.Text <> txtNew.Text) Then
        'confirm new password again
        txtRetype.SetFocus
        txtRetype.SelStart = 0
        txtRetype.SelLength = Len(txtRetype.Text)
        Exit Sub
    Else
        If gstrPassword = txtOld.Text Then
            'set password
            gstrPassword = txtNew.Text
            gblnPassOn = True
            
            'clear all information before unloading
            txtNew.Text = ""
            txtOld.Text = ""
            txtRetype.Text = ""
            
            'save and exit
            Call SaveSettings
            Call CheckPassword
            Unload Me
        Else
            'did not match old password
            txtOld.SetFocus
        End If
    End If
End Sub

Public Sub SetScreen(Optional ByVal strAskOrChange As String = "Ask")
    'Show a different screen depending on which function
    'is needed. If no password is specified, then we
    'ask the user for one and turn off the password.
    
    If gstrPassword = "" Then
        'ask the user for a password and make sure the
        'password is currently disabled
        strAskOrChange = "Change"
        gblnPassOn = False
        Call frmClock.EnableMenus(True)
    End If
    
    Select Case LCase(strAskOrChange)
    Case "ask"
        'ask the user for the password
        'Me.Height = ASK_HEIGHT
        'Me.Width = ASK_WIDTH
        Me.Height = mintOrigHeight - (fraChange.Height + 120)
        fraAsk.Top = 120
        fraAsk.Visible = True
        fraChange.Visible = False
        
        'display the form
        Me.Visible = True
        DoEvents
        
        txtPass.SetFocus
    
    Case "change"
        'change the existing/set a new password
        'Me.Height = CHANGE_HEIGHT
        'Me.Width = CHANGE_WIDTH
        Me.Height = mintOrigHeight - (fraAsk.Height + 120)
        fraAsk.Visible = False
        fraChange.Visible = True
        
        'display the form
        Me.Visible = True
        DoEvents
        
        'if there is no password, set one.
        cmdSet.Enabled = False
        If gstrPassword = "" Then
            txtOld.Enabled = False
            txtOld.BackColor = Me.BackColor
            txtNew.SetFocus
        Else
            txtOld.Enabled = True
            txtOld.BackColor = txtNew.BackColor
            txtOld.SetFocus
        End If
    
    Case ""
        'nothing was passed, unload the form
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    'set and intiial values for the form
    
    Dim intTxtHeight    As Integer      'holds the height to set the text boxes
    
    'set the form fonts to the current system default
    Call SetFormFontsToSystem(Me, FNT_MESSAGE)
    mintOrigHeight = Me.Height
    
    Set Me.Font = txtPass.Font
    intTxtHeight = Me.TextHeight("I") + 60
    
    txtPass.Height = intTxtHeight
    txtOld.Height = intTxtHeight
    txtNew.Height = intTxtHeight
    txtRetype.Height = intTxtHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'unconditionally unload the form
    Unload Me
    
    'flag to clear this form from memory
    gblnDoCleanUp = True
End Sub

Private Sub txtNew_KeyPress(KeyAscii As Integer)
    'move to next text box
    If KeyAscii = Asc(vbCr) Then
        txtRetype.SetFocus
    End If
End Sub

Private Sub txtOld_KeyPress(KeyAscii As Integer)
    'move to the next text box
    If KeyAscii = Asc(vbCr) Then
        txtNew.SetFocus
    End If
End Sub

Private Sub txtOld_Validate(Cancel As Boolean)
    If txtOld.Text <> gstrPassword Then
        'passwords don't match
        txtOld.SetFocus
        txtOld.SelStart = 0
        txtOld.SelLength = Len(txtOld.Text)
    End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    'do we allow the user to activate or deactivate the
    'password
    If KeyAscii = Asc(vbCr) Then
        Call CheckPassword
    End If
End Sub

Private Sub txtRetype_KeyPress(KeyAscii As Integer)
    'set the new password
    If KeyAscii = Asc(vbCr) Then
        'set password
        cmdSet_Click
    End If
End Sub

Private Sub ActivateCheck(ByVal strPass As String)
    'This procedure activtes or deactivates the check
    'box.
    
    If strPass = gstrPassword Then
        Call frmClock.EnableMenus(True)
    Else
        Call frmClock.EnableMenus(False)
    End If
End Sub

Private Sub txtRetype_KeyUp(KeyCode As Integer, Shift As Integer)
    'enable or disable command button
    If txtNew.Text = txtRetype.Text Then
        cmdSet.Enabled = True
    Else
        cmdSet.Enabled = False
    End If
End Sub

Private Sub CheckPassword()
    'did the user enter the correct password
    If txtPass.Text = gstrPassword Then
        'correct password entered, set and exit
        gblnPassOn = True
        txtPass.Text = ""
        Call frmClock.EnableMenus(True)
        Unload Me
    Else
        'incorrect password
        gblnPassOn = False
        txtPass.SelStart = 0
        txtPass.SelLength = Len(txtPass.Text)
    End If
End Sub
