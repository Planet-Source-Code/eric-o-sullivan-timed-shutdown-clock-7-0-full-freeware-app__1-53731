VERSION 5.00
Object = "{072A0CD7-4439-46FD-9BC0-FF8959716B3B}#2.0#0"; "SYSTEM~1.OCX"
Begin VB.Form frmTestSysTray 
   Caption         =   "Test The System Tray Icon"
   ClientHeight    =   3192
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3192
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin SystemTrayIcon.ctlSysTray ctlSysTray1 
      Left            =   120
      Top             =   120
      _ExtentX        =   762
      _ExtentY        =   762
      Icon            =   "frmTestSysTray.frx":0000
      ToolTip         =   "ctlSysTray1"
   End
End
Attribute VB_Name = "frmTestSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'require variable declaration
Option Explicit

Private Sub Form_Load()
    'display the system tray icon
    Call ctlSysTray1.Show
End Sub

Private Sub ctlSysTray1_Click(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Debug.Print "Click" & Button
End Sub

Private Sub ctlSysTray1_DblClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Debug.Print "DblClick" & Button
End Sub

Private Sub ctlSysTray1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Debug.Print "MouseDown " & Button
End Sub

Private Sub ctlSysTray1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Debug.Print "MouseMove" & Button
End Sub

Private Sub ctlSysTray1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Debug.Print "MouseUp" & Button
End Sub

Private Sub ctlSysTray1_TaskbarCreated()
    Debug.Print "TaskbarCreated"
End Sub
