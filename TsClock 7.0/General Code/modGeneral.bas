Attribute VB_Name = "modGeneral"
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     3 September 2003
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    General Code Module
' -----------------------------------------------
'COMMENTS :
'This was made to hold various procedures that
'do not fall under any particular category but
'are usefull in many programs.
'=================================================

'require variable declaration
Option Explicit


'-------------------------------------------------
'                   PROCEDURES
'-------------------------------------------------
Public Sub UnloadAll(Optional ByRef frmUnloadLast As Form = Nothing)
    'This will unload all the forms in the program, with the specified
    'form unloading last
    
    Dim frmFormCounter      As Form     'used to cycle through the Forms collection when unloading
    
    'cycle through all the forms in the project
    For Each frmFormCounter In Forms
        
        'first make sure that this form has not been set to Nothing
        'as this is sometimes necessary to clear memory
        If Not frmFormCounter Is Nothing Then
            
            'make sure that this form is not the form that we want
            'to unload last
            If Not frmUnloadLast Is Nothing Then
                If (frmFormCounter.Name <> frmUnloadLast.Name) Then
                    Unload frmFormCounter
                End If
                
            Else
                'just unload the form - it doesn't match the one we
                'want to unload last
                Unload frmFormCounter
            End If  'is there a form to unload last
        End If  'is there a form to unload
    Next frmFormCounter
    
    'unload the last form is one was specified
    If Not frmUnloadLast Is Nothing Then
        Unload frmUnloadLast
    End If
End Sub

Public Sub MoveTo(ByRef ctlNext As Control)
    'This procedure will move the focus to the specified control
    
    If ctlNext Is Nothing Then
        'invalid control object
        Exit Sub
    End If
    
    'determine what kind of control was passed but ignore any errors
    'generated in case that a SetFocus or property is used that does
    'not exist for the control in question
    On Error GoTo Err_Trap
    Select Case UCase(TypeName(ctlNext))
    Case "TEXTBOX", "MASKEDINPUT"
        'highlight the text before setting focus
        With ctlNext
            .SelStart = 0
            .SelLength = Len(.Text)
            Call .SetFocus
        End With
    
    Case "LISTBOX", "COMBOBOX"
        'select the first item in the list
        With ctlNext
            If (.ListCount > 0) Then
                .ListIndex = 0
            End If
            
            Call .SetFocus
        End With
    
    Case Else
        'just attempt to set the focus
        Call ctlNext.SetFocus
    End Select
    
    
'don't do any thing with the errors, just exit
Err_Trap:
End Sub
