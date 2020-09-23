VERSION 5.00
Begin VB.UserControl ctlProgBar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   372
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1812
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   HitBehavior     =   0  'None
   PropertyPages   =   "ctlProgBar.ctx":0000
   ScaleHeight     =   31
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   151
   ToolboxBitmap   =   "ctlProgBar.ctx":002F
End
Attribute VB_Name = "ctlProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'require variable declaration
Option Explicit

'the border style of the control
Public Enum BorderEnum
    pgrNone = 0
    pgrFixed_Single = 1
End Enum

'the appearance effect of the control
Public Enum AppearanceEnum
    pgr2D = 0
    pgr3D = 1
End Enum

'property variables
Private msngMax         As Single               'the Max value of the progress bar
Private msngMin         As Single               'the Min value of the progress bar
Private msngValue       As Single               'the position of the progress bar
Private mstrCaption     As String               'the user defined caption for the progress bar
Private mblnDefCapt     As Boolean              'the defualt caption is the percentage of the progress bar taken up by Value
Private mlngBackColour  As Long                 'the initial background of the progress bar when Value = Min
Private mlngFillColour  As Long                 'the colour of the progress bar as it takes up space on the screen
Private mlngTextColour  As Long                 'the initial caption colour for the text
Private mlngOverColour  As Long                 'the colour when the progress bar moves over the caption
Private menmBorder      As BorderEnum           'the border style of the user control
Private menmAlignment   As AlignmentConstants   'holds the text alignment for the progress bar
Private menmAppearance  As AppearanceEnum       'holds if the border effect is 2d or 3d

'general variables
Private mbmpBack        As clsBitmap            'the background of the user control

'Event Declarations:
Public Event Click()                            'MappingInfo=UserControl,UserControl,-1,Click
Public Event DblClick()                         'MappingInfo=UserControl,UserControl,-1,DblClick
Public Event MouseDown(Button As Integer, _
                       Shift As Integer, _
                       X As Single, _
                       Y As Single)             'MappingInfo=UserControl,UserControl,-1,MouseDown
Public Event MouseMove(Button As Integer, _
                       Shift As Integer, _
                       X As Single, _
                       Y As Single)             'MappingInfo=UserControl,UserControl,-1,MouseMove
Public Event MouseUp(Button As Integer, _
                     Shift As Integer, _
                     X As Single, _
                     Y As Single)               'MappingInfo=UserControl,UserControl,-1,MouseUp

'properties

Public Property Get Appearance() As AppearanceEnum
Attribute Appearance.VB_Description = "Returns/sets the border effect used for the progress bar"
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
    'This will return border effect style of the control
    Appearance = menmAppearance
End Property

Public Property Let Appearance(ByVal enmNewAppearance As AppearanceEnum)
    'This will set the border style appearance of the control
    menmAppearance = enmNewAppearance
    Call Refresh
    PropertyChanged "Appearance"
End Property

Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of the text in the progress bar"
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Text"
    'This will return the text alignment for the caption of the control
    Alignment = menmAlignment
End Property

Public Property Let Alignment(ByVal enmNewAlign As AlignmentConstants)
    'This will set the text alignment for the caption of the control
    menmAlignment = enmNewAlign
    Call Refresh
    PropertyChanged "Alignment"
End Property

Public Property Get PercentCaption() As Boolean
Attribute PercentCaption.VB_Description = "If set to True, the progress bar text will display the percentage of the control covered by the progress bar"
Attribute PercentCaption.VB_ProcData.VB_Invoke_Property = ";Text"
    'return whether or not to display the
    'default caption
    PercentCaption = mblnDefCapt
End Property

Public Property Let PercentCaption(ByVal blnNewValue As Boolean)
    'set the default caption
    mblnDefCapt = blnNewValue
    Call Refresh
    PropertyChanged "PercentCaption"
End Property

Public Property Get Max() As Single
Attribute Max.VB_Description = "The upper range of the progress bar"
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Scale"
    'return the Max value
    Max = msngMax
End Property

Public Property Let Max(ByVal sngNewValue As Single)
    'set the Max value
    
    If sngNewValue < msngMin Then
        Exit Property
    End If
    
    'if the new max value is less than the current
    'progress Value, then the new max IS the current
    'max value
    If sngNewValue < msngValue Then
        msngValue = sngNewValue
    End If
    
    msngMax = sngNewValue
    Call Refresh
    PropertyChanged "Max"
End Property

Public Property Get Value() As Single
Attribute Value.VB_Description = "This sets the position of the progress bar. It must be between the Max and Min values"
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Scale"
Attribute Value.VB_MemberFlags = "200"
    'return the Value
    Value = msngValue
End Property

Public Property Let Value(ByVal sngNewValue As Single)
    'set the Value
    
    'make sure the new value is not out of the current
    'valid ranges
    Select Case sngNewValue
    Case Is > msngMax
        sngNewValue = msngMax
    Case Is < msngMin
        sngNewValue = msngMin
    End Select
    
    'apply the new value
    msngValue = sngNewValue
    Call Refresh
    PropertyChanged "Value"
End Property

Public Property Get Min() As Single
Attribute Min.VB_Description = "The lower range of the progress bar"
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Scale"
    'return the Min value
    Min = msngMin
End Property

Public Property Let Min(ByVal sngNewValue As Single)
    'set the Min value
    
    If sngNewValue > msngMax Then
        Exit Property
    End If
    
    If sngNewValue > msngValue Then
        msngValue = sngNewValue
    End If
    
    msngMin = sngNewValue
    Call Refresh
    PropertyChanged "Min"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets the text displayed in the progress bar. This is ignored if PercentCaption is set to True"
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518
    'return the caption
    Caption = mstrCaption
End Property

Public Property Let Caption(ByVal strNewValue As String)
    'set the caption
    mstrCaption = strNewValue
    Call Refresh
    PropertyChanged "Caption"
End Property

Public Property Get BackColour() As OLE_COLOR
Attribute BackColour.VB_Description = "Sets the background colour of the area not covered by the progress bar"
Attribute BackColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColour.VB_UserMemId = -501
    'get the background colour
    BackColour = mlngBackColour
End Property

Public Property Let BackColour(ByVal lngNewValue As OLE_COLOR)
    'set the background colour
    mlngBackColour = lngNewValue
    Call Refresh
    PropertyChanged "BackColour"
End Property

Public Property Get FillColour() As OLE_COLOR
Attribute FillColour.VB_Description = "This sets the colour of the progress bar"
Attribute FillColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute FillColour.VB_UserMemId = -510
    'get the fill colour
    FillColour = mlngFillColour
End Property

Public Property Let FillColour(ByVal lngNewValue As OLE_COLOR)
    'set the fill colour
    mlngFillColour = lngNewValue
    Call Refresh
    PropertyChanged "FillColour"
End Property

Public Property Get TextColour() As OLE_COLOR
Attribute TextColour.VB_Description = "This is the colour of the Caption when the progress bar is not covering it"
Attribute TextColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute TextColour.VB_UserMemId = -513
    'get the text colour
    TextColour = mlngTextColour
End Property

Public Property Let TextColour(ByVal lngNewValue As OLE_COLOR)
    'set the text colour
    mlngTextColour = lngNewValue
    Call Refresh
    PropertyChanged "TextColour"
End Property

Public Property Get OverColour() As OLE_COLOR
Attribute OverColour.VB_Description = "This set the colour of the text when the progress bar moves over it"
Attribute OverColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    'get the text over colour
    OverColour = mlngOverColour
End Property

Public Property Let OverColour(ByVal lngNewValue As OLE_COLOR)
    'set the text over colour
    mlngOverColour = lngNewValue
    UserControl.Appearance = menmBorder
    Call Refresh
    PropertyChanged "OverColour"
End Property

Public Property Get BorderStyle() As BorderEnum
Attribute BorderStyle.VB_Description = "Sets the border style of the progress bar"
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -504
    'get the border style
    BorderStyle = menmBorder
End Property

Public Property Let BorderStyle(ByVal enmNewValue As BorderEnum)
    'set the border style
    menmBorder = enmNewValue
    Call Refresh
    PropertyChanged "BorderStyle"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "This sets the font attributes of the text displayed in the progress bar"
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    'get a new font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal fntNewValue As Font)
    'set the new font
    
    Set UserControl.Font = fntNewValue
    
    Call Refresh
    PropertyChanged "Font"
End Property

'methods/procedures

Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Displays the about box for this form"
Attribute ShowAbout.VB_UserMemId = -552
    'display the about screen
    Call frmAboutScreen.Show(vbModal)
End Sub

Private Sub BuildProgressBar()
    'This will rebuild the progress bar picture
    'from scratch, using the current properties
    
    Dim bmpTempFill     As clsBitmap    'this is a bitmap of the section of the progress bar that is being filled
    Dim bmpTempBack     As clsBitmap    'this is a bitmap of the section of the progress bar that is not yet filled
    Dim strCaption      As String       'the text to display in the progress bar
    Dim sngProgress     As Single       'the amount of relative space taken up by Value
    Dim lngTextHeight   As Long         'the height of the text
    Dim lngWidthFill    As Long         'holds the width of the filled section of the progress bar
    Dim lngWidthBack    As Long         'holds the width of the UNfilled section of the progress bar
    
    With UserControl
        'set the border style if necessary
        If (.BorderStyle <> menmBorder) Then
            .BorderStyle = menmBorder
        End If
        
        'check the border effect style
        If (.Appearance <> menmAppearance) Then
            .Appearance = menmAppearance
        End If
    End With    'UserControl
    
    'get the point to fill the progress bar to
    If msngMax > 0 Then
        sngProgress = (1 / (msngMax - msngMin)) * _
                      (msngValue - msngMin)
    End If
    
    'set the text to the default if necessary
    If mblnDefCapt Then
        'get the current percentage
        strCaption = Int(sngProgress * 100) & "%"
    Else
        strCaption = mstrCaption
    End If
    
    'get the width of the filled and unfilled section
    lngWidthFill = Int(mbmpBack.Width * sngProgress)
    lngWidthBack = mbmpBack.Width - lngWidthFill
    
    'set up the bitmaps
    Set bmpTempFill = New clsBitmap
    With bmpTempFill
        Call .SetBitmap(lngWidthFill, _
                        mbmpBack.Height, _
                        .GetSystemColour(mlngFillColour))
    End With
    Set bmpTempBack = New clsBitmap
    With bmpTempBack
        Call .SetBitmap(lngWidthBack, _
                        mbmpBack.Height, _
                        .GetSystemColour(mlngBackColour))
    End With
    
    'get the text height
    lngTextHeight = UserControl.TextHeight(strCaption)
    
    'draw the text based on the selected alignment
    With bmpTempFill
        Call .DrawString(strCaption, _
                         (.Height - lngTextHeight) \ 2, _
                         0, _
                         lngTextHeight, _
                         mbmpBack.Width, _
                         UserControl.Font, _
                         .GetSystemColour(mlngOverColour), _
                         menmAlignment)
    End With
    With bmpTempBack
        Call .DrawString(strCaption, _
                         (.Height - lngTextHeight) \ 2, _
                         -lngWidthFill, _
                         lngTextHeight, _
                         mbmpBack.Width, _
                         UserControl.Font, _
                         .GetSystemColour(mlngTextColour), _
                         menmAlignment)
    End With
    
    'copy the proper sections of the bitmaps
    'onto the background bitmap
    With bmpTempFill
        'Call .Paint(mbmpBack.hDc, _
                    intWidth:=Int(.Width * sngProgress))
        Call .Paint(mbmpBack.hDc)
    End With
    With bmpTempBack
        'Call .Paint(mbmpBack.hDc, _
                    Int(.Width * sngProgress), _
                    intSourceX:=Int(.Width * sngProgress))
        Call .Paint(mbmpBack.hDc, _
                    lngWidthFill)
    End With
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Rebuilds the progress bar from scratch"
    'This will update the display on the
    'progress bar
    Call BuildProgressBar
    Call ShowProgressBar
End Sub

Private Sub ShowProgressBar()
    'display the background
    With mbmpBack
        Call .Paint(UserControl.hDc, _
                    0, _
                    0, _
                    .Height, _
                    .Width)
    End With
End Sub

'events

Private Sub UserControl_Click()
    'raise a click event for the progress bar
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    'raise a double click event for the progress bar
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Paint()
    Call ShowProgressBar
End Sub

Private Sub UserControl_Initialize()
    'set up the default values before
    'reading the property bag
    
    Dim intHeight As Integer    'the user control height
    Dim intWidth As Integer     'the user control width
    
    'create the background bitmap
    Set mbmpBack = New clsBitmap
    With mbmpBack
        .Width = UserControl.ScaleWidth - UserControl.ScaleLeft
        .Height = UserControl.ScaleHeight - UserControl.ScaleTop
        Call .SetBitmap(intWidth, intHeight)
    End With
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'raise a mouse down event for the progress bar
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'raise a mouse move event for the progress bar
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'raise a mouse up event for the progress bar
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_InitProperties()
    'set the default properties for the control
    
    menmBorder = pgrFixed_Single
    msngMax = 100
    msngMin = 0
    msngValue = 0
    mblnDefCapt = True
    mstrCaption = ""
    menmAlignment = vbCenter
    menmAppearance = pgr3D
    mlngBackColour = vbButtonFace
    mlngFillColour = vbHighlight
    mlngTextColour = vbButtonText
    mlngOverColour = vbHighlightText
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'get the properties from the property bag
    
    With PropBag
        msngMax = .ReadProperty("Max", 100)
        msngMin = .ReadProperty("Min", 0)
        msngValue = .ReadProperty("Value", 0)
        mblnDefCapt = .ReadProperty("PercentCaption", True)
        mstrCaption = .ReadProperty("Caption", Ambient.DisplayName)
        menmAppearance = .ReadProperty("Appearance", pgr3D)
        menmAlignment = .ReadProperty("Alignment", vbCenter)
        mlngBackColour = .ReadProperty("BackColour", vbButtonFace)
        mlngFillColour = .ReadProperty("FillColour", vbHighlight)
        mlngTextColour = .ReadProperty("TextColour", vbButtonText)
        mlngOverColour = .ReadProperty("OverColour", vbHighlightText)
        menmBorder = .ReadProperty("BorderStyle", pgrFixed_Single)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
    End With
    
    'refresh the progress bar display
    Call Refresh
End Sub

Private Sub UserControl_Resize()
    'set the bitmap dimensions
    With mbmpBack
        .Width = UserControl.ScaleWidth - (UserControl.ScaleLeft * 2)
        .Height = UserControl.ScaleHeight - (UserControl.ScaleTop * 2)
    End With
    Call BuildProgressBar
    Call ShowProgressBar
End Sub

Private Sub UserControl_Terminate()
    'remove the background bitmap from memory
    Set mbmpBack = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'save the current properties
    
    With PropBag
        Call .WriteProperty("Max", msngMax, 100)
        Call .WriteProperty("Min", msngMin, 0)
        Call .WriteProperty("Value", msngValue, 0)
        Call .WriteProperty("Caption", mstrCaption, Ambient.DisplayName)
        Call .WriteProperty("Appearance", menmAppearance, pgr3D)
        Call .WriteProperty("Alignment", menmAlignment, vbCenter)
        Call .WriteProperty("PercentCaption", mblnDefCapt, True)
        Call .WriteProperty("BackColour", mlngBackColour, vbButtonFace)
        Call .WriteProperty("FillColour", mlngFillColour, vbHighlight)
        Call .WriteProperty("TextColour", mlngTextColour, vbButtonText)
        Call .WriteProperty("OverColour", mlngOverColour, vbHighlightText)
        Call .WriteProperty("BorderStyle", menmBorder, pgrFixed_Single)
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
    End With
End Sub
