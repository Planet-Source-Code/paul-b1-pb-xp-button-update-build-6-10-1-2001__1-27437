VERSION 5.00
Begin VB.UserControl PBXPButton 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   PropertyPages   =   "PBXPButton2.ctx":0000
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   120
   ToolboxBitmap   =   "PBXPButton2.ctx":003F
End
Attribute VB_Name = "PBXPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'***********************************************
'Active-X Name PB XP Button
'Coded By Paul Beviss pbtools@ntlworld.com
'http://pbtools.port5.com
'***********************************************

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal Clr As Long, ByVal hPal As Long, ByRef lpcolorref As Long)
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" _
   (ByVal hDC As Long, _
   ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal lParam As Long, _
   ByVal wParam As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal cX As Long, _
   ByVal cY As Long, _
   ByVal fuFlags As Long) As Long
Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4

Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_RIGHT = &H8000

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Enum CaptionAlign
        Right
        Center
        Left
End Enum

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10

Const m_def_BorderColor = vbWhite
Const m_def_OverBorderColor = 0
Const m_def_BorderColorDown = 0
Const m_def_BackColor = 0
Const m_def_BackColorOver = 0
Const m_def_BackColorDown = 0
Const m_def_BackColorIcon = &HE0E0E0
Const m_def_BackColorIconOver = vbWhite
Const m_def_BackColorIconDown = &HF1F1F1
Const m_def_IconSizeWidth = 16
Const m_def_IconSizeHeight = 16
Const m_def_Caption = "PB XPButton"
Const m_def_Disabled = False
Const m_def_ShadowSize = 0
Const m_def_FontColor = 0
Const m_def_ShadowColor = 0
Const m_def_ShowShadowOver = False
Const m_def_MultiLine = False
Const m_def_AlignCaption = CaptionAlign.Center
Const m_def_ShowFocus = False
Const m_def_FontColorOver = 0
Const m_def_FontColorDown = 0
Const m_def_Checked = False
Const m_def_Value = False
Const m_def_CheckedColor = &HFFC0C0
'Property Variables:
Dim m_CheckedColor As OLE_COLOR
Dim m_Checked As Boolean
Dim m_Value As Boolean
Dim m_FontColorOver As OLE_COLOR
Dim m_FontColorDown As OLE_COLOR
Dim m_ShowFocus As Boolean
Dim m_AlignCaption As CaptionAlign
Dim m_MultiLine As Boolean
Dim m_ShowShadowOver As Boolean
Dim m_FontColor As OLE_COLOR
Dim m_ShadowColor As OLE_COLOR
Dim m_ShadowSize As Long
Dim m_Disabled As Boolean
Dim m_IconSizeWidth As Long
Dim m_IconSizeHeight As Long
Dim m_BackColorIcon As OLE_COLOR
Dim m_BackColorIconOver As OLE_COLOR
Dim m_BackColorIconDown As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_BackColorOver As OLE_COLOR
Dim m_BackColorDown As OLE_COLOR
Dim m_BorderColorDown As OLE_COLOR
Dim m_BorderColor As OLE_COLOR
Dim m_OverBorderColor As OLE_COLOR
Dim m_Icon As Picture
Dim M_IconOK As Boolean
Dim m_run As Long

Private UsrRect As RECT
Private Ret As Long
Private Clicked As Boolean
Private m_OnButt As Boolean
Private ButtCaption As String
Private InFocus As Boolean

Event Click()
Event MouseOver()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)

'Subclass for double click..
Dim WithEvents mSubClass As SmartSubClass
Attribute mSubClass.VB_VarHelpID = -1

'API declarations:
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204

Private Sub UserControl_Initialize()
    Set mSubClass = New SmartSubClass
    mSubClass.SubClassHwnd UserControl.hWnd, True
End Sub

Private Sub UserControl_Resize()
Call RefreshControl
End Sub

Private Sub UserControl_Terminate()
    mSubClass.SubClassHwnd UserControl.hWnd, False
    Set mSubClass = Nothing
End Sub

Private Sub mSubClass_NewMessage( _
    ByVal hWnd As Long, _
    uMsg As Long, _
    wParam As Long, _
    lParam As Long, _
    Cancel As Boolean)

    Select Case uMsg
        Case WM_LBUTTONDBLCLK
            uMsg = WM_LBUTTONDOWN

        Case WM_RBUTTONDBLCLK
            uMsg = WM_RBUTTONDOWN
    End Select

End Sub

' end subclass

Private Sub UserControl_Click()
If Not m_Disabled Then
    If m_Value Then
    m_Value = False
    Else
    m_Value = True
    End If
    RaiseEvent Click ' Event Click()
End If
End Sub
Public Sub RefreshControl()
If m_Checked And m_Value Then
Call DrawControl(4)
Else
Call DrawControl(0)
End If
End Sub
Public Sub Refresh()
'UserControl.Refresh
End Sub
Public Sub DrawControl(ByVal buttontype As Long)
    Dim Brsh As Long, Clr As Long
    Dim lx As Long, ty As Long
    Dim rx As Long, by As Long
    Dim xx As Long
    Dim sh As Long
    Dim textline As Long
    Dim align As Long
Dim lr As Long
    lx = ScaleLeft: ty = ScaleTop
    rx = ScaleWidth: by = ScaleHeight
    On Error Resume Next
'If m_run = 0 Then
Call SetRect(UsrRect, 0, 0, rx, by)
Select Case m_AlignCaption
Case Center
align = DT_CENTER
Case Left
align = DT_LEFT
Case Right
align = DT_RIGHT
End Select

If m_MultiLine Then
textline = align Or DT_VCENTER
Else
textline = DT_SINGLELINE Or align Or DT_VCENTER
End If
If m_Disabled Then
'If buttontype < 4 Then
buttontype = 0
End If
Cls
Select Case buttontype
'Draw Button Normal
Case 0
    Call SetRect(UsrRect, by, 0, rx, by)
      Call OleTranslateColor(m_BackColor, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FillRect(hDC, UsrRect, Brsh)
      DeleteObject Brsh
    Call SetRectEmpty(UsrRect)
    Call SetRect(UsrRect, 0, 0, by, by)
      Call OleTranslateColor(m_BackColorIcon, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FillRect(hDC, UsrRect, Brsh)
      DeleteObject Brsh
    Call SetRectEmpty(UsrRect)
    Call SetRect(UsrRect, 0, 0, rx, by)
      Call OleTranslateColor(m_BorderColor, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FrameRect(hDC, UsrRect, Brsh)
      DeleteObject Brsh
    Call SetRectEmpty(UsrRect)
    If m_Disabled Then
    lr = DrawState(hDC, 0, 0, m_Icon, 0, (by / 2) - (m_IconSizeWidth / 2), (by / 2) - (m_IconSizeHeight / 2), m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_DISABLED)
    Else
       lr = DrawState(hDC, 0, 0, m_Icon, 0, (by / 2) - (m_IconSizeWidth / 2), (by / 2) - (m_IconSizeHeight / 2) - 1, m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_NORMAL)
 End If
'Draw Button Over
Case 1

Call SetRect(UsrRect, by, 0, rx, by)
    Call OleTranslateColor(m_BackColorOver, ByVal 0&, Clr)
    Brsh = CreateSolidBrush(Clr)
    Call FillRect(hDC, UsrRect, Brsh)
    DeleteObject Brsh
Call SetRectEmpty(UsrRect)
 Call SetRect(UsrRect, 0, 0, by, by)
    Call OleTranslateColor(m_BackColorIconOver, ByVal 0&, Clr)
    Brsh = CreateSolidBrush(Clr)
    Call FillRect(hDC, UsrRect, Brsh)
    DeleteObject Brsh
Call SetRectEmpty(UsrRect)
If m_Checked And m_Value Or m_Disabled Then
  Call SetRect(UsrRect, 1, 1, by - 1, by - 1)
    Call OleTranslateColor(m_BorderColorDown, ByVal 0&, Clr)
    Brsh = CreateSolidBrush(Clr)
    Call FrameRect(hDC, UsrRect, Brsh)
    DeleteObject Brsh
 Call SetRectEmpty(UsrRect)
End If
    Call SetRect(UsrRect, 0, 0, rx, by)
        Call OleTranslateColor(m_OverBorderColor, ByVal 0&, Clr)
        Brsh = CreateSolidBrush(Clr)
        Call FrameRect(hDC, UsrRect, Brsh)
        DeleteObject Brsh
Call SetRectEmpty(UsrRect)
 Brsh = CreateSolidBrush(RGB(136, 141, 157))
    lr = DrawState(hDC, Brsh, 0, m_Icon, 0, (by / 2) - (m_IconSizeWidth / 2) + 1, (by / 2) - (m_IconSizeHeight / 2) + 1, m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_MONO)
    DeleteObject Brsh
    lr = DrawState(hDC, 0, 0, m_Icon, 0, (by / 2) - (m_IconSizeWidth / 2) - 1, (by / 2) - (m_IconSizeHeight / 2) - 1, m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_NORMAL)

'Draw Button Down
Case 2
    Call SetRect(UsrRect, by, 0, rx, by)
      Call OleTranslateColor(m_BackColorDown, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FillRect(hDC, UsrRect, Brsh)
      DeleteObject Brsh
    Call SetRectEmpty(UsrRect)
    Call SetRect(UsrRect, 0, 0, by, by)
      Call OleTranslateColor(m_BackColorIconDown, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FillRect(hDC, UsrRect, Brsh)
      DeleteObject Brsh
    Call SetRectEmpty(UsrRect)
    If m_Checked And m_Value Then
    Call SetRect(UsrRect, 1, 1, by, by)
      Call OleTranslateColor(m_BorderColorDown, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FrameRect(hDC, UsrRect, Brsh)
      DeleteObject Brsh
    Call SetRectEmpty(UsrRect)
    End If
    Call SetRect(UsrRect, 0, 0, rx, by)
      Call OleTranslateColor(m_BorderColorDown, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FrameRect(hDC, UsrRect, Brsh)
      DeleteObject Brsh
      Call SetRectEmpty(UsrRect)
 lr = DrawState(hDC, 0, 0, m_Icon, 0, (by / 2) - (m_IconSizeWidth / 2), (by / 2) - (m_IconSizeHeight / 2) - 1, m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_NORMAL)
    Case 4
'Draw Button Checked
    Call SetRect(UsrRect, by, 0, rx, by)
      Call OleTranslateColor(m_BackColor, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FillRect(hDC, UsrRect, Brsh)
      DeleteObject Brsh
      Call SetRectEmpty(UsrRect)
    Call SetRect(UsrRect, 0, 0, by, by)
      Call OleTranslateColor(m_CheckedColor, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FillRect(hDC, UsrRect, Brsh)
      DeleteObject Brsh
      Call SetRectEmpty(UsrRect)
    Call SetRect(UsrRect, 0, 0, rx, by)
      Call OleTranslateColor(m_BorderColor, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FrameRect(hDC, UsrRect, Brsh)
      DeleteObject Brsh
     Call SetRectEmpty(UsrRect)
    Call SetRect(UsrRect, 1, 1, by - 1, by - 1)
      Call OleTranslateColor(m_BorderColorDown, ByVal 0&, Clr)
      Brsh = CreateSolidBrush(Clr)
      Call FrameRect(hDC, UsrRect, Brsh)
      DeleteObject Brsh
    Call SetRectEmpty(UsrRect)

 lr = DrawState(hDC, 0, 0, m_Icon, 0, (by / 2) - (m_IconSizeWidth / 2), (by / 2) - (m_IconSizeHeight / 2) - 1, m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_NORMAL)


End Select

'Draw Caption Shadow
If ButtCaption > "" Then
    If m_Disabled Then
    sh = m_ShadowSize
    m_ShadowSize = 1
    End If
        If m_ShadowSize > 0 Or m_ShadowSize < 10 Then
        If m_Disabled Then
        ForeColor = vbWhite
        Else
        ForeColor = m_ShadowColor
        End If
            For xx = 1 To m_ShadowSize
            Call SetRect(UsrRect, by + xx + 2, xx + 2, rx + xx - 2, by + xx - 2)
            If m_ShowShadowOver And Not m_Disabled Then
            If buttontype = 1 Then
            Call DrawText(hDC, ButtCaption, -1, UsrRect, textline)
            End If
            Else
            Call DrawText(hDC, ButtCaption, -1, UsrRect, textline)
            End If
            Next xx
        End If
Call SetRectEmpty(UsrRect)
'Draw Caption
If m_Disabled Then
ForeColor = vb3DShadow
m_ShadowSize = sh
Else
Select Case buttontype
Case 0
ForeColor = m_FontColor
Case 1
ForeColor = m_FontColorOver
Case 2
ForeColor = m_FontColorDown
Case Else
ForeColor = m_FontColor
End Select
End If
Call SetRect(UsrRect, by + 2, 2, rx - 2, by - 2)
Call DrawText(hDC, ButtCaption, -1, UsrRect, textline)
   Call SetRectEmpty(UsrRect)
    If InFocus And m_ShowFocus And Not m_Disabled Then
     Call SetRect(UsrRect, 2, 2, rx - 2, by - 2)
     Call DrawFocusRect(hDC, UsrRect)
        Call SetRectEmpty(UsrRect)
    End If
End If
'End If
End Sub
Public Sub About()
MsgBox "PB XP Button" & vbCrLf & "pbtools@ntlworld.com"
End Sub
Private Sub UserControl_DblClick()
If Not m_Disabled Then
RaiseEvent Click ' Event Click()
End If
End Sub

Private Sub UserControl_GotFocus()
InFocus = True
End Sub

Private Sub UserControl_InitProperties()
    m_BorderColor = m_def_BorderColor
    m_OverBorderColor = RGB(10, 36, 106)
    Set m_Icon = LoadPicture("")
    m_BorderColorDown = RGB(10, 36, 106)
    m_BackColor = RGB(219, 216, 209)
    m_BackColorOver = RGB(182, 189, 210)
    m_BackColorDown = RGB(133, 146, 181)
    m_BackColorIcon = m_def_BackColorIcon
    m_BackColorIconOver = m_def_BackColorIconOver
    m_BackColorIconDown = m_def_BackColorIconDown
    m_IconSizeWidth = m_def_IconSizeWidth
    m_IconSizeHeight = m_def_IconSizeHeight
    ButtCaption = m_def_Caption
    Set UserControl.Font = Ambient.Font
    m_Disabled = m_def_Disabled
    m_ShadowSize = m_def_ShadowSize
    m_FontColor = m_def_FontColor
    m_ShadowColor = m_def_ShadowColor
    m_ShowShadowOver = m_def_ShowShadowOver
    m_MultiLine = m_def_MultiLine
    m_AlignCaption = m_def_AlignCaption
    m_ShowFocus = m_def_ShowFocus
    m_FontColorOver = m_def_FontColorOver
    m_FontColorDown = m_def_FontColorDown
    m_Checked = m_def_Checked
    m_Value = m_def_Value
     m_CheckedColor = RGB(213, 215, 216)
    Call RefreshControl
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Not m_Disabled Then
Call DrawControl(1)
RaiseEvent KeyDown(KeyCode, Shift)
End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
If Not m_Disabled Then
RaiseEvent KeyPress(KeyAscii)
End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Not m_Disabled Then
    RaiseEvent Click ' Event Click()
    Call DrawControl(0)
    RaiseEvent KeyDown(KeyCode, Shift)
End If
End Sub

Private Sub UserControl_LostFocus()
InFocus = False
Call RefreshControl
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y) ' Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clicked = True
    m_run = 0
     Call DrawControl(2)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y) ' MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Clicked Then
        'm_OnButt = False
        UserControl_MouseOut Button, Shift, X, Y
    End If
End Sub
Private Sub UserControl_MouseOver(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseOver
End Sub
Function UserControl_MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_OnButt = False
    
    If X <= 0 Or X >= ScaleWidth Or Y <= 0 Or Y >= ScaleHeight Then
        m_OnButt = False
        Ret = ReleaseCapture()
        m_run = 0
       Call RefreshControl
        RaiseEvent MouseOut(Button, Shift, X, Y)
    Else
        If m_OnButt = False Then
            m_OnButt = True
            Ret = SetCapture(UserControl.hWnd)
        
        Call DrawControl(1)
        m_run = 1
        RaiseEvent MouseOver
        End If
    End If

End Function

Public Property Get Icon() As Picture
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set m_Icon = New_Icon
    PropertyChanged "Icon"
    If New_Icon > 0 Then
    M_IconOK = True
    Else
    M_IconOK = False
    End If
      Call RefreshControl
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
   Call RefreshControl
End Property
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    m_run = 0
    Clicked = False
    Call DrawControl(0)
    UserControl_MouseOut Button, Shift, X, Y
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "Options"
    Caption = ButtCaption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    ButtCaption = New_Caption
    PropertyChanged "Caption"
   Call RefreshControl
    
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ButtCaption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_OverBorderColor = PropBag.ReadProperty("BorderColorOver", m_def_OverBorderColor)
    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
    m_BorderColorDown = PropBag.ReadProperty("BorderColorDown", m_def_BorderColorDown)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_BackColorOver = PropBag.ReadProperty("BackColorOver", m_def_BackColorOver)
    m_BackColorDown = PropBag.ReadProperty("BackColorDown", m_def_BackColorDown)
    m_BackColorIcon = PropBag.ReadProperty("BackColorIcon", m_def_BackColorIcon)
    m_BackColorIconOver = PropBag.ReadProperty("BackColorIconOver", m_def_BackColorIconOver)
    m_BackColorIconDown = PropBag.ReadProperty("BackColorIconDown", m_def_BackColorIconDown)
    m_IconSizeWidth = PropBag.ReadProperty("IconSizeWidth", m_def_IconSizeWidth)
    m_IconSizeHeight = PropBag.ReadProperty("IconSizeHeight", m_def_IconSizeHeight)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Disabled = PropBag.ReadProperty("Disabled", m_def_Disabled)
    m_ShadowSize = PropBag.ReadProperty("ShadowSize", m_def_ShadowSize)
    m_FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
    m_ShadowColor = PropBag.ReadProperty("ShadowColor", m_def_ShadowColor)
    m_ShowShadowOver = PropBag.ReadProperty("ShowShadowOver", m_def_ShowShadowOver)
    m_MultiLine = PropBag.ReadProperty("MultiLine", m_def_MultiLine)
    m_AlignCaption = PropBag.ReadProperty("AlignCaption", m_def_AlignCaption)
    m_ShowFocus = PropBag.ReadProperty("ShowFocus", m_def_ShowFocus)
    m_FontColorOver = PropBag.ReadProperty("FontColorOver", m_def_FontColorOver)
    m_FontColorDown = PropBag.ReadProperty("FontColorDown", m_def_FontColorDown)
    m_Checked = PropBag.ReadProperty("Checked", m_def_Checked)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)

    m_CheckedColor = PropBag.ReadProperty("CheckedColor", m_def_CheckedColor)

   Call RefreshControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", ButtCaption, m_def_Caption)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BorderColorOver", m_OverBorderColor, m_def_OverBorderColor)
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("BorderColorDown", m_BorderColorDown, m_def_BorderColorDown)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("BackColorOver", m_BackColorOver, m_def_BackColorOver)
    Call PropBag.WriteProperty("BackColorDown", m_BackColorDown, m_def_BackColorDown)
    Call PropBag.WriteProperty("BackColorIcon", m_BackColorIcon, m_def_BackColorIcon)
    Call PropBag.WriteProperty("BackColorIconOver", m_BackColorIconOver, m_def_BackColorIconOver)
    Call PropBag.WriteProperty("BackColorIconDown", m_BackColorIconDown, m_def_BackColorIconDown)
    Call PropBag.WriteProperty("IconSizeWidth", m_IconSizeWidth, m_def_IconSizeWidth)
    Call PropBag.WriteProperty("IconSizeHeight", m_IconSizeHeight, m_def_IconSizeHeight)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Disabled", m_Disabled, m_def_Disabled)
    Call PropBag.WriteProperty("ShadowSize", m_ShadowSize, m_def_ShadowSize)
    Call PropBag.WriteProperty("FontColor", m_FontColor, m_def_FontColor)
    Call PropBag.WriteProperty("ShadowColor", m_ShadowColor, m_def_ShadowColor)
    Call PropBag.WriteProperty("ShowShadowOver", m_ShowShadowOver, m_def_ShowShadowOver)
    Call PropBag.WriteProperty("MultiLine", m_MultiLine, m_def_MultiLine)
    Call PropBag.WriteProperty("AlignCaption", m_AlignCaption, m_def_AlignCaption)
    Call PropBag.WriteProperty("ShowFocus", m_ShowFocus, m_def_ShowFocus)
    Call PropBag.WriteProperty("FontColorOver", m_FontColorOver, m_def_FontColorOver)
    Call PropBag.WriteProperty("FontColorDown", m_FontColorDown, m_def_FontColorDown)
    Call PropBag.WriteProperty("Checked", m_Checked, m_def_Checked)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("CheckedColor", m_CheckedColor, m_def_CheckedColor)

    End Sub

Public Property Get BorderColorOver() As OLE_COLOR
    BorderColorOver = m_OverBorderColor
End Property

Public Property Let BorderColorOver(ByVal New_OverBorderColor As OLE_COLOR)
    m_OverBorderColor = New_OverBorderColor
    PropertyChanged "BorderColorOver"
    Call RefreshControl
End Property
Public Property Get BorderColorDown() As OLE_COLOR
    BorderColorDown = m_BorderColorDown
End Property

Public Property Let BorderColorDown(ByVal New_BorderColorDown As OLE_COLOR)
    m_BorderColorDown = New_BorderColorDown
    PropertyChanged "BorderColorDown"
    Call RefreshControl
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BackColorOver() As OLE_COLOR
    BackColorOver = m_BackColorOver
End Property

Public Property Let BackColorOver(ByVal New_BackColorOver As OLE_COLOR)
    m_BackColorOver = New_BackColorOver
    PropertyChanged "BackColorOver"
    Call RefreshControl
End Property

Public Property Get BackColorDown() As OLE_COLOR
    BackColorDown = m_BackColorDown
End Property

Public Property Let BackColorDown(ByVal New_BackColorDown As OLE_COLOR)
    m_BackColorDown = New_BackColorDown
    PropertyChanged "BackColorDown"
    Call RefreshControl
End Property
Public Property Get BackColorIcon() As OLE_COLOR
    BackColorIcon = m_BackColorIcon
End Property

Public Property Let BackColorIcon(ByVal New_BackColorIcon As OLE_COLOR)
    m_BackColorIcon = New_BackColorIcon
    PropertyChanged "BackColorIcon"
    Call RefreshControl
End Property


Public Property Get BackColorIconOver() As OLE_COLOR
    BackColorIconOver = m_BackColorIconOver
End Property

Public Property Let BackColorIconOver(ByVal New_BackColorIconOver As OLE_COLOR)
    m_BackColorIconOver = New_BackColorIconOver
    PropertyChanged "BackColorIconOver"
    Call RefreshControl
End Property

Public Property Get BackColorIconDown() As OLE_COLOR
    BackColorIconDown = m_BackColorIconDown
End Property
Public Property Let BackColorIconDown(ByVal New_BackColorIconDown As OLE_COLOR)
    m_BackColorIconDown = New_BackColorIconDown
    PropertyChanged "BackColorIconDown"
    Call RefreshControl
End Property
Public Property Get IconSizeWidth() As Long
Attribute IconSizeWidth.VB_ProcData.VB_Invoke_Property = "Options"
    IconSizeWidth = m_IconSizeWidth
End Property

Public Property Let IconSizeWidth(ByVal New_IconSizeWidth As Long)
    m_IconSizeWidth = New_IconSizeWidth
    PropertyChanged "IconSizeWidth"
   Call RefreshControl
End Property

Public Property Get IconSizeHeight() As Long
Attribute IconSizeHeight.VB_ProcData.VB_Invoke_Property = "Options"
    IconSizeHeight = m_IconSizeHeight
End Property

Public Property Let IconSizeHeight(ByVal New_IconSizeHeight As Long)
    m_IconSizeHeight = New_IconSizeHeight
    PropertyChanged "IconSizeHeight"
    Call RefreshControl
End Property
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Call RefreshControl
End Property
Public Property Get Disabled() As Boolean
Attribute Disabled.VB_ProcData.VB_Invoke_Property = "Options"
    Disabled = m_Disabled
End Property

Public Property Let Disabled(ByVal New_Disabled As Boolean)
    m_Disabled = New_Disabled
    PropertyChanged "Disabled"
   Call RefreshControl
End Property
Public Property Get ShadowSize() As Long
Attribute ShadowSize.VB_ProcData.VB_Invoke_Property = "Options"
    ShadowSize = m_ShadowSize
End Property

Public Property Let ShadowSize(ByVal New_ShadowSize As Long)
    m_ShadowSize = New_ShadowSize
    PropertyChanged "ShadowSize"
   Call RefreshControl
End Property
Public Property Get FontColor() As OLE_COLOR
    FontColor = m_FontColor
End Property

Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
    m_FontColor = New_FontColor
    PropertyChanged "FontColor"
   Call RefreshControl
End Property
Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(ByVal New_ShadowColor As OLE_COLOR)
    m_ShadowColor = New_ShadowColor
    PropertyChanged "ShadowColor"
   Call RefreshControl
End Property
Public Property Get ShowShadowOver() As Boolean
Attribute ShowShadowOver.VB_ProcData.VB_Invoke_Property = "Options"
    ShowShadowOver = m_ShowShadowOver
End Property

Public Property Let ShowShadowOver(ByVal New_ShowShadowOver As Boolean)
    m_ShowShadowOver = New_ShowShadowOver
    PropertyChanged "ShowShadowOver"
 Call DrawControl(0)
End Property

Public Property Get MultiLine() As Boolean
    MultiLine = m_MultiLine
End Property

Public Property Let MultiLine(ByVal New_MultiLine As Boolean)
    m_MultiLine = New_MultiLine
    PropertyChanged "MultiLine"
End Property
Public Property Get AlignCaption() As CaptionAlign
    AlignCaption = m_AlignCaption
End Property

Public Property Let AlignCaption(ByVal New_AlignCaption As CaptionAlign)
    m_AlignCaption = New_AlignCaption
    PropertyChanged "AlignCaption"
     Call RefreshControl
End Property

Public Property Get ShowFocus() As Boolean
    ShowFocus = m_ShowFocus
End Property

Public Property Let ShowFocus(ByVal New_ShowFocus As Boolean)
    m_ShowFocus = New_ShowFocus
    PropertyChanged "ShowFocus"
  Call RefreshControl
End Property

Public Property Get FontColorOver() As OLE_COLOR
    FontColorOver = m_FontColorOver
End Property

Public Property Let FontColorOver(ByVal New_FontColorOver As OLE_COLOR)
    m_FontColorOver = New_FontColorOver
    PropertyChanged "FontColorOver"
 Call RefreshControl
End Property

Public Property Get FontColorDown() As OLE_COLOR
    FontColorDown = m_FontColorDown
End Property

Public Property Let FontColorDown(ByVal New_FontColorDown As OLE_COLOR)
    m_FontColorDown = New_FontColorDown
    PropertyChanged "FontColorDown"
Call RefreshControl
End Property
Public Property Get Checked() As Boolean
    Checked = m_Checked
End Property

Public Property Let Checked(ByVal New_Checked As Boolean)
    m_Checked = New_Checked
    PropertyChanged "Checked"
Call RefreshControl
End Property

Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    PropertyChanged "Value"
Call RefreshControl
End Property

Public Property Get CheckedColor() As OLE_COLOR
    CheckedColor = m_CheckedColor
End Property

Public Property Let CheckedColor(ByVal New_CheckedColor As OLE_COLOR)
    m_CheckedColor = New_CheckedColor
    PropertyChanged "CheckedColor"
Call RefreshControl
End Property
