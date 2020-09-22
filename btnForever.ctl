VERSION 5.00
Begin VB.UserControl Command 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   DrawWidth       =   2
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   24
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   Begin VB.Timer tmrHover 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   225
      Top             =   2745
   End
End
Attribute VB_Name = "Command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project:ForeverBlue ActiveX OCX Button Control
'Version:1.00
'Author:Dana Seaman - www.cyberactivex.com
'Creation:2002-08-12
'Modified:2003-02-09
'Example:Screenshot of the Visual Basic Demo included in download. All buttons are resizable and don't rely on bitmaps or skins. <BR><BR><IMG src="foreverblue.gif"><BR><BR>

Option Explicit

Private Const PS_SOLID = 0
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_VCENTER = &H4
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800

'String constants
Private Const s_BorderEndColor      As String = "BorderEndColor"
Private Const s_BorderStartColor    As String = "BorderStartColor"
Private Const s_Caption             As String = "Caption"
Private Const s_Enabled             As String = "Enabled"
Private Const s_FocusRectangle      As String = "FocusRectangle"
Private Const s_Font                As String = "Font"
Private Const s_ForeColor           As String = "ForeColor"
Private Const s_GradientEndColor    As String = "GradientEndColor"
Private Const s_GradientStartColor  As String = "GradientStartColor"

' Pen functions:
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Type TRIVERTEX
   x           As Long
   y           As Long
   Red         As Integer
   Green       As Integer
   Blue        As Integer
   Alpha       As Integer
End Type
Private Type RGB
   Red         As Integer
   Green       As Integer
   Blue        As Integer
End Type

Private Type GradientRECT
   UpperLeft   As Long
   LowerRight  As Long  '
End Type

Private Type POINTAPI
   x           As Long
   y           As Long
End Type

' API Declares
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawStateW Lib "user32" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer, ByVal n4 As Integer, ByVal un As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextExW Lib "user32" () As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function ExtTextOutW Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, ByVal lpRect As Long, ByVal lpString As Long, ByVal nCount As Long, ByVal lpDx As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function TextOutW Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type RECT
   left                 As Long
   tOp                  As Long
   Right                As Long
   Bottom               As Long
End Type

Private Enum ButtontState
   sUp
   sDown
   sOver
   sDisable
End Enum

'Default Property Values:
Const m_def_BorderEndColor = vbWhite
Const m_def_BorderStartColor = vbBlack
Const m_def_Caption = "Command"
Const m_def_FocusRectangle = True
Const m_def_GradientEndColor = &HD08746
Const m_def_GradientStartColor = vbWhite

'Property Variables:
Private bHasFocus             As Boolean
Private BorderEndAPI          As RGB
Private BorderStartAPI        As RGB
Private curState              As ButtontState
Private GradientEndAPI        As RGB
Private GradientStartAPI      As RGB
Private He                    As Long
Private isOver                As Boolean
Private m_BorderEndColor      As OLE_COLOR
Private m_BorderStartColor    As OLE_COLOR
Private m_Button              As Integer
Private m_CapRect             As RECT
Private m_Caption             As String
Private m_FocusRectangle      As Boolean
Private m_GradientEndColor    As OLE_COLOR
Private m_GradientStartColor  As OLE_COLOR
Private m_GradRect            As RECT
Private m_MouseX              As Single
Private m_MouseY              As Single
Private m_Shift               As Integer
Private rgnNorm               As Long
Private Wi                    As Long

'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event Click()
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseEnter()
Event MouseExit()

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get BorderEndColor() As OLE_COLOR
   BorderEndColor = m_BorderEndColor
End Property

Public Property Let BorderEndColor(ByVal New_BorderEndColor As OLE_COLOR)
   m_BorderEndColor = New_BorderEndColor
   SetGradientColors
   DrawButton
   PropertyChanged s_BorderEndColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbblack
Public Property Get BorderStartColor() As OLE_COLOR
   BorderStartColor = m_BorderStartColor
End Property

Public Property Let BorderStartColor(ByVal New_BorderStartColor As OLE_COLOR)
   m_BorderStartColor = New_BorderStartColor
   SetGradientColors
   DrawButton
   PropertyChanged s_BorderStartColor
End Property

'MemberInfo=13,0,0,
Public Property Get Caption() As String
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   SetAcccessKey
   DrawButton
   PropertyChanged s_Caption
End Property

Private Sub DrawBorder()
   Dim tR               As RECT
   Dim He23             As Long
   Dim hBrush           As Long

   He23 = He * 2 \ 3

   'Top Line(Penwidth 2)
   DrawLine 0, 1, Wi, 1, m_BorderStartColor, 2
   'Bottom Line(Penwidth 2)
   DrawLine 0, He - 2, Wi, He - 2, m_BorderStartColor, 2
   'Left Top Gradient
   SetRect tR, 0, 0, 3, He23
   DrawRectGradient hdc, tR, BorderStartAPI, BorderEndAPI, 1
   'Right Top Gradient
   OffsetRect tR, Wi - 4, 0
   DrawRectGradient hdc, tR, BorderStartAPI, BorderEndAPI, 1
   'Right Bottom Gradient
   OffsetRect tR, 0, He23 - 4
   tR.Bottom = He - 2
   DrawRectGradient hdc, tR, BorderEndAPI, BorderStartAPI, 1
   'Left Bottom Gradient
   OffsetRect tR, -Wi + 4, 0
   DrawRectGradient hdc, tR, BorderEndAPI, BorderStartAPI, 1
   'Create brush for boxes
   hBrush = CreateSolidBrush(m_BorderStartColor)
   'Left Solid Box(3x3 Pixel)
   SetRect tR, 0, He23 - 3, 3, He23
   FillRect hdc, tR, hBrush
   'Right Solid Box(3x3 Pixel)
   OffsetRect tR, Wi - 4, 0
   FillRect hdc, tR, hBrush
   'Cleanup
   DeleteObject hBrush

End Sub

Private Sub DrawButton()

   On Error Resume Next
   Cls

   If curState = sOver Or curState = sDown Then
      DrawRectGradient hdc, m_GradRect, GradientStartAPI, GradientEndAPI, 1
   ElseIf curState = sDisable Then
      DrawRectGradient hdc, m_GradRect, GradientStartAPI, GradientEndAPI, 1
   ElseIf curState = sUp Then
      DrawRectGradient hdc, m_GradRect, GradientEndAPI, GradientStartAPI, 1
   End If

   DrawBorder
   DrawCaption

End Sub

Private Sub DrawCaption()
   Dim cR               As RECT

   If (bHasFocus) And (m_FocusRectangle) Then
      DrawFocusRect hdc, m_CapRect
   End If
   LSet cR = m_CapRect
   If curState = sOver Then
      OffsetRect cR, 1, 1
   End If
   DrawText hdc, _
      m_Caption, _
      Len(m_Caption), _
      cR, _
      DT_SINGLELINE Or DT_VCENTER Or DT_CENTER

End Sub

Private Sub DrawLine(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Color As Long, Optional ByVal PenWidth As Long = 1, Optional ByVal PenStyle As Long = PS_SOLID)
   'a fast way to draw lines
   Dim pt               As POINTAPI
   Dim oldPen           As Long
   Dim hPen             As Long

   hPen = CreatePen(PenStyle, PenWidth, Color)
   oldPen = SelectObject(hdc, hPen)

   MoveToEx hdc, x1, y1, pt
   LineTo hdc, x2, y2

   SelectObject hdc, oldPen
   DeleteObject hPen

End Sub

Private Sub DrawRectGradient(lHdc As Long, _
   tR As RECT, _
   color1 As RGB, _
   color2 As RGB, _
   Direction As Long)

   Dim V(1)             As TRIVERTEX
   Dim GRct             As GradientRECT

   '# from
   With V(0)
      .x = tR.left
      .y = tR.tOp
      .Red = color1.Red
      .Green = color1.Green
      .Blue = color1.Blue
      .Alpha = 0
   End With
   '# to
   With V(1)
      .x = tR.Right
      .y = tR.Bottom
      .Red = color2.Red
      .Green = color2.Green
      .Blue = color2.Blue
      .Alpha = 0
   End With

   GRct.UpperLeft = 0
   GRct.LowerRight = 1

   GradientFillRect lHdc, V(0), 2, GRct, 1, Direction

End Sub

'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean

   Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

   UserControl.Enabled() = New_Enabled
   If Not New_Enabled Then
      curState = sDisable
      isOver = False
      tmrHover.Enabled = False
   Else
      If isOver Then
         If m_Button = 1 Then
            curState = sDown
         Else
            curState = sUp
         End If
      Else
         curState = sOver
      End If
   End If
   UserControl_Paint

   PropertyChanged s_Enabled

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get FocusRectangle() As Boolean
   FocusRectangle = m_FocusRectangle
End Property

Public Property Let FocusRectangle(ByVal New_FocusRectangle As Boolean)
   m_FocusRectangle = New_FocusRectangle
   DrawButton
   PropertyChanged s_FocusRectangle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
   Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set UserControl.Font = New_Font
   DrawButton
   PropertyChanged s_Font
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
   ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   UserControl.ForeColor() = New_ForeColor
   DrawButton
   PropertyChanged s_ForeColor
End Property


Private Function GetRGBColours(lColour As Long) As RGB

   Dim HexColour        As String

   OleTranslateColor lColour, 0, lColour
   HexColour = String(6 - Len(Hex$(lColour)), "0") & Hex$(lColour)
   GetRGBColours.Red = "&H" & Mid$(HexColour, 5, 2) & "00"
   GetRGBColours.Green = "&H" & Mid$(HexColour, 3, 2) & "00"
   GetRGBColours.Blue = "&H" & Mid$(HexColour, 1, 2) & "00"

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00D08746&
Public Property Get GradientEndColor() As OLE_COLOR
   GradientEndColor = m_GradientEndColor
End Property

Public Property Let GradientEndColor(ByVal New_GradientEndColor As OLE_COLOR)
   m_GradientEndColor = New_GradientEndColor
   SetGradientColors
   DrawButton
   PropertyChanged s_GradientEndColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get GradientStartColor() As OLE_COLOR
   GradientStartColor = m_GradientStartColor
End Property

Public Property Let GradientStartColor(ByVal New_GradientStartColor As OLE_COLOR)
   m_GradientStartColor = New_GradientStartColor
   SetGradientColors
   DrawButton
   PropertyChanged s_GradientStartColor
End Property

Private Function isMouseOver() As Boolean
   Dim pt               As POINTAPI
   GetCursorPos pt
   isMouseOver = (WindowFromPoint(pt.x, pt.y) = hwnd)
End Function

Private Sub MakeRegion()
   'this function creates the region to "cut" the UserControl
   'so it will be transparent in certain areas

   DeleteObject rgnNorm

   rgnNorm = CreateRoundRectRgn(0, 0, Wi, He, 4, 4)

   'Set Usercontrol to new region
   If rgnNorm Then
      SetWindowRgn UserControl.hwnd, rgnNorm, True
   End If

End Sub

Private Sub SetAcccessKey()
   Dim pos              As Integer

   pos = InStr(1, m_Caption, "&")
   If pos Then
      UserControl.AccessKeys = Mid$(m_Caption, pos + 1, 1)
   End If

End Sub

Private Sub SetGradientColors()
   'Special for API Gradients
   GradientStartAPI = GetRGBColours(m_GradientStartColor)
   GradientEndAPI = GetRGBColours(m_GradientEndColor)
   BorderStartAPI = GetRGBColours(m_BorderStartColor)
   BorderEndAPI = GetRGBColours(m_BorderEndColor)

End Sub

Private Sub tmrHover_Timer()
   If Not isMouseOver Then
      tmrHover.Enabled = False
      isOver = False
      curState = sUp
      DrawButton
      RaiseEvent MouseExit
   End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
   curState = sDown
   DrawButton
   UserControl_MouseUp m_Button, m_Shift, m_MouseX, m_MouseY
End Sub

Private Sub UserControl_DblClick()
   'RaiseEvent DblClick
   UserControl_MouseDown m_Button, m_Shift, m_MouseX, m_MouseY
End Sub

Private Sub UserControl_EnterFocus()
   bHasFocus = True
   DrawButton
End Sub

Private Sub UserControl_ExitFocus()
   bHasFocus = False
   DrawButton
End Sub

Private Sub UserControl_InitProperties()

   m_Caption = m_def_Caption
   m_BorderStartColor = m_def_BorderStartColor
   m_BorderEndColor = m_def_BorderEndColor
   m_GradientStartColor = m_def_GradientStartColor
   m_GradientEndColor = m_def_GradientEndColor
   m_FocusRectangle = m_def_FocusRectangle
   Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   m_Button = Button
   If curState <> sDown Then
      curState = sDown
      DrawButton
   End If
   RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, x, y)
   
   m_Button = 0
   m_MouseX = x
   m_MouseY = y
   m_Shift = Shift
   
   curState = sOver
   If Button < 2 Then
      If x < 0 Or y < 0 Or x > Wi Or y > He Then
         'we are outside the button
         curState = sUp
      Else
         If Button = 1 Then curState = sDown
         If isOver = False Then
            isOver = True
            RaiseEvent MouseEnter
            DrawButton
         End If
      End If
   End If

   tmrHover.Enabled = True

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseUp(Button, Shift, x, y)
   RaiseEvent Click

   curState = sUp
   DrawButton

End Sub

Private Sub UserControl_Paint()
   DrawButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_BorderEndColor = PropBag.ReadProperty(s_BorderEndColor, m_def_BorderEndColor)
   m_BorderStartColor = PropBag.ReadProperty(s_BorderStartColor, m_def_BorderStartColor)
   m_Caption = PropBag.ReadProperty(s_Caption, m_def_Caption)
   m_FocusRectangle = PropBag.ReadProperty(s_FocusRectangle, m_def_FocusRectangle)
   m_GradientEndColor = PropBag.ReadProperty(s_GradientEndColor, m_def_GradientEndColor)
   m_GradientStartColor = PropBag.ReadProperty(s_GradientStartColor, m_def_GradientStartColor)
   Set UserControl.Font = PropBag.ReadProperty(s_Font, Ambient.Font)
   UserControl.Enabled = PropBag.ReadProperty(s_Enabled, True)
   UserControl.ForeColor = PropBag.ReadProperty(s_ForeColor, &H80000012)

   SetGradientColors
   SetAcccessKey
   UserControl_Resize

End Sub

Private Sub UserControl_Resize()
   Wi = ScaleWidth
   He = ScaleHeight
   MakeRegion
   SetRect m_GradRect, 3, 2, Wi - 4, He - 3
   LSet m_CapRect = m_GradRect
   InflateRect m_CapRect, -1, -1
   SetAcccessKey
   curState = sUp
   DrawButton

End Sub

Private Sub UserControl_Terminate()
   DeleteObject rgnNorm
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty s_BorderEndColor, m_BorderEndColor, m_def_BorderEndColor
      .WriteProperty s_BorderStartColor, m_BorderStartColor, m_def_BorderStartColor
      .WriteProperty s_Caption, m_Caption, m_def_Caption
      .WriteProperty s_Enabled, UserControl.Enabled, True
      .WriteProperty s_FocusRectangle, m_FocusRectangle, m_def_FocusRectangle
      .WriteProperty s_Font, UserControl.Font, Ambient.Font
      .WriteProperty s_ForeColor, UserControl.ForeColor, &H80000012
      .WriteProperty s_GradientEndColor, m_GradientEndColor, m_def_GradientEndColor
      .WriteProperty s_GradientStartColor, m_GradientStartColor, m_def_GradientStartColor
   End With

End Sub
