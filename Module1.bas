Attribute VB_Name = "QuickGDI"
Option Explicit

'Visual Basic Thunder
'www.vbthunder.com

Dim m_hDC As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Declare Function CreateSolidBrush Lib "gdi32" _
     (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "gdi32" _
     (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" _
     (ByVal hObject As Long) As Integer

Declare Function GetSysColor Lib "user32" _
     (ByVal nIndex As ColConst) As Long
'Color constants for GetSysColor
Public Enum ColConst
    COLOR_ACTIVEBORDER = 10
    COLOR_ACTIVECAPTION = 2
    COLOR_ADJ_MAX = 100
    COLOR_ADJ_MIN = -100
    COLOR_APPWORKSPACE = 12
    COLOR_BACKGROUND = 1
    COLOR_BTNFACE = 15
    COLOR_BTNHIGHLIGHT = 20
    COLOR_BTNSHADOW = 16
    COLOR_BTNTEXT = 18
    COLOR_CAPTIONTEXT = 9
    COLOR_GRAYTEXT = 17
    COLOR_HIGHLIGHT = 13
    COLOR_HIGHLIGHTTEXT = 14
    COLOR_INACTIVEBORDER = 11
    COLOR_INACTIVECAPTION = 3
    COLOR_INACTIVECAPTIONTEXT = 19
    COLOR_MENU = 4
    COLOR_MENUTEXT = 7
    COLOR_SCROLLBAR = 0
    COLOR_WINDOW = 5
    COLOR_WINDOWFRAME = 6
    COLOR_WINDOWTEXT = 8
End Enum


Private Declare Function GetTextColor Lib "gdi32" _
     (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" _
     (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" _
     (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
     ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" _
     (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Const NEWTRANSPARENT = 3 'use with SetBkMode()

Private Declare Function CreatePen Lib "gdi32" _
     (ByVal nPenStyle As Long, ByVal nWidth As Long, _
     ByVal crColor As Long) As Long
 Declare Function MoveToEx Lib "gdi32" _
     (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
     lpPoint As POINTAPI) As Long
 Declare Function LineTo Lib "gdi32" _
     (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" _
     (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
     ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Sub DrawRect(ByVal X1 As Long, ByVal Y1 As Long, _
ByVal X2 As Long, ByVal Y2 As Long)
     If m_hDC = 0 Then Exit Sub
     Call Rectangle(m_hDC, X1, Y1, X2, Y2)
End Sub

Public Function GetPen(ByVal nWidth As Long, _
ByVal Clr As Long) As Long
     GetPen = CreatePen(0, nWidth, Clr)
End Function

Public Function hPrint(ByVal X As Long, ByVal Y As Long, _
ByVal hStr As String, ByVal Clr As Long) As Long
     If m_hDC = 0 Then Exit Function
     'Equivalent to setting a form's property
     'FontTransparent = True
     SetBkMode m_hDC, NEWTRANSPARENT

     Dim OT As Long
     OT = GetTextColor(m_hDC)
     SetTextColor m_hDC, Clr
     'Print the text
     hPrint = TextOut(m_hDC, X, Y, hStr, Len(hStr))
     'Restore old text color
     SetTextColor m_hDC, OT
End Function

Public Property Get TargethDC() As Long
     TargethDC = m_hDC
End Property
Public Property Let TargethDC(ByVal vNewValue As Long)
     'The hDC to draw to when performing operations
     'from this module's subroutines.
     m_hDC = vNewValue
End Property

Public Sub ThreedBox(ByVal X1 As Long, ByVal Y1 As Long, _
ByVal X2 As Long, ByVal Y2 As Long, _
Optional Sunken As Boolean = False)
     'Draw a raised box around the specified
     'coordinates.

     If m_hDC = 0 Then Exit Sub

     Dim CurPen As Long, OldPen As Long
     Dim dm As POINTAPI

     If Sunken = False Then
         CurPen = GetPen(1, GetSysColor(COLOR_BTNHIGHLIGHT))
     Else
          CurPen = GetPen(1, GetSysColor(COLOR_BTNSHADOW))
     End If
     OldPen = SelectObject(m_hDC, CurPen)
     'FirstLightLine
     MoveToEx m_hDC, X1, Y2, dm
     LineTo m_hDC, X1, Y1
     'SecondLightLine
     LineTo m_hDC, X2, Y1

     SelectObject m_hDC, OldPen
     DeleteObject CurPen
     If Sunken = False Then
          CurPen = GetPen(1, GetSysColor(COLOR_BTNSHADOW))
     Else
          CurPen = GetPen(1, GetSysColor(COLOR_BTNHIGHLIGHT))
     End If
     OldPen = SelectObject(m_hDC, CurPen)
     'FirstDarkLine
     MoveToEx m_hDC, X2, Y1, dm
     LineTo m_hDC, X2, Y2
     'SecondDarkLine
     LineTo m_hDC, X1, Y2

     SelectObject m_hDC, OldPen
     DeleteObject CurPen
End Sub

'--end block--'
Public Function DrawFilledRect(hdc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Clr As Long)
    Dim hHBR As Long
    Dim r As RECT
    Dim a As Long
    SetRect r, X1, Y1, X2, Y2
    hHBR = CreateSolidBrush(Clr)
    a = CreateRectRgnIndirect(r)
    FillRgn hdc, a, hHBR
    DeleteObject hHBR
End Function

Public Sub DrawFilledRect1(ByVal X1 As Long, ByVal Y1 As Long, _
ByVal X2 As Long, ByVal Y2 As Long)
    Dim hHBR As Long
    Dim hPEN As Long

    hHBR = CreateSolidBrush(RGB(211, 211, 220))
    hPEN = CreatePen(0, 1, RGB(180, 182, 214))
    SelectObject m_hDC, hHBR
    SelectObject m_hDC, hPEN
    Rectangle m_hDC, X1, Y1, X2, Y2
    DeleteObject hHBR
    DeleteObject hPEN
End Sub


