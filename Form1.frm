VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   690
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   690
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      Height          =   375
      Left            =   3000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   7
      Left            =   2640
      Picture         =   "Form1.frx":0000
      Top             =   120
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   6
      Left            =   2280
      Picture         =   "Form1.frx":0102
      Top             =   120
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   5
      Left            =   1920
      Picture         =   "Form1.frx":0204
      Top             =   120
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   4
      Left            =   1560
      Top             =   120
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   3
      Left            =   1200
      Picture         =   "Form1.frx":0306
      Top             =   120
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   2
      Left            =   840
      Picture         =   "Form1.frx":0408
      Top             =   120
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   1
      Left            =   480
      Top             =   120
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu hdfn 
         Caption         =   "kio"
      End
      Begin VB.Menu dgfndfh 
         Caption         =   "-"
      End
      Begin VB.Menu ebrbff 
         Caption         =   "erre"
      End
      Begin VB.Menu efrsdfh 
         Caption         =   "wew"
      End
      Begin VB.Menu sfnsdfh 
         Caption         =   "w3etwernh"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Visual Basic Thunder
'www.vbthunder.com

Dim pnt As PaintEffects

Dim MyFont As Long
Dim OldFont As Long

Dim wlOldProc As Long

Dim Caps(2 To 19) As String
Private Declare Sub CopyMem Lib "kernel32" Alias _
    "RtlMoveMemory" (pDest As Any, pSource As Any, _
    ByVal ByteLen As Long)
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long




Private Sub Form_Load()
    
    Set pnt = New PaintEffects
    'Sets the captions
    Caps(2) = "New"
    Caps(3) = "Copy"
    Caps(4) = "Paste"
    Caps(5) = "Cut"
    Caps(6) = "Open"
    Caps(7) = "Save"
    
    If wlOldProc <> 0 Then Exit Sub
    
    Dim i As Integer
    
    MenuItems.MenuForm = Me
    'Start with File menu
    MenuItems.SubMenu = 0
    
    For i = 0 To 6
        MenuItems.MenuID = i
        OwnerDrawMenu (i + 2)
    Next
    
    
    wlOldProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf OwnMenuProc)

End Sub


Public Function IsSeparator(ByVal IID As Integer) As Boolean
    Dim mii As MENUITEMINFO
    mii.cbSize = Len(mii)
    mii.fMask = MIIM_TYPE
    mii.wID = IID
    GetMenuItemInfo GetMenu(hWnd), IID, False, mii
    IsSeparator = ((mii.fType And MFT_SEPARATOR) = MFT_SEPARATOR)
End Function



Public Function MsgProc(ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
    'This procedure is called because we've subclassed
    'this form. We will catch DRAWITEM and MEASUREITEM
    'messages and pass the rest of them on.
    
    'Various structs we'll need
    Dim MeasureInfo As MEASUREITEMSTRUCT
    Dim DrawInfo As DRAWITEMSTRUCT
    'Set later for separator flag:
    Dim IsSep As Boolean
    'Our custom brush and the old one
    Dim hBr As Long, hOldBr As Long
    'Our custom pen and the old one
    Dim hPEN As Long, hOldPen As Long
    'The text color of the menu items
    Dim lTextColor As Long
    'Now much to bump the menu's selection
    'rectangle over
    Dim iRectOffset As Integer
    
    If wMsg = WM_DRAWITEM Then
        If wParam = 0 Then 'It was sent by the menu
            'Get DRAWINFOSTRUCT -- copy it to our
            'empty structure from the pointer in lParam
            Call CopyMem(DrawInfo, ByVal lParam, LenB(DrawInfo))
            IsSep = IsSeparator(DrawInfo.itemID)
            
            '===Set the menu font through its hDC...===
            MyFont = SendMessage(Me.hWnd, WM_GETFONT, 0&, 0&)
            OldFont = SelectObject(DrawInfo.hdc, MyFont)
            'We draw the item based on Un/Selected:
            If (DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED Then
                hBr = CreateSolidBrush( _
                GetSysColor(COLOR_HIGHLIGHT))
                'hPEN = GetPen(1, GetSysColor(COLOR_HIGHLIGHT))
                lTextColor = GetSysColor(COLOR_MENUTEXT)
            Else
                hBr = CreateSolidBrush(GetSysColor(COLOR_MENU))
                hPEN = GetPen(1, GetSysColor(COLOR_MENU))
                lTextColor = GetSysColor(COLOR_MENUTEXT)
            End If
            'We're going to draw on the menu
            QuickGDI.TargethDC = DrawInfo.hdc
            'Select our new, correctly colored objects:
            

            hOldBr = SelectObject(DrawInfo.hdc, hBr)
            hOldPen = SelectObject(DrawInfo.hdc, hPEN)
            With DrawInfo.rcItem
                'not selected
                If (DrawInfo.itemState And ODS_SELECTED) <> ODS_SELECTED Then
                    'Clear the space where the image is
                    QuickGDI.DrawRect .Left, .Top, _
                        22, .Bottom
                End If
                
                'Check to see if the menu item is one of the
                'ones with a picture. If so, then we need to
                'move the edge of the drawing rectangle a little
                'to the left to make room for the image.
                iRectOffset = IIf(img(DrawInfo.itemID).Picture.Handle <> 0 _
                    , 23, 0)
                'Do we have a separator bar?
                If Not IsSep Then
                    'Draw the rectangle onto the item's space
                    
                    
                    QuickGDI.DrawRect .Left + iRectOffset, _
                        .Top, .Right, .Bottom
                    
                    DrawFilledRect DrawInfo.hdc, .Left, .Top, .Right, .Bottom, vbWhite
                    
                    If (DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED Then
                        DrawFilledRect1 .Left, _
                        .Top, .Right, .Bottom
                    End If
                    
                    'Print the item's text
                    '(held in the Caps() array)
                    hPrint .Left + 30, .Top + 3, _
                        Caps(DrawInfo.itemID), _
                        lTextColor
                End If
            End With
            'Select the old objects into the menu's DC
            Call SelectObject(DrawInfo.hdc, hOldBr)
            Call SelectObject(DrawInfo.hdc, hOldPen)
            'Delete the ones we created
            Call DeleteObject(hBr)
            Call DeleteObject(hPEN)
            With DrawInfo
                'If the item had an image:
                '2 = New, 3 = Open, 4 = Save, etc.
                If img(DrawInfo.itemID).Picture.Handle <> 0 Then
                    
                    If (DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED Then
                        Dim i As Long
                        Dim e As Long
                        Dim a As Long
                        
                        Picture2.Cls
                        
                        'draws the bitmap to picture2
                        pnt.PaintTransparentStdPic Picture2.hdc, _
                            0, 0, _
                            16, 16, img(DrawInfo.itemID).Picture, _
                            0, 0, vbMagenta
                        
                        'makes the shadow for the icon.
                        For i = 0 To 16
                            For e = 0 To 16
                                a = GetPixel(Picture2.hdc, i, e)
                                If a <> vbMagenta Then SetPixel Picture2.hdc, i, e, RGB(158, 158, 165)
                            Next
                        Next
                    
                        Picture2.Refresh
                        
                        'draws icon shadow
                        pnt.PaintTransparentDC .hdc, _
                            5, .rcItem.Top + 6, _
                            16, 16, Picture2.hdc, _
                            0, 0, vbMagenta
                                                    
                        'draws icon
                        pnt.PaintTransparentStdPic .hdc, _
                            3, .rcItem.Top + 4, _
                            16, 16, img(DrawInfo.itemID).Picture, _
                            0, 0, vbMagenta
                    Else
                        'draw the light grey bar on the left
                        DrawFilledRect .hdc, 0, .rcItem.Top, 23, .rcItem.Bottom, RGB(241, 240, 242)
                        
                        'draws the icon
                        pnt.PaintTransparentStdPic .hdc, _
                            4, .rcItem.Top + 5, _
                            16, 16, img(DrawInfo.itemID).Picture, _
                            0, 0, vbMagenta
                    End If
                    
                End If
                If IsSep Then
                    'Draw the special separator bar
                    'ThreedBox .rcItem.Left, _
                        .rcItem.Top + 2, _
                        .rcItem.Right - 1, _
                        .rcItem.Bottom - 2, True
                    Dim pt As POINTAPI
                    DrawFilledRect .hdc, .rcItem.Left, .rcItem.Top, .rcItem.Right, .rcItem.Bottom, vbWhite
                    DrawFilledRect .hdc, 0, .rcItem.Top, 23, .rcItem.Bottom, RGB(241, 240, 242)
                    MoveToEx .hdc, .rcItem.Left + 25, .rcItem.Top + 2, pt
                    LineTo .hdc, .rcItem.Right, .rcItem.Top + 2
                End If
            End With
        End If
        'Don't pass this message on:
        MsgProc = False
        Exit Function
        
    ElseIf wMsg = WM_MEASUREITEM Then
        'Get the MEASUREITEM struct from the pointer
        Call CopyMem(MeasureInfo, ByVal lParam, Len(MeasureInfo))
        IsSep = IsSeparator(MeasureInfo.itemID)
        'Tell Windows how big our items are.
        MeasureInfo.itemWidth = 150
        'If the item being measured is the separator
        'bar, the height should be 5 pixels, 18 if
        'otherwise...
        MeasureInfo.itemHeight = IIf(IsSep, 5, 22)
        'Return the information back to Windows
        Call CopyMem(ByVal lParam, MeasureInfo, Len(MeasureInfo))
        'Don't pass this message on:
        MsgProc = False
        Exit Function
    ElseIf wMsg = WM_MENUSELECT Then
        'lblNum.Caption = LoWord(wParam) & ", (" & HiWord(wParam) & ")"
        
    End If
    
    'We didn't handle this message,
    'pass it on to the next WndProc
    MsgProc = CallWindowProc(wlOldProc, hWnd, wMsg, wParam, lParam)
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    If wlOldProc <> 0 Then
        SetWindowLong hWnd, GWL_WNDPROC, wlOldProc
    End If
    Set pnt = Nothing
    
    'Destroy the font object created in
    'the form's window procedure.
    Call DeleteObject(MyFont)
    
End Sub
'--end block--'
