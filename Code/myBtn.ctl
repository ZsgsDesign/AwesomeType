VERSION 5.00
Begin VB.UserControl Button 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Í¸Ã÷
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   DefaultCancel   =   -1  'True
   MouseIcon       =   "myBtn.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   470
   ToolboxBitmap   =   "myBtn.ctx":33E2
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawStyle       =   6  'Inside Solid
      Height          =   405
      Left            =   0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   6
      Top             =   0
      Width           =   975
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   540
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer TimerPaint 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   2880
      Top             =   0
   End
   Begin VB.PictureBox picBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   720
      MouseIcon       =   "myBtn.ctx":36F4
      MousePointer    =   99  'Custom
      Picture         =   "myBtn.ctx":6AD6
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   4
      Top             =   990
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   540
      MouseIcon       =   "myBtn.ctx":6D98
      MousePointer    =   99  'Custom
      Picture         =   "myBtn.ctx":A17A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   3
      Top             =   1020
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   360
      MouseIcon       =   "myBtn.ctx":A43C
      MousePointer    =   99  'Custom
      Picture         =   "myBtn.ctx":D81E
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   2
      Top             =   990
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   180
      MouseIcon       =   "myBtn.ctx":DAE0
      MousePointer    =   99  'Custom
      Picture         =   "myBtn.ctx":10EC2
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   1
      Top             =   990
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   990
      MouseIcon       =   "myBtn.ctx":11184
      MousePointer    =   99  'Custom
      Picture         =   "myBtn.ctx":14566
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   0
      Top             =   990
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Button"
      Height          =   180
      Left            =   2760
      TabIndex        =   5
      Top             =   1020
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************
'×÷Õß£ºÕÅÓÓ½Ü
'
'Ãû³Æ£ºmyBtn.ctl
'
'ÃèÊö£ºÍ¨ÓÃ°´Å¥¿Ø¼þµÄ´úÂë
'
'ÍøÕ¾£ºhttps://www.johnzhang.xyz/
'
'ÓÊÏä£ºzsgsdesign@gmail.com
'
'×ñÑ­MITÐ­Òé£¬¶þ´Î¿ª·¢Çë×¢Ã÷Ô­×÷Õß£¡
'****************************************************************************

Option Explicit

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawFocusRect Lib "User32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "User32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, _
        ByVal heightSrc As Long, ByVal lBlendFunction As Long) As Boolean
        
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Enum BlendOp
        AC_SRC_OVER = &H0
        AC_SRC_ALPHA = &H1
End Enum

Private Type BLENDFUNCTION
        BlendOp As Byte
        BlendbtnFlags As Byte
        SourceConstantAlpha As Byte
        AlphaFormat As Byte
End Type

Private Type POINTAPI
        X   As Long
        Y   As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Enum btnStatus
        btnNormal = 0
        btnHot
        btnPressed
        btbDefault
        btbDefault2
        btnNoDraw
End Enum

Private Const DT_CALCRECT = &H400
Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_NOCLIP = &H100
Private Const DT_SINGLELINE = &H20
Private Const DT_INTERNAL = &H1000
Private Const DT_NOPREFIX = &H800
Private Const DT_PLOTTER = 0
Private Const DT_RASDISPLAY = 1
Private Const DT_WORDBREAK = &H10
Private Const DT_EDITCONTROL = &H2000

Public Event Click()
Public Event MouseOut()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Const picW = 6
Private Const picH = 32

Private btnDC(4) As Long
Private btnBMP(4) As Long
Private bTime As Integer
Private oldX As Single, oldY As Single
Private rcCaption As RECT
Private rcFocus As RECT

Private defaultFlag As Boolean
Private btnFlag As btnStatus
Private lastStatus As btnStatus
Private bFocused As Boolean

Private szCaption As String
Private nSpeed As Long
Private bEnabled As Boolean
Private bDefault As Boolean

Private nForeColor As Long, nHotColor As Long, nPressedColor As Long

Dim rComicEnabled As Boolean

Public Property Get ComicEnabled() As Boolean
    ComicEnabled = rComicEnabled
End Property

Public Property Let ComicEnabled(cValue As Boolean)
    rComicEnabled = cValue
    PropertyChanged "ComicEnabled"
End Property

Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        btnFlag = btnPressed
        bTime = 0
        RefreshPicMain
    End If
End Sub

Private Sub picMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        RaiseEvent Click
    End If
End Sub

Private Sub picMain_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Static MouseIn As Boolean
        Static pt As POINTAPI
        
        GetCursorPos pt
        ScreenToClient UserControl.hWnd, pt
        oldX = oldX + 1
        MouseIn = (0 <= pt.X) And (pt.X <= UserControl.ScaleWidth) And (0 <= pt.Y) And (pt.Y <= UserControl.ScaleHeight)
        bTime = 0
        picMain_MouseMove 0, 0, CSng(pt.X), CSng(pt.Y)
    End If
End Sub

Private Sub TimerPaint_Timer()
    
    If btnFlag = btnNoDraw Then
        Exit Sub
    ElseIf btnFlag = btnHot Then
        bTime = bTime + nSpeed
        If picMain.ForeColor <> nHotColor Then picMain.ForeColor = nHotColor
    ElseIf btnFlag = btnPressed Then
        bTime = bTime + nSpeed
        If picMain.ForeColor <> nPressedColor Then picMain.ForeColor = nPressedColor
    ElseIf btnFlag = btnNormal Then
        bTime = bTime + 35
        If picMain.ForeColor <> nForeColor Then
            If bEnabled Then
                picMain.ForeColor = nForeColor
            Else
                picMain.ForeColor = vbGrayText
            End If
        End If
    Else
        If defaultFlag = True Then
            bTime = bTime + 35
        Else
            bTime = bTime + 35
        End If
        If picMain.ForeColor <> nForeColor Then picMain.ForeColor = nForeColor
    End If
    
    If bTime >= 255 Then
        ShowTransparency btnDC(btnFlag), 255
        If btnFlag = btbDefault Then
            btnFlag = btbDefault2
            defaultFlag = False
            bTime = 0
        ElseIf btnFlag = btbDefault2 Then
            btnFlag = btbDefault
            bTime = 0
        Else
            btnFlag = btnNoDraw
        End If
        DrawText picMain.hdc, szCaption, lstrlen(szCaption), rcCaption, DT_CENTER Or DT_EDITCONTROL Or DT_WORDBREAK
        
        picMain.Picture = picMain.Image
        If bFocused Then
            picMain.ForeColor = vbBlack
            DrawFocusRect picMain.hdc, rcFocus
        End If
    Else
        ShowTransparency btnDC(btnFlag), bTime
        DrawText picMain.hdc, szCaption, lstrlen(szCaption), rcCaption, DT_CENTER Or DT_EDITCONTROL Or DT_WORDBREAK
        If bFocused Then
            picMain.ForeColor = vbBlack
            DrawFocusRect picMain.hdc, rcFocus
        End If
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then RaiseEvent Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    bDefault = Ambient.DisplayAsDefault
    RefreshPicMain
    If bDefault Then
        btnFlag = btbDefault
        defaultFlag = True
    Else
        btnFlag = btnNormal
    End If
    bTime = 0
End Sub

Private Sub picMain_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_EnterFocus()
'    bFocused = True                           'ÐéÏß
End Sub

Private Sub UserControl_ExitFocus()
    bFocused = False
End Sub

Private Sub UserControl_InitProperties()
    nSpeed = 20
    szCaption = Replace(UserControl.Extender.Name, "RKShade", "")
    bDefault = Ambient.DisplayAsDefault
    If bDefault Then
        btnFlag = btbDefault
        defaultFlag = True
    End If
    PropertyChanged
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    szCaption = PropBag.ReadProperty("Caption", Replace(UserControl.Extender.Name, "RKShade", ""))
    bEnabled = PropBag.ReadProperty("Enabled", True)
    nSpeed = PropBag.ReadProperty("Speed", "20")
    nForeColor = PropBag.ReadProperty("ForeColor", 0)
    nHotColor = PropBag.ReadProperty("HotColor", 0)
    nPressedColor = PropBag.ReadProperty("PressedColor", 0)
    Set picMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    bDefault = Ambient.DisplayAsDefault
    If bDefault Then
        btnFlag = btbDefault
        defaultFlag = True
    End If
    Enabled = bEnabled
    
    lblCaption.Caption = szCaption
    TimerPaint.Enabled = Ambient.UserMode
    
    If bEnabled Then
        picMain.ForeColor = nForeColor
    Else
        picMain.ForeColor = vbGrayText
    End If
    Static j As Double
    Static i As Long
    
    j = picMain.TextWidth(szCaption) / UserControl.ScaleWidth
    If Int(j) <> j Then
        i = (Int(j) + 1) * picMain.TextHeight("1")
    End If
    
    rcCaption.Right = UserControl.ScaleWidth
    rcCaption.Top = (UserControl.ScaleHeight - i) \ 2
    rcCaption.Bottom = rcCaption.Top + i
    
    Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", szCaption, Replace(UserControl.Extender.Name, "RKShade", "")
    PropBag.WriteProperty "Enabled", bEnabled, True
    PropBag.WriteProperty "Speed", nSpeed, "20"
    PropBag.WriteProperty "ForeColor", nForeColor, 0
    PropBag.WriteProperty "HotColor", nHotColor, 0
    PropBag.WriteProperty "PressedColor", nPressedColor, 0
    Call PropBag.WriteProperty("Font", picMain.Font, Ambient.Font)
End Sub

Public Property Get Caption() As String
    Caption = szCaption
End Property

Public Property Let Caption(szCap As String)
    szCaption = szCap
    PropertyChanged "Caption"
    If bEnabled Then
        picMain.ForeColor = nForeColor
    Else
        picMain.ForeColor = vbGrayText
    End If
    
    Static j As Double
    Static i As Long
    
    j = picMain.TextWidth(szCaption) / UserControl.ScaleWidth
    If Int(j) <> j Then
        i = (Int(j) + 1) * picMain.TextHeight("1")
    End If
    
    rcCaption.Right = UserControl.ScaleWidth
    rcCaption.Top = (UserControl.ScaleHeight - i) \ 2
    rcCaption.Bottom = rcCaption.Top + i
    
    Paint
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(en As Boolean)
    If bDefault Then
        btnFlag = btbDefault
        defaultFlag = True
    Else
        btnFlag = btnNormal
    End If
    bTime = 0
    bEnabled = en
    UserControl.Enabled = en
    picMain.Enabled = en
    PropertyChanged "Enabled"
    
    If bEnabled Then
        picMain.ForeColor = nForeColor
    Else
        picMain.ForeColor = vbGrayText
    End If
    Paint
End Property

Public Property Get Speed() As Long
    Speed = nSpeed
End Property

Public Property Let Speed(nSpd As Long)
    nSpeed = nSpd
    PropertyChanged "Speed"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = nForeColor
End Property

Public Property Let ForeColor(clr As OLE_COLOR)
    nForeColor = clr
    lblCaption.ForeColor = clr
    PropertyChanged "ForeColor"
    If bEnabled Then
        picMain.ForeColor = nForeColor
    Else
        picMain.ForeColor = vbGrayText
    End If
    Paint
End Property
Public Property Get Font() As Font
    Set Font = picMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set picMain.Font = New_Font
    PropertyChanged "Font"
    Paint
End Property
Public Property Get HotColor() As OLE_COLOR
    HotColor = nHotColor
End Property

Public Property Let HotColor(clr As OLE_COLOR)
    nHotColor = clr
    PropertyChanged "HotColor"
End Property

Public Property Get PressedColor() As OLE_COLOR
    PressedColor = nPressedColor
End Property

Public Property Let PressedColor(clr As OLE_COLOR)
    nPressedColor = clr
    PropertyChanged "PressedColor"
End Property

Private Sub UserControl_Initialize()
    Dim i As Byte
    For i = 0 To 4
        btnDC(i) = CreateCompatibleDC(picMain.hdc)
    Next i
    rcFocus.Left = 3
    rcFocus.Top = 3
End Sub

Private Sub Paint()
    Dim i As Long
    Dim W As Long
    Dim H As Long

    W = UserControl.ScaleWidth
    H = UserControl.ScaleHeight
    picMain.Width = W
    picMain.Height = H
    For i = 0 To 4
        With picBtn(i)
            btnBMP(i) = CreateCompatibleBitmap(.hdc, W, H)
            Call SelectObject(btnDC(i), btnBMP(i))
            StretchBlt btnDC(i), 2, 2, W - 4, H - 4, .hdc, 2, 2, picW - 4, picH - 4, vbSrcCopy
            BitBlt btnDC(i), 0, 0, 2, 2, .hdc, 0, 0, vbSrcCopy '×óÉÏ½Ç
            BitBlt btnDC(i), W - 2, 0, 2, 2, .hdc, picW - 2, 0, vbSrcCopy 'ÓÒÉÏ½Ç
            BitBlt btnDC(i), 0, H - 2, 2, 2, .hdc, 0, picH - 2, vbSrcCopy '×óÏÂ½Ç
            BitBlt btnDC(i), W - 2, H - 2, 2, 2, .hdc, picW - 2, picH - 2, vbSrcCopy 'ÓÒÏÂ½Ç
            StretchBlt btnDC(i), 2, 0, W - 4, 2, .hdc, 2, 0, picW - 4, 2, vbSrcCopy         'ÉÏ±ß¿ò
            StretchBlt btnDC(i), 2, H - 2, W - 4, 2, .hdc, 2, picH - 2, picW - 4, 2, vbSrcCopy 'ÏÂ±ß¿ò
            StretchBlt btnDC(i), 0, 2, 2, H - 4, .hdc, 0, 2, 2, picH - 4, vbSrcCopy  '×ó±ß¿ò
            StretchBlt btnDC(i), W - 2, 2, 2, H - 4, .hdc, picW - 2, 2, 2, picH - 4, vbSrcCopy 'ÓÒ±ß¿ò
            If i >= 3 Then
                BitBlt btnDC(i), 2, 2, W - 4, 1, btnDC(i), 2, 1, vbSrcCopy 'ÉÏ±ß¿ò
                BitBlt btnDC(i), 2, H - 3, W - 4, 1, btnDC(i), 2, H - 2, vbSrcCopy 'ÏÂ±ß¿ò
                BitBlt btnDC(i), 2, 2, 1, H - 4, btnDC(i), 1, 2, vbSrcCopy   '×ó±ß¿ò
                BitBlt btnDC(i), W - 3, 2, 1, H - 4, btnDC(i), W - 2, 2, vbSrcCopy  'ÓÒ±ß¿ò
            End If
            If i = 4 Then
                BitBlt btnDC(4), 2, 3, W - 4, 1, btnDC(4), 2, 1, vbSrcCopy       'ÉÏ±ß¿ò
                BitBlt btnDC(4), 2, H - 4, W - 4, 1, btnDC(4), 2, H - 2, vbSrcCopy 'ÏÂ±ß¿ò
                BitBlt btnDC(4), 3, 2, 1, H - 4, btnDC(4), 1, 2, vbSrcCopy    '×ó±ß¿ò
                BitBlt btnDC(4), W - 4, 2, 1, H - 4, btnDC(4), W - 2, 2, vbSrcCopy   'ÓÒ±ß¿ò
            End If
        End With
    Next i

    BitBlt picMain.hdc, 0, 0, W, H, btnDC(0), 0, 0, vbSrcCopy

    Call DrawText(picMain.hdc, szCaption, lstrlen(szCaption), rcCaption, DT_CENTER Or DT_EDITCONTROL Or DT_WORDBREAK)
    
    RefreshPicMain
    For i = 0 To 4
        DeleteObject btnBMP(i)
    Next i
End Sub

Private Sub UserControl_Resize()
    Static j As Double
    Static i As Long
    
    j = picMain.TextWidth(szCaption) / UserControl.ScaleWidth
    If Int(j) <> j Then
        i = (Int(j) + 1) * picMain.TextHeight("1")
    End If
    
    rcCaption.Right = UserControl.ScaleWidth
    rcCaption.Top = (UserControl.ScaleHeight - i) \ 2
    rcCaption.Bottom = rcCaption.Top + i
    
    rcFocus.Right = UserControl.ScaleWidth - 3
    rcFocus.Bottom = UserControl.ScaleHeight - 3
    
    Paint
End Sub

Private Sub UserControl_Terminate()
    Dim i As Byte
    For i = 0 To 2
        DeleteObject btnBMP(i)
        DeleteDC btnDC(i)
    Next i
End Sub

Private Sub RefreshPicMain()
    
    If bFocused Then
        picMain.ForeColor = vbBlack
        DrawFocusRect picMain.hdc, rcFocus
    End If
    picMain.Picture = picMain.Image
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    oldX = X: oldY = Y
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
        bTime = 0
        btnFlag = btnPressed
        lastStatus = btnPressed
        RefreshPicMain
    End If
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If oldX = X And oldY = Y Then Exit Sub

    Static MouseIn As Boolean
    
    oldX = X: oldY = Y

    MouseIn = (0 <= X) And (X <= UserControl.ScaleWidth) And (0 <= Y) And (Y <= UserControl.ScaleHeight)
    If MouseIn Then
        'ReleaseCapture
        RaiseEvent MouseMove(Button, Shift, X, Y)
        SetCapture picMain.hWnd
        
        If Button = 0 Then
            If lastStatus <> btnHot Then
                bTime = 0
                btnFlag = btnHot
                lastStatus = btnHot
            End If
        ElseIf Button = 1 Then
            If lastStatus <> btnPressed Then
                bTime = 0
                btnFlag = btnPressed
                lastStatus = btnPressed
                RefreshPicMain
            End If
        End If
    Else
        If Button = 0 Then ReleaseCapture
        RaiseEvent MouseOut
        If lastStatus = btnHot Or lastStatus = btnPressed Then
            bTime = 0
            If bDefault Then
                btnFlag = btbDefault
                defaultFlag = True
                lastStatus = btbDefault
            Else
                btnFlag = btnNormal
                lastStatus = btnNormal
            End If
            RefreshPicMain
        End If
    End If
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static MouseIn As Boolean
    
    MouseIn = (0 <= X) And (X <= UserControl.ScaleWidth) And (0 <= Y) And (Y <= UserControl.ScaleHeight)
    ReleaseCapture
    RaiseEvent MouseUp(Button, Shift, X, Y)
    bTime = 0
    If MouseIn Then
        btnFlag = btnHot
    Else
        If bDefault Then
            btnFlag = btbDefault
            defaultFlag = True
            lastStatus = btbDefault
        Else
            btnFlag = btnNormal
            lastStatus = btnNormal
        End If
    End If
    RefreshPicMain
    oldX = oldX + 1
End Sub

Private Sub ShowTransparency(cSrc As Long, ByVal nLevel As Byte)
    Dim LrProps As Long
    With picMain
        .Cls
        LrProps = nLevel * &H10000

        AlphaBlend .hdc, 0, 0, .ScaleWidth, .ScaleHeight, _
            cSrc, 0, 0, .ScaleWidth, .ScaleHeight, LrProps
    End With
End Sub





