VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FrmMain 
   Caption         =   "AwesomeType"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12315
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   536
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   821
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox PicRMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   4980
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   139
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   2115
      Begin AwesomeType.Button MnSelAll 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   -15
         TabIndex        =   10
         Top             =   2220
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   688
         Caption         =   "È«Ñ¡(&A)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AwesomeType.Button MnRe 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   1050
         TabIndex        =   9
         Top             =   1860
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   661
         Caption         =   "ÖØ×ö"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AwesomeType.Button mnUn 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   -15
         TabIndex        =   8
         Top             =   1860
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "³·Ïú"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AwesomeType.Button MnCut 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   -15
         TabIndex        =   3
         Top             =   -15
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   688
         Caption         =   "¼ôÌù(&T)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AwesomeType.Button MnCopy 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   -15
         TabIndex        =   4
         Top             =   360
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   688
         Caption         =   "¸´ÖÆ(&C)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AwesomeType.Button mnPaste 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   -15
         TabIndex        =   5
         Top             =   720
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   688
         Caption         =   "Õ³Ìù(&P)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AwesomeType.Button MnDelete 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   -15
         TabIndex        =   6
         Top             =   1080
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   688
         Caption         =   "É¾³ý(&D)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AwesomeType.Button MnSearch 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1440
         TabIndex        =   7
         Top             =   1455
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   741
         Caption         =   "ËÑË÷"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox TxSearch 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   30
         TabIndex        =   2
         Top             =   1515
         Width           =   1425
      End
      Begin AwesomeType.Button MnAdd 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   -15
         TabIndex        =   11
         Top             =   2580
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   688
         Caption         =   "¹ØÓÚ(&B)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13996
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"FrmMain.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'×÷Õß£ºÕÅÓÓ½Ü
'
'Ãû³Æ£ºFrmMain.frm
'
'ÃèÊö£ºAwesomeTypeÖ÷´°¿ÚµÄ´úÂë
'
'ÍøÕ¾£ºhttps://www.johnzhang.xyz/
'
'ÓÊÏä£ºzsgsdesign@gmail.com
'
'×ñÑ­MITÐ­Òé£¬¶þ´Î¿ª·¢Çë×¢Ã÷Ô­×÷Õß£¡
'****************************************************************************

Private WithEvents oSyntax As CSyntax
Attribute oSyntax.VB_VarHelpID = -1
'³·ÏúÓëÖØ×ö
Private trapUndo As Boolean           'flag to indicate whether actions should be trapped
Private UndoStack As New Collection   'collection of undo elements
Private RedoStack As New Collection   'collection of redo elements
'End
Dim FindText As String

Private Sub Form_Load()
trapUndo = True     'Enable Undo Trapping
End Sub

Private Sub Form_Resize()
RTB.Move -2, -2, Me.ScaleWidth + 4, Me.ScaleHeight + 4
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub MnAdd_Click()
FrmBriskAbout.Show
End Sub

Private Sub MnCopy_Click()
Clipboard.Clear
Clipboard.SetText RTB.SelText
PicRMenu.Visible = False
End Sub

Private Sub MnCut_Click()
Clipboard.Clear
Clipboard.SetText RTB.SelText
RTB.SelText = ""
PicRMenu.Visible = False
End Sub

Private Sub MnDelete_Click()
RTB.SelText = ""
PicRMenu.Visible = False
End Sub

Private Sub mnPaste_Click()
RTB.SelText = Clipboard.GetText
PicRMenu.Visible = False
End Sub

Private Sub MnRe_Click()
Redo
End Sub

Private Sub MnSearch_Click()
RTB.SetFocus
If FindText = TxSearch.Text Then
    RTB.SelStart = RTB.SelStart + RTB.SelLength + 1
    RTB.Find TxSearch.Text, , Len(RTB)
Else
    RTB.Find TxSearch.Text
    FindText = TxSearch.Text
End If
End Sub

Private Sub MnSelAll_Click()
RTB.SetFocus
RTB.SelStart = 0
RTB.SelLength = Len(RTB)
PicRMenu.Visible = False
End Sub

Private Sub mnUn_Click()
Undo
End Sub

Private Sub RTB_Change()
On Error Resume Next
'¸ßÁÁ
    Set oSyntax = New CSyntax
    oSyntax.HighLightRichEdit RTB
    Set oSyntax = Nothing
'end
'³·ÏúÓëÖØ×ö
    If Not trapUndo Then Exit Sub 'because trapping is disabled

    Dim newElement As New UndoElement   'create new undo element
    Dim c%, l&

    'remove all redo items because of the change
    For c% = 1 To RedoStack.Count
        RedoStack.Remove 1
    Next c%

    'set the values of the new element
    newElement.SelStart = RTB.SelStart
    newElement.TextLen = Len(RTB.Text)
    newElement.Text = RTB.Text

    'add it to the undo stack
    UndoStack.Add Item:=newElement
    'enable controls accordingly

End Sub

Private Sub RTB_KeyUp(KeyCode As Integer, Shift As Integer)
'·ÀÖ¹richtextbox¾­µäÐÔ¿Õ¸ñbug
   If (Shift = vbCtrlMask And KeyCode = vbKeySpace) Or _
      (KeyCode = vbKeySpace And Shift = vbCtrlMask) Then
      Dim Position As Long
      Dim SelectiveText As Long
      With RTB
           Position = .SelStart
           SelectiveText = .SelLength
          .Text = .Text
          .SelStart = Position
          .SelLength = SelectiveText
      End With
   End If
'end
End Sub

Public Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String    '³·ÏúÓëÖØ×ö
Dim tempParam$
Dim d&
    If Len(lParam1) > Len(lParam2) Then 'swap
        tempParam$ = lParam1
        lParam1 = lParam2
        lParam2 = tempParam$
    End If
    d& = Len(lParam2) - Len(lParam1)
    Change = Mid(lParam2, startSearch - d&, d&)
End Function

Public Sub Undo()    '³·Ïú
Dim chg$, X&
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object, objElement2 As Object
    If UndoStack.Count > 1 And trapUndo Then 'we can proceed
        trapUndo = False
        DeleteFlag = UndoStack(UndoStack.Count - 1).TextLen < UndoStack(UndoStack.Count).TextLen
        If DeleteFlag Then  'delete some text
            X& = SendMessage(RTB.hWnd, EM_HIDESELECTION, 1&, 1&)
            Set objElement = UndoStack(UndoStack.Count)
            Set objElement2 = UndoStack(UndoStack.Count - 1)
            RTB.SelStart = objElement.SelStart - (objElement.TextLen - objElement2.TextLen)
            RTB.SelLength = objElement.TextLen - objElement2.TextLen
            RTB.SelText = ""
            X& = SendMessage(RTB.hWnd, EM_HIDESELECTION, 0&, 0&)
        Else 'append something
            Set objElement = UndoStack(UndoStack.Count - 1)
            Set objElement2 = UndoStack(UndoStack.Count)
            chg$ = Change(objElement.Text, objElement2.Text, _
                objElement2.SelStart + 1 + Abs(Len(objElement.Text) - Len(objElement2.Text)))
            RTB.SelStart = objElement2.SelStart
            RTB.SelLength = 0
            RTB.SelText = chg$
            RTB.SelStart = objElement2.SelStart
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                RTB.SelLength = Len(chg$)
            Else
                RTB.SelStart = RTB.SelStart + Len(chg$)
            End If
        End If
        RedoStack.Add Item:=UndoStack(UndoStack.Count)
        UndoStack.Remove UndoStack.Count
    End If
    
    trapUndo = True
    RTB.SetFocus
End Sub

Public Sub Redo()    'ÖØ×ö
Dim chg$
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object
    If RedoStack.Count > 0 And trapUndo Then
        trapUndo = False
        DeleteFlag = RedoStack(RedoStack.Count).TextLen < Len(RTB.Text)
        If DeleteFlag Then  'delete last item
            Set objElement = RedoStack(RedoStack.Count)
            RTB.SelStart = objElement.SelStart
            RTB.SelLength = Len(RTB.Text) - objElement.TextLen
            RTB.SelText = ""
        Else 'append something
            Set objElement = RedoStack(RedoStack.Count)
            chg$ = Change(RTB.Text, objElement.Text, objElement.SelStart + 1)
            RTB.SelStart = objElement.SelStart - Len(chg$)
            RTB.SelLength = 0
            RTB.SelText = chg$
            RTB.SelStart = objElement.SelStart - Len(chg$)
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                RTB.SelLength = Len(chg$)
            Else
                RTB.SelStart = RTB.SelStart + Len(chg$)
            End If
        End If
        UndoStack.Add Item:=objElement
        RedoStack.Remove RedoStack.Count
    End If
    
    trapUndo = True
    RTB.SetFocus
End Sub

Private Sub RTB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    FindText = ""
    PicRMenu.Move X / 15 + 2, Y / 15 + 2
    If PicRMenu.Left > Me.ScaleWidth - PicRMenu.Width Then PicRMenu.Left = Me.ScaleWidth - PicRMenu.Width
    If PicRMenu.Top > Me.ScaleHeight - PicRMenu.Height Then PicRMenu.Top = Me.ScaleHeight - PicRMenu.Height
    PicRMenu.Visible = True
Else
    PicRMenu.Visible = False
End If
End Sub
