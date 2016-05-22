VERSION 5.00
Begin VB.Form FrmBriskAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin AwesomeType.Button MnSelAll 
      CausesValidation=   0   'False
      Height          =   390
      Left            =   1920
      TabIndex        =   0
      Top             =   2880
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   688
      Caption         =   "关闭(&C)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "版本：张佑杰"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "作者：张佑杰"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AwesomeType是一款简洁的开源VB代码高亮工具"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "AwesomeType"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   4575
   End
   Begin VB.Line Line4 
      X1              =   5992
      X2              =   5992
      Y1              =   0
      Y2              =   3708
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3708
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6000
      Y1              =   3697
      Y2              =   3697
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6000
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "FrmBriskAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'作者：张佑杰
'
'名称：FrmBriskAbout.frm
'
'描述：AwesomeType关于窗口的代码
'
'网站：https://www.johnzhang.xyz/
'
'邮箱：zsgsdesign@gmail.com
'
'遵循MIT协议，二次开发请注明原作者！
'****************************************************************************

Private Sub Form_Load()
Label1.Left = (FrmBriskAbout.ScaleWidth - Label1.Width) / 2
Label2.Left = (FrmBriskAbout.ScaleWidth - Label2.Width) / 2
Label3.Left = (FrmBriskAbout.ScaleWidth - Label3.Width) / 2
Label4.Left = (FrmBriskAbout.ScaleWidth - Label4.Width) / 2
MnSelAll.Left = (FrmBriskAbout.ScaleWidth - MnSelAll.Width) / 2
Label4.Caption = "版本：" & App.Major & "." & App.Minor & "." & App.Revision
FrmMain.PicRMenu.Visible = False
End Sub

Private Sub MnSelAll_Click()
Unload Me
End Sub
