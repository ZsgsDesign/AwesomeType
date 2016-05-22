VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_color()   As OLE_COLOR

Private m_iCount    As Integer

Public Sub AddColor(ByRef color As OLE_COLOR)
'*添加一种新的颜色
    m_iCount = m_iCount + 1
    ReDim Preserve m_color(1 To m_iCount)
    m_color(m_iCount) = color
End Sub

Public Function getColorPos(ByRef color As OLE_COLOR) As Integer
'*得到一种颜色的排列位置
Dim i   As Integer
    For i = 1 To m_iCount
        If m_color(i) = color Then
            getColorPos = i
            Exit Function
        End If
    Next i
    '*找不到，返回-1
    getColorPos = -1
End Function

Public Function getColorHex(ByRef color As OLE_COLOR) As String
'*得到一种颜色的十六进制字符串
Dim s   As String
    s = Hex(RedColor(color))
    s = String(2 - Len(s), "0") & s
    getColorHex = s
    s = Hex(GreenColor(color))
    s = String(2 - Len(s), "0") & s
    getColorHex = getColorHex & s
    s = Hex(BlueColor(color))
    s = String(2 - Len(s), "0") & s
    getColorHex = getColorHex & s
End Function

'**************************************************************
'*ComStr
'*功能：返回所有颜色在rtf格式中的定义
'*说明：
'*传入参数：
'*    color      类型：OLE_COLOR
'*返回参数：
'*
'*作者：progame  日期：2002-09-29  13:50:55
'**************************************************************
Public Function ComStr() As String
Dim i       As Integer

    For i = 1 To m_iCount
        ComStr = ComStr & "\red" & RedColor(m_color(i)) & "\green" & GreenColor(m_color(i)) & "\blue" & BlueColor(m_color(i)) & ";"
    Next i
End Function

Private Function RedColor(ByRef color As OLE_COLOR) As Long
    RedColor = color And 255
End Function

Private Function GreenColor(ByRef color As OLE_COLOR) As Long
    GreenColor = (color And 65280) / 256
End Function

Private Function BlueColor(ByRef color As OLE_COLOR) As Long
    BlueColor = (color And 16711680) / 65536
End Function

Private Sub Class_Initialize()
'*定义默认的两种颜色
    m_iCount = 2
    AddColor (vbWhite)
    AddColor (vbBlack)
End Sub

