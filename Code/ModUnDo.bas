Attribute VB_Name = "ModUnDo"
'****************************************************************************
'作者：张佑杰
'
'名称：ModUnDo.bas
'
'描述：AwesomeType撤销功能的模块代码
'
'网站：https://www.johnzhang.xyz/
'
'邮箱：zsgsdesign@gmail.com
'
'遵循MIT协议，二次开发请注明原作者！
'****************************************************************************
Public Declare Function SendMessage Lib "User32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long


Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63




