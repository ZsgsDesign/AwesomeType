Attribute VB_Name = "ModUnDo"
'****************************************************************************
'���ߣ����ӽ�
'
'���ƣ�ModUnDo.bas
'
'������AwesomeType�������ܵ�ģ�����
'
'��վ��https://www.johnzhang.xyz/
'
'���䣺zsgsdesign@gmail.com
'
'��ѭMITЭ�飬���ο�����ע��ԭ���ߣ�
'****************************************************************************
Public Declare Function SendMessage Lib "User32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long


Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63




