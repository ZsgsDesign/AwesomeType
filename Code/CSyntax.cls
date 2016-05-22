VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSyntax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Compare Binary

Public Event Progress(ByVal Value As Integer) '*��������

Private m_initSTX As Boolean            '�﷨��ʼ�����
Private m_oString As CStringBuilder     '���ַ���������

Private m_sTitle            As String   '*�﷨�ļ�Title

Private m_sLnCommFlag1      As String   '*��ע��ͷһ
Private m_sLnCommFlag2      As String   '*��ע��ͷ��
Private m_sBlkCommHead1     As String   '*��ע��ͷһ
Private m_sBlkCommFoot1     As String   '*��ע��βһ
Private m_sBlkCommHead2     As String   '*��ע��ͷ��
Private m_sBlkCommFoot2     As String   '*��ע��β��
Private m_objCommFont       As CFont    '*ע����������

Private m_sCommRtfHead      As String   '*ע������RTF��ʽͷ
Private m_sCommRtfFoot      As String   '*ע������RTF��ʽβ

Private m_objNumFont        As CFont    '*������������

Private m_sNumRtfHead       As String   '*��������RTF��ʽͷ
Private m_sNumRtfFoot       As String   '*��������RTF��ʽβ

Private m_iKwdCount         As Integer  '*�ؼ���������
Private m_sKwdSet()         As String   '*����һ��ؼ��ֵ�ȫ�����(��chr(3)�ָ�)
Private m_sKwdTitle()       As String   '*�˹ؼ������͵�����
Private m_objKwdFont()      As CFont    '*�˹ؼ������͵���������
Private m_sKwdRtfHead()     As String   '*�ؼ�������RTF��ʽͷ
Private m_sKwdRtfFoot()     As String   '*�ؼ�������RTF��ʽβ

Private m_bCase             As Boolean  '*�Ƿ��С������

Private m_sQuotFlag1        As String   '*�ַ����߽�һ
Private m_sQuotFlag2        As String   '*�ַ����߽��
Private m_bMultiLine        As Boolean  '*�Ƿ�������Ч
Private m_objQuotFont       As CFont    '*�ַ�����������
Private m_sQuotRtfHead      As String   '*�ַ�������RTF��ʽͷ
Private m_sQuotRtfFoot      As String   '*�ַ�������RTF��ʽβ

Private m_sEscape           As String   '*����������ַ������ַ������ַ�����

Private m_objColorManager   As CColor   '*��ɫ������
Private m_sColorRtf         As String   '*RTF����ɫ��Ϣ

Private m_sWord()         As String     '*��һ���ַ����ָ���һ��������

Private m_iQuotState        As Integer  '*��ǰ�Ƿ����ַ�����[0..�����ַ���/1..���ַ���һ��/2..���ַ�������]
Private m_iCommState        As Integer  '*��ǰע��״̬[0..��ע��/1..����ע��һ/2..����ע�Ͷ�/3..����ע��һ/4..����ע�Ͷ�]

Private Const DELIMITER = ",(){}[]-+*%/='~!&|\<>?:;."

'*RTF�ļ�ͷ(First��Last֮�������ɫ��������Ϣ
Private Const HEAD_FIRST = "{\rtf1\ansi\ansicpg936\deff0{\fonttbl}{\colortbl ;"
Private Const HEAD_LAST = "}\viewkind4\uc1\pard\lang2052\f0\fs"  'fs** ���������С

Private Const HEAD_HTML = "<PRE>"
Private Const TAIL_HTML = "</PRE>"

'----------------------VB �﷨����------------------------
Private Const STX_BASE = "#QUOTATION1="",#QUOTATION2=,#CONTINUE_QUOTE=,#LINECOMMENT=',#LINECOMMENT2=rem,#COMMENTON=,#COMMENTOFF=,#COMMENTON2=,#COMMENTOFF2=,#ESCAPE=,#CASE=n"
Private Const STX_QUOTATION = "ForeColor=0,backcolor=,bold=,italic=,underline=,strike="
Private Const STX_COMMENT = "ForeColor=33280,backcolor=,bold=,italic=,underline=,strike="
Private Const STX_NUMBER = "ForeColor=0,backcolor=,bold=,italic=,underline=,strike="
Private Const STX_OPERATOR = "ForeColor=0,backcolor=,bold=,italic=,underline=,strike="
Private Const STX_KEYWORD = "ForeColor=8388608,backcolor=,bold=,italic=,underline=,strike="
Private Const STX_FUNCTION = "ForeColor=8388608,backcolor=,bold=,italic=,underline=,strike="

Private Const VB_OPERATORS = "&,+,-,*,/,\,^,>,<,>=,<=,<>,=,%"
Private Const VB_KEYWORDS = "Access,AddressOf,Alias,And,Any,AppActivate,Append,Array,As,Base,Beep,Binary,Boolean,ByRef,ByVal,Byte," _
        & "Cancel,Call,Case,ChDir,ChDrive,Choose,Close,Collection,Compare,Const,Currency,Declare,Dim,Do,DoEvents,Double," _
        & "Each,Else,ElseIf,Empty,End,Enum,Environ,Erase,Err,Error,Exit,Explicit,False,FileCopy,For,Form,Function,Get,GoSub,GoTo," _
        & "If,In,Input,Integer,Is,Kill,Let,Lib,Like,Line,Load,Lock,Long,Loop,Me,MkDir,Mod,Module,MsgBox,New,Next," _
        & "Not,Nothing,Null,Object,On,Open,Option,Optional,Or,Output,Print,Print,Private,Property,Public,Put," _
        & "ReDim,Read,Reset,Resume,Return,RmDir,Seek,Select,SendKeys,Set,SetAttr,Shared,Shell,Single,Static,Step,Stop,String," _
        & "Sub,Switch,Tab,Then,To,True,Type,TypeName,Unlock,Until,Variant,Wend,While,With,Write,Xor"
Private Const VB_FUNCTIONS = "Abs,Asc,Atn,CBool,CByte,CCur,CDate,CDbl,CDec,CInt,CLng,CSng,CStr,CVErr,CVar,Chr,Command," _
        & "Cos,CreateObject,CurDir,DDB,Date,DateAdd,DateDiff,DatePart,DateSerial,DateValue,Day,DeleteSetting,EOF," _
        & "Eqv,Exp,FV,FileAttr,FileDateTime,FileLen,Filter,Fix,Format,FormatCurrency,FormatDateTime,FormatNumber,FormatPercent,FreeFile," _
        & "GetAllSettings,GetAttr,GetObject,GetSetting,Hex,Hour,IIf,IPmt,IRR,Imp,InStr,InStrRev,InputBox," _
        & "IsArray,IsDate,IsEmpty,IsError,IsMissing,IsNull,IsNumeric,IsObject,Join,LBound,LCase,LOF,LSet,LTrim,Left,Len," _
        & "LoadPicture,LoadResData,LoadResPicture,LoadResString,Loc,Log,MIRR,MacID,MacScript,Mid,Minute,Month,MonthName," _
        & "NPV,NPer,Now,Oct,PPmt,QBColor,RGB,RSet,RTrim,Random,Randomize,Rate,Replace,Right,Rnd,Round,SLN,SYD,SaveSetting,Second," _
        & "Sgn,Sin,Space,Split,Sqr,StrComp,StrConv,StrReverse,Tan,Time,TimeSerial,TimeValue,UBound,UCase,Val,VarPtr,VarType," _
        & "Weekday,WeekdayName,,Year"

'**************************************************************
'*HighLightRichEdit
'*���ܣ��Կؼ������﷨������ʾ
'*˵����
'*���������
'*    ctrl      ���ͣ�RichTextBox
'*���ز�����
'*
'*���ߣ�progame  ���ڣ�2002-09-29  17:38:22
'**************************************************************
Public Sub HighLightRichEdit(ByRef ctrl As RichTextBox)
Dim oLoc As Long
Dim oLen As Long
    
    If Not m_initSTX Then ReadSTX     '�����﷨
    
    oLoc = ctrl.SelStart
    oLen = ctrl.SelLength

    HighLightCtrl ctrl
    
    ctrl.SelStart = oLoc
    ctrl.SelLength = oLen
End Sub

Private Sub ReadSTX()
'*���﷨����
Dim i As Integer
    
    '���﷨ϸ��
    Call ReadDetail
    
    '*��������ַ���Ԥ������
    m_sQuotRtfHead = m_objQuotFont.getRtfHead(m_objColorManager)
    m_sQuotRtfFoot = m_objQuotFont.getRtfTail
    m_sCommRtfHead = m_objCommFont.getRtfHead(m_objColorManager)
    m_sCommRtfFoot = m_objCommFont.getRtfTail
    
    m_sNumRtfHead = m_objNumFont.getRtfHead(m_objColorManager)
    m_sNumRtfFoot = m_objNumFont.getRtfTail
    
    m_sColorRtf = m_objColorManager.ComStr
    
    For i = 1 To m_iKwdCount
        m_sKwdRtfHead(i) = m_objKwdFont(i).getRtfHead(m_objColorManager)
        m_sKwdRtfFoot(i) = m_objKwdFont(i).getRtfTail
    Next i
    
    m_initSTX = True
End Sub

Private Sub ReadDetail()
'*��ȡ�﷨�ľ���ʵ��
Dim sTmp        As String
Dim i           As Integer
Dim l           As Integer
Dim lstr        As String
Dim rstr        As String
Dim readArr()   As String

    
    '*��ȡ�ַ�������������
    readArr = Split(STX_QUOTATION, ",")
    For l = 0 To UBound(readArr)
        sTmp = readArr(l)
        Call PreFormat(sTmp)
        
        '��ȡ��ϸ����
        i = InStr(sTmp, "=")
        Call ReadFont(sTmp, i, m_objQuotFont)
    Next
    
     '*��ȡע�͵���������
    readArr = Split(STX_COMMENT, ",")
    For l = 0 To UBound(readArr)
        sTmp = readArr(l)
        Call PreFormat(sTmp)
            
        '��ȡ��ϸ����
        i = InStr(sTmp, "=")
        Call ReadFont(sTmp, i, m_objCommFont)
    Next
    
     '*��ȡ���ֵ���������
    readArr = Split(STX_NUMBER, ",")
    For l = 0 To UBound(readArr)
        sTmp = readArr(l)
        Call PreFormat(sTmp)
        
        '��ȡ��ϸ����
        i = InStr(sTmp, "=")
        Call ReadFont(sTmp, i, m_objNumFont)
    Next
    
    '*��ȡ�ؼ���
    ReadKeys "Keyword", STX_KEYWORD, VB_KEYWORDS
    ReadKeys "Operator", STX_OPERATOR, VB_OPERATORS
    ReadKeys "Function", STX_FUNCTION, VB_FUNCTIONS
    'ReadKeys "Propertie", STX_PROPERTIE, VB_PROPERTIES
    
    '*��ȡ��������
    readArr = Split(STX_BASE, ",")
    For l = 0 To UBound(readArr)
        sTmp = readArr(l)
        i = InStr(sTmp, "=")
        
        lstr = Trim(Left(sTmp, i - 1))
        rstr = Trim(Right(sTmp, Len(sTmp) - i))
        
        Select Case UCase(lstr)
            Case "#TITLE"
                rstr = rstr
            Case "#QUOTATION1"
                m_sQuotFlag1 = rstr
            Case "#QUOTATION2"
                m_sQuotFlag2 = rstr
            Case "#CONTINUE_QUOTE"
                m_bMultiLine = IIf(LCase(rstr) = "y", True, False)
            Case "#LINECOMMENT"
                m_sLnCommFlag1 = rstr
            Case "#LINECOMMENT2"
                m_sLnCommFlag2 = rstr
            Case "#COMMENTON"
                m_sBlkCommHead1 = rstr
            Case "#COMMENTOFF"
                m_sBlkCommFoot1 = rstr
            Case "#COMMENTON2"
                m_sBlkCommHead2 = rstr
            Case "#COMMENTOFF2"
                m_sBlkCommFoot2 = rstr
            Case "#CASE"
                m_bCase = IIf(LCase(rstr) = "y", True, False)
            Case "#ESCAPE"
                m_sEscape = rstr
            Case Else
               '
        End Select
    Next
End Sub

Private Sub PreFormat(ByRef s As String)
'*���﷨�ļ�Ԥ��������Ϊ��ʱ�������Ƚ��и�ʽ������Ӧ"}" "{" "\"
    s = Replace(Replace(Replace(s, "\", "\\"), "{", "\{"), "}", "\}")
End Sub

Private Sub ReadKeys(ByVal KeyName As String, ByVal KeySTX As String, ByVal Keywords As String)
'*��ȡ�ؼ���

Dim sTmp As String
Dim readArr() As String
Dim i As Integer
Dim l As Integer

    sTmp = KeyName
    m_iKwdCount = m_iKwdCount + 1
    
    ReDim Preserve m_sKwdSet(1 To m_iKwdCount)
    ReDim Preserve m_sKwdTitle(1 To m_iKwdCount)
    ReDim Preserve m_objKwdFont(1 To m_iKwdCount)
    ReDim Preserve m_sKwdRtfHead(1 To m_iKwdCount)
    ReDim Preserve m_sKwdRtfFoot(1 To m_iKwdCount)
    
    Set m_objKwdFont(m_iKwdCount) = New CFont
    
    m_sKwdTitle(m_iKwdCount) = sTmp
    m_sKwdSet(m_iKwdCount) = Chr(3)
    
    '�ؼ��ֵ���������
    readArr = Split(KeySTX, ",")
    For l = 0 To UBound(readArr)
        sTmp = readArr(l)
        Call PreFormat(sTmp)
        
        i = InStr(sTmp, "=")
        Call ReadFont(sTmp, i, m_objKwdFont(m_iKwdCount))
    Next
    
    '��ȡ�ؼ��ֵ���ϸ����
    readArr = Split(Keywords, ",")
    For l = 0 To UBound(readArr)
        sTmp = readArr(l)
        m_sKwdSet(m_iKwdCount) = m_sKwdSet(m_iKwdCount) & IIf(m_bCase, Trim(sTmp), UCase(Trim(sTmp))) & Chr(3)
    Next

End Sub

Private Sub ReadFont(ByVal sValue As String, ByVal i As Integer, ByRef Font As CFont)
'*��ȡ�ļ��е���������
'*      i       --"="����λ��
'*      font    --Ҫ�����������

Dim lstr        As String
Dim rstr        As String
    lstr = Trim(Left(sValue, i - 1))
    rstr = Trim(Right(sValue, Len(sValue) - i))
    If lstr <> "" And rstr <> "" Then
        Select Case LCase(lstr)
            Case "forecolor"
                Font.bForecolor = True
                Font.ForeColor = rstr
                m_objColorManager.AddColor CLng(rstr)
            Case "backcolor"
                Font.bBackcolor = True
                Font.BackColor = rstr
                m_objColorManager.AddColor CLng(rstr)
            Case "bold"
                Font.bold = IIf(LCase(rstr) = "y", True, False)
            Case "italic"
                Font.italic = IIf(LCase(rstr) = "y", True, False)
            Case "underline"
                Font.underline = IIf(LCase(rstr) = "y", True, False)
            Case "strike"
                Font.strike = IIf(LCase(rstr) = "y", True, False)
            Case Else
                '
                
        End Select
    End If
End Sub

Private Sub Class_Initialize()
    Set m_objColorManager = New CColor
    Set m_objCommFont = New CFont
    Set m_objQuotFont = New CFont
    Set m_objNumFont = New CFont
    Set m_oString = New CStringBuilder
End Sub

Private Sub Class_Terminate()
Dim i   As Integer

    Set m_objColorManager = Nothing
    Set m_objCommFont = Nothing
    Set m_objQuotFont = Nothing
    Set m_objNumFont = Nothing
    Set m_oString = Nothing
    
    For i = 1 To m_iKwdCount
        Set m_objKwdFont(i) = Nothing
    Next i
    
    Erase m_objKwdFont
    Erase m_sKwdTitle
    Erase m_sKwdSet
    Erase m_sKwdRtfHead
    Erase m_sKwdRtfFoot
    Erase m_sWord
End Sub

'**************************************************************
'*HighLightCtrl
'*���ܣ��Կؼ������﷨������ʾ
'*˵����
'*���������
'*    ctrl           ���ͣ�RichTextBox      --�ؼ�
'*���ز�����
'*
'*���ߣ�progame  ���ڣ�2002-09-29  17:38:22
'**************************************************************
Private Sub HighLightCtrl(ByRef ctrl As RichTextBox)
Dim sLine() As String
Dim i       As Long
Dim l       As Long
Dim sTmp        As String
Dim iProgress   As Integer      '*��ǰ��������
Dim iTmp        As Integer
'Dim rtfText As String
Dim rtfFontSize As Integer
    
    sTmp = ctrl.Text
    l = Len(sTmp)
    If l = 0 Then Exit Sub      '���ݳ���Ϊ0���˳�
    
    rtfFontSize = ctrl.Font.Size    '�����С
    If rtfFontSize Mod 3 = 2 Then
        rtfFontSize = rtfFontSize * 2 + 1
    Else
        rtfFontSize = rtfFontSize * 2
    End If
    
    '���� RTF ��ʽ�ַ�
    'rtfText = HEAD_FIRST & m_sColorRtf & HEAD_LAST & CStr(rtfFontSize) & " "
    m_oString.Append HEAD_FIRST & m_sColorRtf & HEAD_LAST & CStr(rtfFontSize) & " "
    
    m_iQuotState = 0
    m_iCommState = 0
    
    iProgress = 0
    
    '*���ڴ��ı����зֿ鴦����ÿ1KΪһ����λ
    Dim aTmp()      As String
    Dim j           As Long
    Dim lstPos      As Long
    lstPos = 1
    ReDim aTmp(1 To l / 4000 + 1)
    For i = 1 To UBound(aTmp)
        j = InStr(lstPos + 4000, sTmp, vbCrLf, vbTextCompare)
        If j = 0 Then
            j = l
            aTmp(i) = Mid(sTmp, lstPos, j - lstPos + 1)
            lstPos = -1
        Else
            aTmp(i) = Mid(sTmp, lstPos, j - lstPos + 1)
            lstPos = j + 1
        End If
        
        '*���������ַ��ͷָ��ַ�
        aTmp(i) = Replace(aTmp(i), "\", "\\")
        aTmp(i) = Replace(aTmp(i), "{", "\{")
        aTmp(i) = Replace(aTmp(i), "}", "\}")
        aTmp(i) = Replace(aTmp(i), "(", " ( ")
        aTmp(i) = Replace(aTmp(i), ")", " ) ")
        aTmp(i) = Replace(aTmp(i), ";", " ; ")
        aTmp(i) = Replace(aTmp(i), "#", " # ")
        aTmp(i) = Replace(aTmp(i), Chr(9), " " & Chr(9) & " ")
        aTmp(i) = Replace(aTmp(i), vbTab, " " & vbTab & " ")
    
        sLine = Split(aTmp(i), vbCrLf)
        
        For j = 0 To UBound(sLine)
        
            Call HighLightLine(sLine(j))    '*����һ��
            
            '*�ָ��ָ��ַ�
            aTmp(i) = Replace(sLine(j), " ( ", "(")
            aTmp(i) = Replace(aTmp(i), " ) ", ")")
            aTmp(i) = Replace(aTmp(i), " ; ", ";")
            aTmp(i) = Replace(aTmp(i), " # ", "#")
            aTmp(i) = Replace(aTmp(i), " " & Chr(9) & " ", Chr(9))
            aTmp(i) = Replace(aTmp(i), " " & vbTab & " ", vbTab)
            
            'rtfText = rtfText & aTmp(i) & "\par "
            aTmp(i) = aTmp(i) & "\par "
            m_oString.Append aTmp(i)
    
        Next j
        
        If lstPos = -1 Then
            Exit For
        End If
        
        iTmp = i * 100 / UBound(aTmp)
        If iProgress <> iTmp Then
            iProgress = iTmp
            RaiseEvent Progress(iProgress)
        End If
            
    Next i
    
    Erase sLine
    
    'ctrl.TextRTF = rtfText & "}"
    ctrl.TextRTF = m_oString.toString & "}"
    m_oString.Reset
    
End Sub

Private Sub HighLightLine(ByRef sLine As String)
'*ɨ���ַ��������ʽ������ַ���
Dim i       As Long
Dim j       As Integer
Dim k       As Integer
Dim bComm As Boolean

    m_sWord = Split(sLine, " ")
    k = UBound(m_sWord)
    For i = 0 To k

        '*��������ַ�����������Ƚ���ע���жϣ��ٽ��йؼ����ж�
        If m_iQuotState = 0 Then
        
            bComm = TreatComment(m_sWord(i), (i = k))         '*����ע��
        
        End If

        If Not bComm Then

                If TreatQuot(m_sWord(i), (i = k), 1) Then       '*�����ַ���
                
                    '*�ָ������������ַ�����־
                    If m_sQuotFlag1 <> "" Then
                        m_sWord(i) = Replace(m_sWord(i), Chr(1), m_sQuotFlag1)
                    End If
                    If m_sQuotFlag2 <> "" Then
                        m_sWord(i) = Replace(m_sWord(i), Chr(2), m_sQuotFlag2)
                    End If
                Else
                
                    If m_iQuotState = 0 Then
                        Call TreatKeyWord(m_sWord(i))         '*�����ؼ���
                    End If
                    
                End If

        End If
        If i <> k Then m_sWord(i) = m_sWord(i) & " "
        
    Next i
    
    sLine = Join(m_sWord, "")

End Sub

Private Function TreatComment(ByRef sWord As String, ByVal bLast As Boolean) As Boolean
'*����ע��
'*      bLast       --�Ƿ����һ���ַ�
'*      ����ֵ��    sWord�Ƿ�ע��
    Select Case m_iCommState
        Case 0
            If m_sLnCommFlag1 <> "" Then
                If InStr(1, sWord, m_sLnCommFlag1, vbBinaryCompare) = 1 Then  '*�ǵ���ע��һ
                    m_iCommState = 1
                    '*����ע������RTF��ʽͷ
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    '*������һ�У��ٽ��д���
                    If bLast Then
                        TreatComment sWord, True
                    End If
                    Exit Function
                End If
            End If
            If m_sLnCommFlag2 <> "" Then
                If InStr(1, sWord, m_sLnCommFlag2, vbBinaryCompare) = 1 Then  '*�ǵ���ע�Ͷ�
                    m_iCommState = 2
                    '*����ע������RTF��ʽͷ
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    '*������һ���ʣ��ٽ��д���
                    If bLast Then
                        TreatComment sWord, True
                    End If
                    Exit Function
                End If
            End If
            If m_sBlkCommHead1 <> "" And m_sBlkCommFoot1 <> "" Then
                If InStr(1, sWord, m_sBlkCommHead1, vbBinaryCompare) = 1 Then  '*�Ƕ���ע��һ
                    m_iCommState = 3
                    '*����ע������RTF��ʽͷ
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    TreatComment sWord, bLast
                    Exit Function
                End If
            End If
            If m_sBlkCommHead2 <> "" And m_sBlkCommFoot2 <> "" Then
                If InStr(1, sWord, m_sBlkCommHead2, vbBinaryCompare) = 1 Then '*�Ƕ���ע�Ͷ�
                    m_iCommState = 4
                    '*����ע������RTF��ʽͷ
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    TreatComment sWord, bLast
                    Exit Function
                End If
            End If
        Case 1, 2
            TreatComment = True
            If bLast Then       '*����ע�ͣ��ҵ�ǰΪ�������һ���ʣ���ע�ͱ�־�ָ�Ϊ0
                sWord = sWord & m_sCommRtfFoot
                m_iCommState = 0
            End If
        Case 3
            TreatComment = True
            If m_sBlkCommHead1 <> "" And m_sBlkCommFoot1 <> "" And Len(sWord) > 1 Then
                If InStrRev(sWord, m_sBlkCommFoot1, , vbBinaryCompare) <> 0 Then '*�Ƕ���ע��һ
                    m_iCommState = 0
                    '*����ע������RTF��ʽβ
                    sWord = sWord & m_sCommRtfFoot
                    Exit Function
                End If
            End If
        Case 4
            TreatComment = True
            If m_sBlkCommHead2 <> "" And m_sBlkCommFoot2 <> "" And Len(sWord) > 1 Then
                If InStrRev(sWord, m_sBlkCommFoot2, , vbBinaryCompare) <> 0 Then '*�Ƕ���ע�Ͷ�
                    m_iCommState = 0
                    '*����ע������RTF��ʽβ
                    sWord = sWord & m_sCommRtfFoot
                    Exit Function
                End If
            End If
        Case Else
        
    End Select
    
End Function

Private Function TreatKeyWord(ByRef sWord As String) As Boolean
'*�����ؼ���
'*      ����ֵ��    sWord�Ƿ�ؼ���
Dim i           As Integer
Dim bKeyWord    As Boolean
    '* �ȴ����ǲ�������
    If IsNumeric(sWord) Then
        sWord = m_sNumRtfHead & sWord & m_sNumRtfFoot
        Exit Function
    End If
    
    For i = 1 To m_iKwdCount
        If m_bCase Then
            If (InStr(1, m_sKwdSet(i), Chr(3) & sWord & Chr(3), vbBinaryCompare) <> 0) Then
               bKeyWord = True
            End If
        Else
            If (InStr(1, m_sKwdSet(i), Chr(3) & UCase(sWord) & Chr(3), vbBinaryCompare) <> 0) Then
                bKeyWord = True
            End If
        End If
        
        If bKeyWord Then
            '*��������RTF��ʽ��ͷβ�ڴ�����
            sWord = m_sKwdRtfHead(i) & sWord & m_sKwdRtfFoot(i)
            TreatKeyWord = True
            Exit Function
        End If
    Next i
End Function

Private Function TreatQuot(ByRef sWord As String, ByVal bLast As Boolean, ByVal iBegin As Integer) As Boolean
'*�����ַ���
'*      bLast       �Ƿ�������һ����
'*      iBegin      ����λ��
'*      ����ֵ��    �Ƿ������ַ�������
Dim i       As Integer
Dim j       As Integer
Dim sQuot   As String       '*��ǰ�����жϵ��ַ����߽��
    Select Case m_iQuotState
        Case 0
            '*�ж������ַ�����־�ȳ��֣�ʹ���ȳ��ֵ��ַ�����־�������ַ�����ʼλ��
            i = 0
            j = 0
            If m_sQuotFlag1 <> "" Then
                i = InStr(iBegin, sWord, m_sQuotFlag1, vbBinaryCompare)
            End If
            If m_sQuotFlag2 <> "" Then
                j = InStr(iBegin, sWord, m_sQuotFlag2, vbBinaryCompare)
            End If
            If i = 0 And j = 0 Then     '*û���ַ�����־���֣������ж�
                Exit Function
            End If
            If (j = 0) Or (i <> 0 And i < j) Then
                sQuot = m_sQuotFlag1
            End If
            If (j <> 0 And j < i) Or (i = 0) Then
                sQuot = m_sQuotFlag2
            End If
        Case 1
            sQuot = m_sQuotFlag1
        Case 2
            sQuot = m_sQuotFlag2
        Case Else
    End Select
    
    '*���ҵ�ǰ���ַ�����־
    i = InStr(iBegin, sWord, sQuot, vbBinaryCompare)
    
    '*���ǰһ���ַ��������ַ�ǰ׺����˴β�����Ч
    If m_sEscape <> "" Then
        If i > 1 Then
            j = InStr(iBegin, sWord, m_sEscape)
            If j <> 0 And j = i - Len(m_sEscape) Then
                '*�����ַ�����־����
                If sQuot = m_sQuotFlag1 Then
                    sWord = Replace(sWord, sQuot, Chr(1), , 1)
                Else
                    sWord = Replace(sWord, sQuot, Chr(2), , 1)
                End If
                TreatQuot = True
                '*�����ݹ鴦��
                Call TreatQuot(sWord, bLast, i + 1)
                Exit Function
            End If
        End If
    End If
    
    If i <> 0 Then
    
        TreatQuot = True
        
        '*��ʽ���ַ���
        If sQuot = m_sQuotFlag1 Then
            If m_iQuotState = 0 Then
                sWord = Replace(sWord, sQuot, m_sQuotRtfHead & Chr(1), , 1)
                m_iQuotState = 1
            Else
                sWord = Replace(sWord, sQuot, Chr(1) & m_sQuotRtfFoot, , 1)
                m_iQuotState = 0
            End If
        Else
            If m_iQuotState = 0 Then
                sWord = Replace(sWord, sQuot, m_sQuotRtfHead & Chr(2), , 1)
                m_iQuotState = 2
            Else
                sWord = Replace(sWord, sQuot, Chr(2) & m_sQuotRtfFoot, , 1)
                m_iQuotState = 0
            End If
        End If

        '*�����ݹ鴦��
        Call TreatQuot(sWord, bLast, i)
        
    Else
    
        '*����Ѿ�û���˱�־����������Ч������£���ǰ�����Ϊ���һ��ʱ��ǿ�н���
        If bLast And Not m_bMultiLine And m_iQuotState <> 0 Then
            sWord = sWord & m_sQuotRtfFoot
            m_iQuotState = 0
        End If
        
    End If

End Function


