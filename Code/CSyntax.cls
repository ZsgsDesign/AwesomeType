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

Public Event Progress(ByVal Value As Integer) '*处理进度

Private m_initSTX As Boolean            '语法初始化标记
Private m_oString As CStringBuilder     '大字符串连接类

Private m_sTitle            As String   '*语法文件Title

Private m_sLnCommFlag1      As String   '*行注释头一
Private m_sLnCommFlag2      As String   '*行注释头二
Private m_sBlkCommHead1     As String   '*块注释头一
Private m_sBlkCommFoot1     As String   '*块注释尾一
Private m_sBlkCommHead2     As String   '*块注释头二
Private m_sBlkCommFoot2     As String   '*块注释尾二
Private m_objCommFont       As CFont    '*注释字体设置

Private m_sCommRtfHead      As String   '*注释字体RTF格式头
Private m_sCommRtfFoot      As String   '*注释字体RTF格式尾

Private m_objNumFont        As CFont    '*数字字体设置

Private m_sNumRtfHead       As String   '*数字字体RTF格式头
Private m_sNumRtfFoot       As String   '*数字字体RTF格式尾

Private m_iKwdCount         As Integer  '*关键字类型数
Private m_sKwdSet()         As String   '*存入一类关键字的全部组合(以chr(3)分隔)
Private m_sKwdTitle()       As String   '*此关键字类型的名称
Private m_objKwdFont()      As CFont    '*此关键字类型的字体设置
Private m_sKwdRtfHead()     As String   '*关键字字体RTF格式头
Private m_sKwdRtfFoot()     As String   '*关键字字体RTF格式尾

Private m_bCase             As Boolean  '*是否大小字敏感

Private m_sQuotFlag1        As String   '*字符串边界一
Private m_sQuotFlag2        As String   '*字符串边界二
Private m_bMultiLine        As Boolean  '*是否续行有效
Private m_objQuotFont       As CFont    '*字符串字体设置
Private m_sQuotRtfHead      As String   '*字符串字体RTF格式头
Private m_sQuotRtfFoot      As String   '*字符串字体RTF格式尾

Private m_sEscape           As String   '*后面跟敏感字符，此字符后面字符忽略

Private m_objColorManager   As CColor   '*颜色管理类
Private m_sColorRtf         As String   '*RTF内颜色信息

Private m_sWord()         As String     '*对一行字符串分隔成一个词数组

Private m_iQuotState        As Integer  '*当前是否在字符串内[0..不是字符串/1..在字符串一内/2..在字符串二内]
Private m_iCommState        As Integer  '*当前注释状态[0..无注释/1..单行注释一/2..单行注释二/3..多行注释一/4..多行注释二]

Private Const DELIMITER = ",(){}[]-+*%/='~!&|\<>?:;."

'*RTF文件头(First和Last之间插入颜色和字体信息
Private Const HEAD_FIRST = "{\rtf1\ansi\ansicpg936\deff0{\fonttbl}{\colortbl ;"
Private Const HEAD_LAST = "}\viewkind4\uc1\pard\lang2052\f0\fs"  'fs** 控制字体大小

Private Const HEAD_HTML = "<PRE>"
Private Const TAIL_HTML = "</PRE>"

'----------------------VB 语法常量------------------------
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
'*功能：对控件进行语法高亮显示
'*说明：
'*传入参数：
'*    ctrl      类型：RichTextBox
'*返回参数：
'*
'*作者：progame  日期：2002-09-29  17:38:22
'**************************************************************
Public Sub HighLightRichEdit(ByRef ctrl As RichTextBox)
Dim oLoc As Long
Dim oLen As Long
    
    If Not m_initSTX Then ReadSTX     '读入语法
    
    oLoc = ctrl.SelStart
    oLen = ctrl.SelLength

    HighLightCtrl ctrl
    
    ctrl.SelStart = oLoc
    ctrl.SelLength = oLen
End Sub

Private Sub ReadSTX()
'*将语法读入
Dim i As Integer
    
    '读语法细节
    Call ReadDetail
    
    '*对字体的字符串预先生成
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
'*读取语法的具体实现
Dim sTmp        As String
Dim i           As Integer
Dim l           As Integer
Dim lstr        As String
Dim rstr        As String
Dim readArr()   As String

    
    '*读取字符串的字体配置
    readArr = Split(STX_QUOTATION, ",")
    For l = 0 To UBound(readArr)
        sTmp = readArr(l)
        Call PreFormat(sTmp)
        
        '读取详细内容
        i = InStr(sTmp, "=")
        Call ReadFont(sTmp, i, m_objQuotFont)
    Next
    
     '*读取注释的字体配置
    readArr = Split(STX_COMMENT, ",")
    For l = 0 To UBound(readArr)
        sTmp = readArr(l)
        Call PreFormat(sTmp)
            
        '读取详细内容
        i = InStr(sTmp, "=")
        Call ReadFont(sTmp, i, m_objCommFont)
    Next
    
     '*读取数字的字体配置
    readArr = Split(STX_NUMBER, ",")
    For l = 0 To UBound(readArr)
        sTmp = readArr(l)
        Call PreFormat(sTmp)
        
        '读取详细内容
        i = InStr(sTmp, "=")
        Call ReadFont(sTmp, i, m_objNumFont)
    Next
    
    '*读取关键字
    ReadKeys "Keyword", STX_KEYWORD, VB_KEYWORDS
    ReadKeys "Operator", STX_OPERATOR, VB_OPERATORS
    ReadKeys "Function", STX_FUNCTION, VB_FUNCTIONS
    'ReadKeys "Propertie", STX_PROPERTIE, VB_PROPERTIES
    
    '*读取基本配置
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
'*进语法文件预处理，因为到时候代码会先进行格式化以适应"}" "{" "\"
    s = Replace(Replace(Replace(s, "\", "\\"), "{", "\{"), "}", "\}")
End Sub

Private Sub ReadKeys(ByVal KeyName As String, ByVal KeySTX As String, ByVal Keywords As String)
'*读取关键字

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
    
    '关键字的字体配置
    readArr = Split(KeySTX, ",")
    For l = 0 To UBound(readArr)
        sTmp = readArr(l)
        Call PreFormat(sTmp)
        
        i = InStr(sTmp, "=")
        Call ReadFont(sTmp, i, m_objKwdFont(m_iKwdCount))
    Next
    
    '读取关键字的详细内容
    readArr = Split(Keywords, ",")
    For l = 0 To UBound(readArr)
        sTmp = readArr(l)
        m_sKwdSet(m_iKwdCount) = m_sKwdSet(m_iKwdCount) & IIf(m_bCase, Trim(sTmp), UCase(Trim(sTmp))) & Chr(3)
    Next

End Sub

Private Sub ReadFont(ByVal sValue As String, ByVal i As Integer, ByRef Font As CFont)
'*读取文件中的字体设置
'*      i       --"="所在位置
'*      font    --要传入的字体类

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
'*功能：对控件进行语法高亮显示
'*说明：
'*传入参数：
'*    ctrl           类型：RichTextBox      --控件
'*返回参数：
'*
'*作者：progame  日期：2002-09-29  17:38:22
'**************************************************************
Private Sub HighLightCtrl(ByRef ctrl As RichTextBox)
Dim sLine() As String
Dim i       As Long
Dim l       As Long
Dim sTmp        As String
Dim iProgress   As Integer      '*当前处理进度
Dim iTmp        As Integer
'Dim rtfText As String
Dim rtfFontSize As Integer
    
    sTmp = ctrl.Text
    l = Len(sTmp)
    If l = 0 Then Exit Sub      '内容长度为0，退出
    
    rtfFontSize = ctrl.Font.Size    '字体大小
    If rtfFontSize Mod 3 = 2 Then
        rtfFontSize = rtfFontSize * 2 + 1
    Else
        rtfFontSize = rtfFontSize * 2
    End If
    
    '生成 RTF 格式字符
    'rtfText = HEAD_FIRST & m_sColorRtf & HEAD_LAST & CStr(rtfFontSize) & " "
    m_oString.Append HEAD_FIRST & m_sColorRtf & HEAD_LAST & CStr(rtfFontSize) & " "
    
    m_iQuotState = 0
    m_iCommState = 0
    
    iProgress = 0
    
    '*对于大文本进行分块处理，每1K为一个单位
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
        
        '*处理功能字符和分隔字符
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
        
            Call HighLightLine(sLine(j))    '*处理一行
            
            '*恢复分隔字符
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
'*扫描字符，输出格式化后的字符串
Dim i       As Long
Dim j       As Integer
Dim k       As Integer
Dim bComm As Boolean

    m_sWord = Split(sLine, " ")
    k = UBound(m_sWord)
    For i = 0 To k

        '*如果不是字符串，则进行先进行注释判断，再进行关键字判断
        If m_iQuotState = 0 Then
        
            bComm = TreatComment(m_sWord(i), (i = k))         '*处理注释
        
        End If

        If Not bComm Then

                If TreatQuot(m_sWord(i), (i = k), 1) Then       '*处理字符串
                
                    '*恢复被处理过的字符串标志
                    If m_sQuotFlag1 <> "" Then
                        m_sWord(i) = Replace(m_sWord(i), Chr(1), m_sQuotFlag1)
                    End If
                    If m_sQuotFlag2 <> "" Then
                        m_sWord(i) = Replace(m_sWord(i), Chr(2), m_sQuotFlag2)
                    End If
                Else
                
                    If m_iQuotState = 0 Then
                        Call TreatKeyWord(m_sWord(i))         '*处理关键字
                    End If
                    
                End If

        End If
        If i <> k Then m_sWord(i) = m_sWord(i) & " "
        
    Next i
    
    sLine = Join(m_sWord, "")

End Sub

Private Function TreatComment(ByRef sWord As String, ByVal bLast As Boolean) As Boolean
'*处理注释
'*      bLast       --是否最后一个字符
'*      返回值：    sWord是否注释
    Select Case m_iCommState
        Case 0
            If m_sLnCommFlag1 <> "" Then
                If InStr(1, sWord, m_sLnCommFlag1, vbBinaryCompare) = 1 Then  '*是单行注释一
                    m_iCommState = 1
                    '*加上注释字体RTF格式头
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    '*如果最后一行，再进行处理
                    If bLast Then
                        TreatComment sWord, True
                    End If
                    Exit Function
                End If
            End If
            If m_sLnCommFlag2 <> "" Then
                If InStr(1, sWord, m_sLnCommFlag2, vbBinaryCompare) = 1 Then  '*是单行注释二
                    m_iCommState = 2
                    '*加上注释字体RTF格式头
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    '*如果最后一个词，再进行处理
                    If bLast Then
                        TreatComment sWord, True
                    End If
                    Exit Function
                End If
            End If
            If m_sBlkCommHead1 <> "" And m_sBlkCommFoot1 <> "" Then
                If InStr(1, sWord, m_sBlkCommHead1, vbBinaryCompare) = 1 Then  '*是多行注释一
                    m_iCommState = 3
                    '*加上注释字体RTF格式头
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    TreatComment sWord, bLast
                    Exit Function
                End If
            End If
            If m_sBlkCommHead2 <> "" And m_sBlkCommFoot2 <> "" Then
                If InStr(1, sWord, m_sBlkCommHead2, vbBinaryCompare) = 1 Then '*是多行注释二
                    m_iCommState = 4
                    '*加上注释字体RTF格式头
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    TreatComment sWord, bLast
                    Exit Function
                End If
            End If
        Case 1, 2
            TreatComment = True
            If bLast Then       '*单行注释，且当前为此行最后一个词，将注释标志恢复为0
                sWord = sWord & m_sCommRtfFoot
                m_iCommState = 0
            End If
        Case 3
            TreatComment = True
            If m_sBlkCommHead1 <> "" And m_sBlkCommFoot1 <> "" And Len(sWord) > 1 Then
                If InStrRev(sWord, m_sBlkCommFoot1, , vbBinaryCompare) <> 0 Then '*是多行注释一
                    m_iCommState = 0
                    '*加上注释字体RTF格式尾
                    sWord = sWord & m_sCommRtfFoot
                    Exit Function
                End If
            End If
        Case 4
            TreatComment = True
            If m_sBlkCommHead2 <> "" And m_sBlkCommFoot2 <> "" And Len(sWord) > 1 Then
                If InStrRev(sWord, m_sBlkCommFoot2, , vbBinaryCompare) <> 0 Then '*是多行注释二
                    m_iCommState = 0
                    '*加上注释字体RTF格式尾
                    sWord = sWord & m_sCommRtfFoot
                    Exit Function
                End If
            End If
        Case Else
        
    End Select
    
End Function

Private Function TreatKeyWord(ByRef sWord As String) As Boolean
'*处理关键字
'*      返回值：    sWord是否关键字
Dim i           As Integer
Dim bKeyWord    As Boolean
    '* 先处理是不是数字
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
            '*加上字体RTF格式的头尾于词两端
            sWord = m_sKwdRtfHead(i) & sWord & m_sKwdRtfFoot(i)
            TreatKeyWord = True
            Exit Function
        End If
    Next i
End Function

Private Function TreatQuot(ByRef sWord As String, ByVal bLast As Boolean, ByVal iBegin As Integer) As Boolean
'*处理字符串
'*      bLast       是否此行最后一个词
'*      iBegin      搜索位置
'*      返回值：    是否做过字符串处理
Dim i       As Integer
Dim j       As Integer
Dim sQuot   As String       '*当前用来判断的字符串边界符
    Select Case m_iQuotState
        Case 0
            '*判断哪种字符串标志先出现，使用先出现的字符串标志来决定字符串起始位置
            i = 0
            j = 0
            If m_sQuotFlag1 <> "" Then
                i = InStr(iBegin, sWord, m_sQuotFlag1, vbBinaryCompare)
            End If
            If m_sQuotFlag2 <> "" Then
                j = InStr(iBegin, sWord, m_sQuotFlag2, vbBinaryCompare)
            End If
            If i = 0 And j = 0 Then     '*没有字符串标志出现，不用判断
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
    
    '*查找当前的字符串标志
    i = InStr(iBegin, sWord, sQuot, vbBinaryCompare)
    
    '*如果前一个字符是敏感字符前缀，则此次查找无效
    If m_sEscape <> "" Then
        If i > 1 Then
            j = InStr(iBegin, sWord, m_sEscape)
            If j <> 0 And j = i - Len(m_sEscape) Then
                '*将此字符串标志屏避
                If sQuot = m_sQuotFlag1 Then
                    sWord = Replace(sWord, sQuot, Chr(1), , 1)
                Else
                    sWord = Replace(sWord, sQuot, Chr(2), , 1)
                End If
                TreatQuot = True
                '*继续递归处理
                Call TreatQuot(sWord, bLast, i + 1)
                Exit Function
            End If
        End If
    End If
    
    If i <> 0 Then
    
        TreatQuot = True
        
        '*格式化字符串
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

        '*继续递归处理
        Call TreatQuot(sWord, bLast, i)
        
    Else
    
        '*如果已经没有了标志，但续行无效的情况下，当前这个词为最后一个时，强行结束
        If bLast And Not m_bMultiLine And m_iQuotState <> 0 Then
            sWord = sWord & m_sQuotRtfFoot
            m_iQuotState = 0
        End If
        
    End If

End Function



