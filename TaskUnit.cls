' TaskUnit.cls
Option Explicit

' task的基本信息
Public taskId As Integer         ' taskId int
Public taskFunc As String        ' taskId -> func name
Public taskWorkbook As String    ' taskId -> workbookInfo instance
Public taskStartArgs As Variant  ' taskId -> startArgs array
Public taskResumeSpec As Variant ' taskId -> resumeSpec array
Public taskCell As String        ' taskId -> taskCell address
Public taskStatus As CoStatus    ' taskId -> status string:defined、yield、paused、done、error、terminated
Public taskProgress As Double    ' taskId -> progress number
Public taskMessage As String     ' taskId -> message variant
Public taskValue As Variant      ' taskId -> value variant
Public taskError As String       ' taskId -> error message
Public taskCoThread As LongPtr   ' taskId -> coThread LongPtr
Public taskCoRef As Long         ' Lua 注册表引用号（防止 GC）
' task的调度统计
Public taskLastTime As Double    ' 任务上次运行时间(ms)
Public taskTotalTime As Double   ' 任务总运行时间(ms)
Public taskTickCount As Long     ' 任务调度次数
' CFS 调度字段
Public CFS_vruntime As Double      ' 虚拟运行时间（核心调度依据）
Public CFS_weight As Double        ' 任务权重（默认1024，越大优先级越高）
Public CFS_lastScheduled As Double ' 上次被调度的时间戳（GetTickCount）
' task 的监视器集合
Public taskWatches As Collection ' 任务监视的Watchcell字符串的集合

Private m_ResumeSpecParsed As Variant

Private Sub Class_Initialize()
    taskCoRef = 0
    taskCoThread = 0
    CFS_vruntime = 0
    CFS_lastScheduled = 0
    Set taskWatches = New Collection
    m_ResumeSpecParsed = Array()
End Sub

' 清除协程引用
Public Sub ClearCoroutineRef()
    taskCoRef = 0
    taskCoThread = 0
End Sub

' 添加监控单元格
Public Sub AddWatch(watchCell As String)
    Dim i As Long
    On Error Resume Next
    For i = 1 To Me.taskWatches.Count
        If Me.taskWatches(i) = watchCell Then Exit Sub
    Next i
    On Error GoTo 0
    Me.taskWatches.Add watchCell
End Sub

Public Sub RemoveWatch(watchCell As String)
    Dim i As Long
    On Error Resume Next
    For i = Me.taskWatches.Count To 1 Step -1
        If Me.taskWatches(i) = watchCell Then
            Me.taskWatches.Remove i
            Exit Sub
        End If
    Next i
End Sub

' 获取解析后的参数规格（只读属性）
Public Property Get ResumeSpecParsed() As Variant
    ResumeSpecParsed = m_ResumeSpecParsed
End Property

' 统一参数规格解析
' 参数类型：
'   PARAM_LITERAL (0)    - 字面量（数值、布尔、普通字符串）
'   PARAM_REF (1)        - 单元格/区域引用（$数量=0，来自Range对象或Dictionary）,动态引用（$数量>=1，存储引用层数）
Public Sub ParseResumeSpecs(specs As Variant, callerWb As Workbook, callerWs As Worksheet)
    On Error GoTo ErrorHandler

    ' 安全获取数组边界
    Dim lb As Long, ub As Long
    On Error Resume Next
    lb = LBound(specs): ub = UBound(specs)
    If Err.Number <> 0 Then m_ResumeSpecParsed = Array(): Exit Sub
    On Error GoTo ErrorHandler

    If ub < lb Then m_ResumeSpecParsed = Array(): Exit Sub

    Dim parsed() As Variant
    ReDim parsed(lb To ub)

    Dim i As Long, dollarCount As Long
    Dim spec As Object, dictParam As Object
    Dim strParam As String, refStr As String
    Dim wbN As String, wsN As String, addr As String
    Dim defWb As String, defWs As String
    defWb = callerWb.Name: defWs = callerWs.Name

    For i = lb To ub
        Set spec = CreateObject("Scripting.Dictionary")

        ' Empty/Missing -> 字面量nil
        If IsEmpty(specs(i)) Or IsMissing(specs(i)) Then
            spec("type") = PARAM_LITERAL
            spec("value") = Empty
            GoTo NextSpec
        End If

        Select Case TypeName(specs(i))
            Case "Dictionary"
                Set dictParam = specs(i)
                If dictParam.Exists("isRange") And dictParam("isRange") = True Then
                    spec("type") = PARAM_REF
                    spec("level") = 0  ' 直接引用
                    spec("address") = dictParam("address")
                    spec("workbook") = dictParam("workbook")
                    spec("worksheet") = dictParam("worksheet")
                Else
                    spec("type") = PARAM_LITERAL
                    Set spec("value") = dictParam
                End If

            Case "Range"
                spec("type") = PARAM_REF
                spec("level") = 0
                spec("address") = specs(i).Address(False, False)
                spec("workbook") = specs(i).Worksheet.Parent.Name
                spec("worksheet") = specs(i).Worksheet.Name

            Case "String"
                strParam = CStr(specs(i))
                dollarCount = CountDollars(strParam)
                If dollarCount >= 1 Then
                    ' 引用类型：level = dollarCount - 1
                    ' $addr -> level=0 (直接读addr)
                    ' $$addr -> level=1 (读addr得到ptr1,读ptr1)
                    ' $$$addr -> level=2 (读addr得ptr1,读ptr1得ptr2,读ptr2)
                    spec("type") = PARAM_REF
                    spec("level") = dollarCount - 1
                    refStr = Mid(strParam, dollarCount + 1)
                    spec("refString") = refStr
                    ParseRefString refStr, wbN, wsN, addr, defWb, defWs
                    spec("workbook") = wbN
                    spec("worksheet") = wsN
                    spec("address") = addr
                Else
                    spec("type") = PARAM_LITERAL
                    spec("value") = strParam
                End If

            Case "Empty", "Nothing"
                spec("type") = PARAM_LITERAL
                spec("value") = Empty

            Case Else
                spec("type") = PARAM_LITERAL
                spec("value") = specs(i)
        End Select
NextSpec:
        Set parsed(i) = spec
    Next i

    m_ResumeSpecParsed = parsed
    Exit Sub
ErrorHandler:
    m_ResumeSpecParsed = Array()
End Sub

' 获取 Resume 时的实际参数值
Public Function GetResumeArgs() As Variant
    On Error GoTo ErrHandler

    If Not IsArray(m_ResumeSpecParsed) Then GetResumeArgs = Array(): Exit Function

    Dim lb As Long, ub As Long
    On Error Resume Next
    lb = LBound(m_ResumeSpecParsed): ub = UBound(m_ResumeSpecParsed)
    If Err.Number <> 0 Then GetResumeArgs = Array(): Exit Function
    On Error GoTo ErrHandler

    If ub < lb Then GetResumeArgs = Array(): Exit Function

    Dim result() As Variant
    ReDim result(lb To ub)

    Dim i As Long, specType As Long, level As Long
    Dim spec As Object

    For i = lb To ub
        Set spec = m_ResumeSpecParsed(i)
        specType = CLng(spec("type"))

        Select Case specType
            Case PARAM_LITERAL
                If IsObject(spec("value")) Then
                    Set result(i) = spec("value")
                Else
                    result(i) = spec("value")
                End If

            Case PARAM_REF
                level = CLng(spec("level"))
                result(i) = ResolveRef(spec("workbook"), spec("worksheet"), spec("address"), level)

            Case Else
                result(i) = Empty
        End Select
    Next i

    GetResumeArgs = result
    Exit Function
ErrHandler:
    GetResumeArgs = Array()
End Function

' ===== 内联辅助函数 =====

Private Function CountDollars(s As String) As Long
    Dim c As Long, i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) = "$" Then c = c + 1 Else Exit For
    Next
    CountDollars = c
End Function

' 解析引用字符串 [Book]Sheet!Addr
Private Sub ParseRefString(ref As String, ByRef wb As String, ByRef ws As String, _
                           ByRef addr As String, defWb As String, defWs As String)
    wb = defWb: ws = defWs: addr = ref
    Dim r As String, p As Long
    r = ref

    ' [BookName]
    If Len(r) > 0 And Left$(r, 1) = "[" Then
        p = InStr(r, "]")
        If p > 1 Then wb = Mid$(r, 2, p - 2): r = Mid$(r, p + 1)
    End If

    ' SheetName!
    p = InStr(r, "!")
    If p > 0 Then
        ws = Replace$(Left$(r, p - 1), "'", "")
        addr = Mid$(r, p + 1)
    Else
        addr = r
    End If
End Sub

' 统一的引用解析（支持多层）
' level=0: 直接读取addr的值
' level>0: 循环解引用level次
Private Function ResolveRef(wb As String, ws As String, addr As String, level As Long) As Variant
    Const MAX_DEPTH As Long = 10

    On Error GoTo ErrH

    If level > MAX_DEPTH Then ResolveRef = CVErr(xlErrRef): Exit Function

    Dim curWb As String, curWs As String, curAddr As String
    Dim content As Variant, ptr As String
    Dim lv As Long

    curWb = wb: curWs = ws: curAddr = addr

    ' 循环解引用 level 次
    For lv = 1 To level
        content = ReadCell(curWb, curWs, curAddr)
        If IsError(content) Then ResolveRef = content: Exit Function
        If IsEmpty(content) Or Len(Trim$(CStr(content))) = 0 Then
            ResolveRef = CVErr(xlErrRef): Exit Function
        End If

        ' 解析指针内容
        ptr = Trim$(CStr(content))
        ' 去掉内容中的前导$
        Do While Len(ptr) > 0 And Left$(ptr, 1) = "$"
            ptr = Mid$(ptr, 2)
        Loop
        If Len(ptr) = 0 Then ResolveRef = CVErr(xlErrRef): Exit Function

        ' 解析为下一地址
        ParseRefString ptr, curWb, curWs, curAddr, curWb, curWs
    Next lv

    ' 读取最终值
    ResolveRef = ReadCell(curWb, curWs, curAddr)
    Exit Function
ErrH:
    ResolveRef = CVErr(xlErrRef)
End Function

' 读取单元格值（内联优化版）
Private Function ReadCell(wbName As String, wsName As String, cellAddr As String) As Variant
    Dim w As Workbook, s As Worksheet, r As Range
    On Error Resume Next
    Set w = Application.Workbooks(wbName)
    If w Is Nothing Then ReadCell = CVErr(xlErrRef): Exit Function
    Set s = w.Sheets(wsName)
    If s Is Nothing Then ReadCell = CVErr(xlErrRef): Exit Function
    Set r = s.Range(cellAddr)
    If r Is Nothing Then ReadCell = CVErr(xlErrRef): Exit Function
    On Error GoTo 0
    ReadCell = r.Value
End Function
