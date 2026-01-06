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

' 解析后的参数规格
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

' ============================================
' 统一参数规格解析（优化版）
' ============================================
' 输入格式        | type | level | 说明
' ----------------|------|-------|---------------------------
' Range对象/Dict  |  1   |   0   | 直接读取指定单元格/区域的值
' "$B1"           |  1   |   0   | 直接读取B1的值
' "$B1:C2"        |  1   |   0   | 直接读取B1:C2区域的值
' "$$B1"          |  1   |   1   | 读B1得到地址ptr1，再读ptr1的值
' "$$$B1"         |  1   |   2   | 读B1得ptr1，读ptr1得ptr2，读ptr2的值
' 其他            |  0   |   -   | 字面量，直接传递
Public Sub ParseResumeSpecs(specs As Variant, callerWb As Workbook, callerWs As Worksheet)
    On Error GoTo ErrorHandler

    ' 安全获取数组边界
    Dim lb As Long, ub As Long
    On Error Resume Next
    lb = LBound(specs): ub = UBound(specs)
    If Err.Number <> 0 Then
        Err.Clear
        m_ResumeSpecParsed = Array()
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    If ub < lb Then
        m_ResumeSpecParsed = Array()
        Exit Sub
    End If

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
            Set parsed(i) = spec
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
                    spec("refString") = "[" & dictParam("workbook") & "]" & dictParam("worksheet") & "!" & dictParam("address")
                Else
                    spec("type") = PARAM_LITERAL
                    Set spec("value") = dictParam
                End If

            Case "String"
                strParam = CStr(specs(i))
                dollarCount = CountLeadingDollars(strParam)

                If dollarCount >= 1 Then
                    ' 引用类型：level = dollarCount - 1
                    ' $addr -> level=0 (直接读addr)
                    ' $$addr -> level=1 (读addr得到ptr1,读ptr1)
                    ' $$$addr -> level=2 (读addr得ptr1,读ptr1得ptr2,读ptr2)

                    ' 检查空引用
                    If Len(Trim$(refStr)) = 0 Then
                        ' 当作字面量处理
                        spec("type") = PARAM_LITERAL
                        spec("value") = strParam
                        Set parsed(i) = spec
                        GoTo NextSpec
                    End If

                    spec("type") = PARAM_REF
                    spec("level") = dollarCount - 1
                    refStr = Mid$(strParam, dollarCount + 1)
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

        Set parsed(i) = spec
NextSpec:
    Next i

    m_ResumeSpecParsed = parsed
    Exit Sub

ErrorHandler:
    m_ResumeSpecParsed = Array()
End Sub

' 获取 Resume 时的实际参数值
Public Function GetResumeArgs() As Variant
    On Error GoTo ErrHandler

    If Not IsArray(m_ResumeSpecParsed) Then
        GetResumeArgs = Array()
        Exit Function
    End If

    Dim lb As Long, ub As Long
    On Error Resume Next
    lb = LBound(m_ResumeSpecParsed): ub = UBound(m_ResumeSpecParsed)
    If Err.Number <> 0 Then
        Err.Clear
        GetResumeArgs = Array()
        Exit Function
    End If
    On Error GoTo ErrHandler

    If ub < lb Then
        GetResumeArgs = Array()
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(lb To ub)

    Dim i As Long, specType As Long, level As Long
    Dim spec As Object

    For i = lb To ub
        Set spec = m_ResumeSpecParsed(i)
        specType = CLng(spec("type"))

        Select Case specType
            Case PARAM_LITERAL
                If spec.Exists("value") Then
                    If IsObject(spec("value")) Then
                        Set result(i) = spec("value")
                    Else
                        result(i) = spec("value")
                    End If
                Else
                    result(i) = Empty
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

' 计算字符串开头的$符号数量
Private Function CountLeadingDollars(s As String) As Long
    Dim c As Long, i As Long, sLen As Long
    sLen = Len(s)
    c = 0
    For i = 1 To sLen
        If Mid$(s, i, 1) = "$" Then
            c = c + 1
        Else
            Exit For
        End If
    Next
    CountLeadingDollars = c
End Function

' 解析引用字符串 [Book]Sheet!Addr
' 支持格式：
'   Addr                    -> 使用默认工作簿和工作表
'   Sheet!Addr              -> 使用默认工作簿
'   [Book]Sheet!Addr        -> 完整引用
'   'Sheet Name'!Addr       -> 带空格的工作表名
Private Sub ParseRefString(ref As String, ByRef wb As String, ByRef ws As String, _
                           ByRef addr As String, defWb As String, defWs As String)
    wb = defWb: ws = defWs: addr = ref
    If Len(ref) = 0 Then Exit Sub
    Dim r As String, p As Long, bracketEnd As Long
    r = ref
    ' [BookName] 部分
    If Left$(r, 1) = "[" Then
        bracketEnd = InStr(r, "]")
        If bracketEnd > 1 Then
            wb = Mid$(r, 2, bracketEnd - 2)
            r = Mid$(r, bracketEnd + 1)
        End If
    End If

    ' SheetName! 部分
    p = InStr(r, "!")
    If p > 0 Then
        ws = Left$(r, p - 1)
        ' 移除工作表名两端的单引号
        If Len(ws) >= 2 Then
            If Left$(ws, 1) = "'" And Right$(ws, 1) = "'" Then
                ws = Mid$(ws, 2, Len(ws) - 2)
            End If
        End If
        addr = Mid$(r, p + 1)
    Else
        addr = r
    End If
End Sub

' 统一的引用解析（支持多层间接引用）
' level=0: 直接读取addr的值
' level>0: 循环解引用level次
Private Function ResolveRef(wb As String, ws As String, addr As String, level As Long) As Variant
    Const MAX_DEPTH As Long = 10
    On Error GoTo ErrH

    ' 防止无限循环
    If level > MAX_DEPTH Then
        ResolveRef = CVErr(xlErrRef)
        Exit Function
    End If
    ' 验证初始参数
    If Len(wb) = 0 Or Len(ws) = 0 Or Len(addr) = 0 Then
        ResolveRef = CVErr(xlErrRef)
        Exit Function
    End If

    Dim curWb As String, curWs As String, curAddr As String
    Dim content As Variant, ptr As String
    Dim lv As Long

    curWb = wb: curWs = ws: curAddr = addr

    ' 循环解引用 level 次
    For lv = 1 To level
        content = ReadCellValue(curWb, curWs, curAddr)
        ' 检查读取结果
        If IsError(content) Then
            ResolveRef = content
            Exit Function
        End If
        If IsEmpty(content) Then
            ResolveRef = CVErr(xlErrRef)
            Exit Function
        End If
        ' 转换为字符串并检查
        ptr = Trim$(CStr(content))
        If Len(ptr) = 0 Then
            ResolveRef = CVErr(xlErrRef)
            Exit Function
        End If
        ' 去掉内容中的前导$符号
        Do While Len(ptr) > 0 And Left$(ptr, 1) = "$"
            ptr = Mid$(ptr, 2)
        Loop

        If Len(ptr) = 0 Then
            ResolveRef = CVErr(xlErrRef)
            Exit Function
        End If

        ' 解析为下一地址（使用当前工作簿/工作表作为默认值）
        ParseRefString ptr, curWb, curWs, curAddr, curWb, curWs
    Next lv

    ' 读取最终值
    ResolveRef = ReadCellValue(curWb, curWs, curAddr)
    Exit Function
ErrH:
    ResolveRef = CVErr(xlErrRef)
End Function

' 读取单元格值（优化版，支持区域）
Private Function ReadCellValue(wbName As String, wsName As String, cellAddr As String) As Variant
    Dim w As Workbook, s As Worksheet, r As Range

    On Error Resume Next

    Set w = Application.Workbooks(wbName)
    If w Is Nothing Then
        ReadCellValue = CVErr(xlErrRef)
        Exit Function
    End If

    Set s = w.Sheets(wsName)
    If s Is Nothing Then
        ReadCellValue = CVErr(xlErrRef)
        Exit Function
    End If

    Set r = s.Range(cellAddr)
    If r Is Nothing Then
        ReadCellValue = CVErr(xlErrRef)
        Exit Function
    End If

    On Error GoTo 0

    ' 单个单元格返回值，多个单元格返回数组
    If r.Cells.Count = 1 Then
        ReadCellValue = r.Value
    Else
        ReadCellValue = r.Value  ' 返回二维数组
    End If
End Function
