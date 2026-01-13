' TaskUnit.cls
Option Explicit

' task的基本信息
Public taskWorkbook As String    ' taskId -> workbookInfo instance
Public taskStatus As CoStatus    ' taskId -> status string:defined、yield、paused、done、error、terminated
Public taskCoRef As Long         ' Lua 注册表引用号（防止 GC）
Public taskId As Integer
Public taskFunc As String
Public taskStartArgs As Variant
Public taskResumeSpec As Variant
Public taskCell As String
Public taskProgress As Double
Public taskMessage As String
Public taskValue As Variant
Public taskError As String
Public taskCoThread As LongPtr

' task的调度统计
Public taskLastTime As Double
Public taskTotalTime As Double
Public taskTickCount As Long

' CFS 调度字段
Public CFS_vruntime As Double
Public CFS_weight As Double
Public CFS_lastScheduled As Double

' task 的监视器集合
Public taskWatches As Collection

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

Public Sub ClearCoroutineRef()
    taskCoRef = 0
    taskCoThread = 0
End Sub

Public Sub AddWatch(watchCell As String)
    On Error Resume Next
    ' 尝试直接添加，如果重复会静默失败
    Dim i As Long
    For i = 1 To taskWatches.Count
        If taskWatches(i) = watchCell Then Exit Sub
    Next
    taskWatches.Add watchCell
End Sub

Public Sub RemoveWatch(watchCell As String)
    Dim i As Long
    For i = 1 To taskWatches.Count
        If taskWatches(i) = watchCell Then
            taskWatches.Remove i
            Exit Sub
        End If
    Next
End Sub

Public Property Get ResumeSpecParsed() As Variant
    ResumeSpecParsed = m_ResumeSpecParsed
End Property

' ============================================
' 统一参数规格解析（简化版）
' ============================================
' 输入格式        | type | level | 说明
' ----------------|------|-------|---------------------------
' Range对象/Dict  |  1   |   0   | 直接读取指定单元格/区域的值
' "$B1"           |  1   |   0   | 直接读取B1的值
' "$B1:C2"        |  1   |   0   | 直接读取B1:C2区域的值
' "$$B1"          |  1   |   1   | 读B1得到地址ptr1，再读ptr1的值
' 其他            |  0   |   -   | 字面量，直接传递
Public Sub ParseResumeSpecs(specs As Variant, callerWb As Workbook, callerWs As Worksheet)
    On Error GoTo ErrorHandler
    Dim lb As Long, ub As Long

    ' 快速检查空数组
    If Not TryGetArrayBounds(specs, lb, ub) Then
        m_ResumeSpecParsed = Array()
        Exit Sub
    End If

    Dim parsed() As Variant
    ReDim parsed(lb To ub)

    Dim i As Long
    Dim defWb As String, defWs As String
    defWb = callerWb.Name: defWs = callerWs.Name

    For i = lb To ub
        Set parsed(i) = ParseSingleSpec(specs(i), defWb, defWs)
    Next i

    m_ResumeSpecParsed = parsed
    Exit Sub

ErrorHandler:
    m_ResumeSpecParsed = Array()
End Sub

' 解析单个参数规格
Private Function ParseSingleSpec(spec As Variant, defWb As String, defWs As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    ' Empty/Missing -> 字面量nil
    If IsEmpty(spec) Or IsMissing(spec) Then
        result("type") = PARAM_LITERAL
        result("value") = Empty
        Set ParseSingleSpec = result
        Exit Function
    End If

    ' 获取类型名称
    Dim specTypeName As String
    specTypeName = TypeName(spec)

    Select Case specTypeName
        Case "Dictionary"
            ' 需要显式转换为 Object
            Dim dictSpec As Object
            Set dictSpec = spec
            ParseDictionarySpec dictSpec, result

        Case "String"
            ParseStringSpec CStr(spec), result, defWb, defWs

        Case "Range"
            ' 处理 Range 对象（如果直接传入）
            Dim rngSpec As Object
            Set rngSpec = spec
            result("type") = PARAM_REF
            result("level") = 0
            result("address") = rngSpec.Address(False, False)
            result("workbook") = rngSpec.Worksheet.Parent.Name
            result("worksheet") = rngSpec.Worksheet.Name
            result("refString") = "[" & rngSpec.Worksheet.Parent.Name & "]" & _
                                  rngSpec.Worksheet.Name & "!" & rngSpec.Address(False, False)

        Case Else
            ' 其他类型作为字面量
            result("type") = PARAM_LITERAL
            result("value") = spec  ' VBA 会自动处理值类型赋值

    End Select

    Set ParseSingleSpec = result
End Function

' 解析 Dictionary 类型参数
Private Sub ParseDictionarySpec(dictParam As Object, result As Object)
    If dictParam.Exists("isRange") Then
        If dictParam("isRange") = True Then
            result("type") = PARAM_REF
            result("level") = 0
            result("address") = dictParam("address")
            result("workbook") = dictParam("workbook")
            result("worksheet") = dictParam("worksheet")
            result("refString") = "[" & dictParam("workbook") & "]" & _
                                  dictParam("worksheet") & "!" & dictParam("address")
            Exit Sub
        End If
    End If

    ' 非 Range 的 Dictionary，作为字面量
    result("type") = PARAM_LITERAL
    Set result("value") = dictParam
End Sub

' 解析字符串类型参数
Private Sub ParseStringSpec(strParam As String, result As Object, defWb As String, defWs As String)
    Dim dollarCount As Long
    dollarCount = CountLeadingDollars(strParam)

    If dollarCount = 0 Then
        ' 普通字符串字面量
        result("type") = PARAM_LITERAL
        result("value") = strParam
        Exit Sub
    End If

    ' 引用类型
    Dim refStr As String
    refStr = Trim$(Mid$(strParam, dollarCount + 1))

    ' 空引用 -> 字面量
    If Len(refStr) = 0 Then
        result("type") = PARAM_LITERAL
        result("value") = strParam
        Exit Sub
    End If

    Dim wbN As String, wsN As String, addr As String
    ParseRefString refStr, wbN, wsN, addr, defWb, defWs

    result("type") = PARAM_REF
    result("level") = dollarCount - 1
    result("refString") = refStr
    result("workbook") = wbN
    result("worksheet") = wsN
    result("address") = addr
End Sub

' 获取 Resume 时的实际参数值
Public Function GetResumeArgs() As Variant
    On Error GoTo ErrHandler

    Dim lb As Long, ub As Long
    If Not TryGetArrayBounds(m_ResumeSpecParsed, lb, ub) Then
        GetResumeArgs = Array()
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(lb To ub)

    Dim i As Long
    Dim spec As Object

    For i = lb To ub
        ' 显式转换为 Object
        If IsObject(m_ResumeSpecParsed(i)) Then
            Set spec = m_ResumeSpecParsed(i)
        Else
            ' 如果不是对象，跳过或处理错误
            result(i) = Empty
            GoTo NextSpec
        End If

        result(i) = ResolveSpecValue(spec)
NextSpec:
    Next i
    GetResumeArgs = result
    Exit Function
ErrHandler:
    GetResumeArgs = Array()
End Function

' 解析单个规格的值
Private Function ResolveSpecValue(spec As Object) As Variant
    On Error GoTo ErrH

    Dim specType As Long
    specType = CLng(spec("type"))

    Select Case specType
        Case PARAM_LITERAL
            If spec.Exists("value") Then
                If IsObject(spec("value")) Then
                    Set ResolveSpecValue = spec("value")
                Else
                    ResolveSpecValue = spec("value")
                End If
            Else
                ResolveSpecValue = Empty
            End If

        Case PARAM_REF
            Dim wb As String, ws As String, addr As String, lvl As Long
            wb = spec("workbook")
            ws = spec("worksheet")
            addr = spec("address")
            lvl = CLng(spec("level"))
            ResolveSpecValue = ResolveRef(wb, ws, addr, lvl)

        Case Else
            ResolveSpecValue = Empty
    End Select
    Exit Function

ErrH:
    ResolveSpecValue = CVErr(xlErrValue)
End Function

' ===== 辅助函数 =====

' 尝试获取数组边界（统一的空数组检查）
Private Function TryGetArrayBounds(arr As Variant, ByRef lb As Long, ByRef ub As Long) As Boolean
    On Error Resume Next
    lb = LBound(arr)
    ub = UBound(arr)
    TryGetArrayBounds = (Err.Number = 0) And (ub >= lb)
    Err.Clear
End Function

' 计算字符串开头的$符号数量
Private Function CountLeadingDollars(s As String) As Long
    Dim c As Long, sLen As Long
    sLen = Len(s)
    For c = 1 To sLen
        If Mid$(s, c, 1) <> "$" Then Exit For
    Next
    CountLeadingDollars = c - 1
End Function

' 解析引用字符串 [Book]Sheet!Addr
Private Sub ParseRefString(ref As String, ByRef wb As String, ByRef ws As String, _
                           ByRef addr As String, defWb As String, defWs As String)
    wb = defWb: ws = defWs: addr = ref
    If Len(ref) = 0 Then Exit Sub

    Dim r As String, p As Long
    r = ref
    ' [BookName] 部分
    If Left$(r, 1) = "[" Then
        p = InStr(r, "]")
        If p > 1 Then
            wb = Mid$(r, 2, p - 2)
            r = Mid$(r, p + 1)
        End If
    End If

    ' SheetName! 部分
    p = InStr(r, "!")
    If p > 0 Then
        ws = Left$(r, p - 1)
        If Len(ws) >= 2 And Left$(ws, 1) = "'" And Right$(ws, 1) = "'" Then
            ws = Mid$(ws, 2, Len(ws) - 2)
        End If
        addr = Mid$(r, p + 1)
    Else
        addr = r
    End If
End Sub

' 解析引用（支持多层间接引用）
' level=0: 直接读取 addr 的值（$A1）
' level=1: 读取 addr 得到新地址，再读取新地址的值（$$A1）
Private Function ResolveRef(wb As String, ws As String, addr As String, level As Long) As Variant
    Const MAX_DEPTH As Long = 10

    ' 合并所有前置检查
    If level < 0 Or level > MAX_DEPTH Or Len(wb) = 0 Or Len(ws) = 0 Or Len(addr) = 0 Then
        ResolveRef = CVErr(xlErrRef)
        Exit Function
    End If

    Dim curWb As String: curWb = wb
    Dim curWs As String: curWs = ws
    Dim curAddr As String: curAddr = addr
    Dim content As Variant
    Dim ptr As String

    Dim lv As Long
    For lv = 1 To level
        content = ReadCellValue(curWb, curWs, curAddr)

        ' 合并错误检查
        If IsError(content) Or IsEmpty(content) Or IsNull(content) Then
            ResolveRef = CVErr(xlErrRef)
            Exit Function
        End If

        ptr = Trim$(CStr(content))

        ' 内联去除前导$符号
        Do While Len(ptr) > 0 And Left$(ptr, 1) = "$"
            ptr = Mid$(ptr, 2)
        Loop

        If Len(ptr) = 0 Then
            ResolveRef = CVErr(xlErrRef)
            Exit Function
        End If

        ParseRefString ptr, curWb, curWs, curAddr, curWb, curWs
    Next lv

    ResolveRef = ReadCellValue(curWb, curWs, curAddr)
End Function

' 读取单元格值（支持溢出范围引用 #）
Private Function ReadCellValue(wbName As String, wsName As String, cellAddr As String) As Variant
    On Error GoTo ErrH

    Dim rng As Range
    Dim ws As Worksheet
    Dim actualAddr As String

    Set ws = Application.Workbooks(wbName).Sheets(wsName)
    actualAddr = cellAddr

    ' 检查是否为溢出范围引用（以#结尾）
    If Right$(actualAddr, 1) = "#" Then
        ' 移除#符号获取基础地址
        Dim baseAddr As String
        baseAddr = Left$(actualAddr, Len(actualAddr) - 1)

        ' 获取基础单元格
        Set rng = ws.Range(baseAddr)

        ' 尝试获取溢出范围（SpillRange）
        On Error Resume Next
        Dim spillRng As Range
        Set spillRng = rng.SpillingToRange

        If Err.Number <> 0 Or spillRng Is Nothing Then
            Err.Clear
            ' 如果没有溢出范围，尝试使用 SpillRange 属性（Excel 365/2021+）
            Set spillRng = rng.SpillRange
        End If

        If Err.Number <> 0 Or spillRng Is Nothing Then
            Err.Clear
            On Error GoTo ErrH
            ' 没有溢出范围，返回单个单元格的值
            ReadCellValue = rng.Value
            Exit Function
        End If
        On Error GoTo ErrH

        ' 返回整个溢出范围的值
        ReadCellValue = spillRng.Value
    Else
        ' 普通引用
        ReadCellValue = ws.Range(actualAddr).Value
    End If

    Exit Function
ErrH:
    ReadCellValue = CVErr(xlErrRef)
End Function
