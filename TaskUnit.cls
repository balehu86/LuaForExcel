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
    Dim i As Long
    For i = 1 To Me.taskWatches.Count
        If Me.taskWatches(i) = watchCell Then Exit Sub
    Next i
    Me.taskWatches.Add watchCell
End Sub

Public Sub RemoveWatch(watchCell As String)
    Dim i As Long
    For i = Me.taskWatches.Count To 1 Step -1
        If Me.taskWatches(i) = watchCell Then
            Me.taskWatches.Remove i
            Exit Sub
        End If
    Next i
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
' "$$$B1"         |  1   |   2   | 读B1得ptr1，读ptr1得ptr2，读ptr2的值
' 其他            |  0   |   -   | 字面量，直接传递
Public Sub ParseResumeSpecs(specs As Variant, callerWb As Workbook, callerWs As Worksheet)
    On Error GoTo ErrorHandler

    ' 快速检查空数组
    If Not TryGetArrayBounds(specs, lb, ub) Then
        m_ResumeSpecParsed = Array()
        Exit Sub
    End If

    Dim lb As Long, ub As Long
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

    Select Case TypeName(spec)
        Case "Dictionary"
            ParseDictionarySpec spec, result

        Case "String"
            ParseStringSpec CStr(spec), result, defWb, defWs

        Case Else
            result("type") = PARAM_LITERAL
            If IsObject(spec) Then
                Set result("value") = spec
            Else
                result("value") = spec
            End If
    End Select

    Set ParseSingleSpec = result
End Function

' 解析 Dictionary 类型参数
Private Sub ParseDictionarySpec(dictParam As Object, result As Object)
    If dictParam.Exists("isRange") And dictParam("isRange") = True Then
        result("type") = PARAM_REF
        result("level") = 0
        result("address") = dictParam("address")
        result("workbook") = dictParam("workbook")
        result("worksheet") = dictParam("worksheet")
        result("refString") = "[" & dictParam("workbook") & "]" & dictParam("worksheet") & "!" & dictParam("address")
    Else
        result("type") = PARAM_LITERAL
        Set result("value") = dictParam
    End If
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
        Set spec = m_ResumeSpecParsed(i)
        result(i) = ResolveSpecValue(spec)
    Next i

    GetResumeArgs = result
    Exit Function
ErrHandler:
    GetResumeArgs = Array()
End Function

' 解析单个规格的值
Private Function ResolveSpecValue(spec As Object) As Variant
    Select Case CLng(spec("type"))
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
            ResolveSpecValue = ResolveRef(spec("workbook"), spec("worksheet"), spec("address"), CLng(spec("level")))

        Case Else
            ResolveSpecValue = Empty
    End Select
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
Private Function ResolveRef(wb As String, ws As String, addr As String, level As Long) As Variant
    Const MAX_DEPTH As Long = 10
    On Error GoTo ErrH
    If level > MAX_DEPTH Or Len(wb) = 0 Or Len(ws) = 0 Or Len(addr) = 0 Then
        ResolveRef = CVErr(xlErrRef)
        Exit Function
    End If

    Dim curWb As String, curWs As String, curAddr As String
    Dim content As Variant, ptr As String

    curWb = wb: curWs = ws: curAddr = addr

    ' 循环解引用
    Dim lv As Long
    For lv = 1 To level
        content = ReadCellValue(curWb, curWs, curAddr)
        If IsError(content) Or IsEmpty(content) Then
            ResolveRef = CVErr(xlErrRef)
            Exit Function
        End If

        ptr = Trim$(CStr(content))

        ' 去掉前导$符号
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
    Exit Function
ErrH:
    ResolveRef = CVErr(xlErrRef)
End Function

' 读取单元格值
Private Function ReadCellValue(wbName As String, wsName As String, cellAddr As String) As Variant
    On Error GoTo ErrH
    Dim w As Workbook, s As Worksheet, r As Range
    Set w = Application.Workbooks(wbName)
    Set s = w.Sheets(wsName)
    Set r = s.Range(cellAddr)

    ReadCellValue = r.Value
    Exit Function
ErrH:
    ReadCellValue = CVErr(xlErrRef)
End Function
