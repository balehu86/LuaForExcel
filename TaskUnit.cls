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

Private m_ResumeSpecParsed As Variant  ' 

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
    Dim alreadyExists As Boolean
    alreadyExists = False

    Dim i As Long
    On Error Resume Next
    For i = 1 To Me.taskWatches.Count
        If Me.taskWatches(i) = watchCell Then
            alreadyExists = True
            Exit For
        End If
    Next i
    On Error GoTo 0

    If Not alreadyExists Then
        Me.taskWatches.Add watchCell
    End If
End Sub

' 移除监控单元格
Public Sub RemoveWatch(watchCell As String)
    On Error Resume Next
    Dim i As Long
    For i = Me.taskWatches.Count To 1 Step -1
        If Me.taskWatches(i) = watchCell Then
            Me.taskWatches.Remove i
            Exit For
        End If
    Next i
    On Error GoTo 0
End Sub

' 获取解析后的参数规格（只读属性）
Public Property Get ResumeSpecParsed() As Variant
    ResumeSpecParsed = m_ResumeSpecParsed
End Property

' 简化后的参数规格解析
' 参数类型：
'   PARAM_LITERAL (0)        - 字面量（数值、布尔、普通字符串）
'   PARAM_RANGE_REF (1)      - 单元格/区域引用
'   PARAM_DYNAMIC_STRING (2) - 动态字符串（"$B1" 格式）
Public Sub ParseResumeSpecs(specs As Variant, callerWb As Workbook, callerWs As Worksheet)
    On Error GoTo ErrorHandler

    ' 安全获取数组边界
    Dim lowerBound As Long, upperBound As Long
    On Error Resume Next
    lowerBound = LBound(specs)
    upperBound = UBound(specs)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrorHandler
        m_ResumeSpecParsed = Array()
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    If upperBound < lowerBound Then
        m_ResumeSpecParsed = Array()
        Exit Sub
    End If

    Dim parsedSpecs() As Variant
    ReDim parsedSpecs(lowerBound To upperBound)

    Dim i As Long
    For i = lowerBound To upperBound
        Dim spec As Object
        Set spec = CreateObject("Scripting.Dictionary")

        Select Case TypeName(specs(i))
            Case "Dictionary"
                ' Dictionary 类型：检查是否是 Range 信息
                Dim dictParam As Object
                Set dictParam = specs(i)

                If dictParam.Exists("isRange") And dictParam("isRange") = True Then
                    ' Range 引用信息
                    spec("type") = PARAM_RANGE_REF
                    spec("address") = dictParam("address")
                    spec("workbook") = dictParam("workbook")
                    spec("worksheet") = dictParam("worksheet")
                Else
                    ' 普通 Dictionary，作为字面量
                    spec("type") = PARAM_LITERAL
                    Set spec("value") = dictParam
                End If

            Case "Range"
                ' Range 对象（理论上不会出现，因为 LuaTask 已转换为 Dictionary）
                Dim rangeParam As Range
                Set rangeParam = specs(i)
                spec("type") = PARAM_RANGE_REF
                spec("address") = rangeParam.Address(False, False)
                spec("workbook") = rangeParam.Worksheet.Parent.Name
                spec("worksheet") = rangeParam.Worksheet.Name

            Case "String"
                ' 字符串：检查是否是动态引用
                Dim strParam As String
                strParam = CStr(specs(i))

                If Len(strParam) > 0 And Left(strParam, 1) = "$" Then
                    ' 动态引用格式 "$B1"
                    spec("type") = PARAM_DYNAMIC_STRING
                    Dim refString As String
                    refString = Mid(strParam, 2)
                    spec("refString") = refString

                    ' 解析并存储完整引用信息
                    Dim wbName As String, wsName As String, addr As String
                    ParseReferenceString refString, wbName, wsName, addr, _
                                         callerWb.Name, callerWs.Name
                    spec("workbook") = wbName
                    spec("worksheet") = wsName
                    spec("address") = addr
                Else
                    ' 普通字符串
                    spec("type") = PARAM_LITERAL
                    spec("value") = strParam
                End If

            Case Else
                ' 其他类型（数值、布尔等）作为字面量
                spec("type") = PARAM_LITERAL
                spec("value") = specs(i)
        End Select

        Set parsedSpecs(i) = spec
    Next i

    m_ResumeSpecParsed = parsedSpecs
    Exit Sub

ErrorHandler:
    Debug.Print "TaskUnit.ParseResumeSpecs Error: " & Err.Description
    m_ResumeSpecParsed = Array()
End Sub

' 获取 Resume 时的实际参数值
Public Function GetResumeArgs() As Variant
    On Error GoTo ErrorHandler

    If Not IsArray(m_ResumeSpecParsed) Then
        GetResumeArgs = Array()
        Exit Function
    End If

    Dim lowerBound As Long, upperBound As Long
    On Error Resume Next
    lowerBound = LBound(m_ResumeSpecParsed)
    upperBound = UBound(m_ResumeSpecParsed)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrorHandler
        GetResumeArgs = Array()
        Exit Function
    End If
    On Error GoTo ErrorHandler

    If upperBound < lowerBound Then
        GetResumeArgs = Array()
        Exit Function
    End If

    Dim resultArgs() As Variant
    ReDim resultArgs(lowerBound To upperBound)

    Dim i As Long
    For i = lowerBound To upperBound
        Dim spec As Object
        Set spec = m_ResumeSpecParsed(i)

        Dim specType As Long
        specType = CLng(spec("type"))

        Select Case specType
            Case PARAM_LITERAL
                If IsObject(spec("value")) Then
                    Set resultArgs(i) = spec("value")
                Else
                    resultArgs(i) = spec("value")
                End If

            Case PARAM_RANGE_REF, PARAM_DYNAMIC_STRING
                ' 两种引用类型处理相同：读取当前值
                resultArgs(i) = ReadRangeValue( _
                    CStr(spec("workbook")), _
                    CStr(spec("worksheet")), _
                    CStr(spec("address")) _
                )

            Case Else
                resultArgs(i) = Empty
        End Select
    Next i

    GetResumeArgs = resultArgs
    Exit Function

ErrorHandler:
    Debug.Print "TaskUnit.GetResumeArgs Error: " & Err.Description
    GetResumeArgs = Array()
End Function

' ===== 私有辅助函数 =====

Private Function ReadRangeValue(wbName As String, wsName As String, cellAddress As String) As Variant
    On Error GoTo ErrorHandler

    Dim targetWb As Workbook
    On Error Resume Next
    Set targetWb = Application.Workbooks(wbName)
    On Error GoTo ErrorHandler

    If targetWb Is Nothing Then
        ReadRangeValue = CVErr(xlErrRef)
        Exit Function
    End If

    Dim targetWs As Worksheet
    On Error Resume Next
    Set targetWs = targetWb.Sheets(wsName)
    On Error GoTo ErrorHandler

    If targetWs Is Nothing Then
        ReadRangeValue = CVErr(xlErrRef)
        Exit Function
    End If

    Dim targetRange As Range
    On Error Resume Next
    Set targetRange = targetWs.Range(cellAddress)
    On Error GoTo ErrorHandler

    If targetRange Is Nothing Then
        ReadRangeValue = CVErr(xlErrRef)
        Exit Function
    End If

    ReadRangeValue = targetRange.Value
    Exit Function

ErrorHandler:
    ReadRangeValue = CVErr(xlErrRef)
End Function

Private Sub ParseReferenceString(refString As String, _
                                  ByRef outWbName As String, _
                                  ByRef outWsName As String, _
                                  ByRef outCellAddress As String, _
                                  defaultWbName As String, _
                                  defaultWsName As String)
    outWbName = defaultWbName
    outWsName = defaultWsName
    outCellAddress = refString

    Dim remaining As String
    remaining = refString

    ' 解析工作簿 [BookName]
    If Len(remaining) > 0 And Left(remaining, 1) = "[" Then
        Dim bracketEnd As Long
        bracketEnd = InStr(remaining, "]")
        If bracketEnd > 1 Then
            outWbName = Mid(remaining, 2, bracketEnd - 2)
            remaining = Mid(remaining, bracketEnd + 1)
        End If
    End If

    ' 解析工作表 SheetName!
    Dim exclamPos As Long
    exclamPos = InStr(remaining, "!")
    If exclamPos > 0 Then
        outWsName = Replace(Left(remaining, exclamPos - 1), "'", "")
        outCellAddress = Mid(remaining, exclamPos + 1)
    Else
        outCellAddress = remaining
    End If
End Sub
