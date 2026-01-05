' TaskUnit.cls
Option Explicit

Public taskId As Integer         ' taskId int
Public taskFunc As String        ' taskId -> func name
Public taskWorkbook As String    ' taskId -> workbookInfo instance
Public taskStartArgs As Variant  ' taskId -> startArgs array
Public taskResumeSpec As Variant ' taskId -> resumeSpec array
Public taskCell As String        ' taskId -> taskCell address
Public taskStatus As CoStatus ' taskId -> status string:defined、yield、paused、done、error、terminated
Public taskProgress As Double    ' taskId -> progress number
Public taskMessage As String     ' taskId -> message variant
Public taskValue As Variant      ' taskId -> value variant
Public taskError As String       ' taskId -> error message
Public taskCoThread As LongPtr   ' taskId -> coThread LongPtr
Public taskCoRef As Long         ' Lua 注册表引用号（防止 GC）

Public taskLastTime As Double    ' 任务上次运行时间(ms)
Public taskTotalTime As Double   ' 任务总运行时间(ms)
Public taskTickCount As Long     ' 任务调度次数
' === CFS 调度字段 ===
Public CFS_vruntime As Double      ' 虚拟运行时间（核心调度依据）
Public CFS_weight As Double        ' 任务权重（默认1024，越大优先级越高）
Public CFS_lastScheduled As Double ' 上次被调度的时间戳（GetTickCount）

Public taskWatches As Collection ' 任务监视的Watchcell字符串的集合

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
    Dim exists As Boolean
    exists = False
    
    On Error Resume Next
    Dim i As Long
    For i = 1 To Me.taskWatches.Count
        If Me.taskWatches(i) = watchCell Then
            exists = True
            Exit For
        End If
    Next
    On Error GoTo 0
    
    If Not exists Then
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
    Next
End Sub

' 获取解析后的参数规格（只读属性）
Public Property Get ResumeSpecParsed() As Variant
    ResumeSpecParsed = m_ResumeSpecParsed
End Property

' 解析 Resume 参数规格
Public Sub ParseResumeSpecs(specs As Variant, wb As Workbook, ws As Worksheet)
    On Error GoTo ErrorHandler
    
    If Not IsArray(specs) Then
        m_ResumeSpecParsed = Array()
        Exit Sub
    End If
    
    Dim lb As Long, ub As Long
    On Error Resume Next
    lb = LBound(specs)
    ub = UBound(specs)
    If Err.Number <> 0 Then
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
    
    Dim i As Long
    For i = lb To ub
        Dim spec As Object
        Set spec = CreateObject("Scripting.Dictionary")
        
        ' 检查参数类型
        If TypeName(specs(i)) = "Dictionary" Then
            Dim dictParam As Object
            Set dictParam = specs(i)
            
            If dictParam.Exists("isRange") And dictParam("isRange") = True Then
                If dictParam("cellCount") = 1 Then
                    spec("type") = PARAM_CELL_REF
                Else
                    spec("type") = PARAM_RANGE_REF
                End If
                spec("address") = dictParam("address")
                spec("workbook") = dictParam("workbook")
                spec("worksheet") = dictParam("worksheet")
            Else
                spec("type") = PARAM_LITERAL
                Set spec("value") = dictParam
            End If
            
        ElseIf TypeName(specs(i)) = "Range" Then
            Dim rng As Range
            Set rng = specs(i)
            
            If rng.Cells.Count = 1 Then
                spec("type") = PARAM_CELL_REF
            Else
                spec("type") = PARAM_RANGE_REF
            End If
            
            spec("address") = rng.Address(False, False)
            spec("workbook") = rng.Worksheet.Parent.Name
            spec("worksheet") = rng.Worksheet.Name
            
        ElseIf VarType(specs(i)) = vbString Then
            Dim paramStr As String
            paramStr = CStr(specs(i))
            
            If Len(paramStr) > 0 And Left(paramStr, 1) = "$" Then
                spec("type") = PARAM_DYNAMIC_STRING
                spec("refString") = Mid(paramStr, 2)
                spec("workbook") = wb.Name
                spec("worksheet") = ws.Name
            Else
                spec("type") = PARAM_LITERAL
                spec("value") = paramStr
            End If
            
        Else
            spec("type") = PARAM_LITERAL
            spec("value") = specs(i)
        End If
        
        Set parsed(i) = spec
    Next i
    
    m_ResumeSpecParsed = parsed
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
    
    Dim lb As Long, ub As Long
    On Error Resume Next
    lb = LBound(m_ResumeSpecParsed)
    ub = UBound(m_ResumeSpecParsed)
    If Err.Number <> 0 Then
        GetResumeArgs = Array()
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    If ub < lb Then
        GetResumeArgs = Array()
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(lb To ub)
    
    Dim i As Long
    For i = lb To ub
        Dim spec As Object
        Set spec = m_ResumeSpecParsed(i)
        
        Dim specType As Long
        specType = CLng(spec("type"))
        
        Select Case specType
            Case PARAM_LITERAL
                If IsObject(spec("value")) Then
                    Set result(i) = spec("value")
                Else
                    result(i) = spec("value")
                End If
                
            Case PARAM_CELL_REF, PARAM_RANGE_REF
                result(i) = ReadRangeValue(CStr(spec("workbook")), CStr(spec("worksheet")), CStr(spec("address")))
                
            Case PARAM_DYNAMIC_STRING
                result(i) = ResolveDynamicRef(CStr(spec("refString")), CStr(spec("workbook")), CStr(spec("worksheet")))
                
            Case Else
                result(i) = Empty
        End Select
    Next i
    
    GetResumeArgs = result
    Exit Function
    
ErrorHandler:
    Debug.Print "TaskUnit.GetResumeArgs Error: " & Err.Description
    GetResumeArgs = Array()
End Function

' ===== 私有辅助函数 =====

Private Function ReadRangeValue(wbName As String, wsName As String, addr As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    
    Set wb = Nothing
    On Error Resume Next
    Set wb = Application.Workbooks(wbName)
    On Error GoTo ErrorHandler
    
    If wb Is Nothing Then
        ReadRangeValue = CVErr(xlErrRef)
        Exit Function
    End If
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Sheets(wsName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        ReadRangeValue = CVErr(xlErrRef)
        Exit Function
    End If
    
    Set rng = Nothing
    On Error Resume Next
    Set rng = ws.Range(addr)
    On Error GoTo ErrorHandler
    
    If rng Is Nothing Then
        ReadRangeValue = CVErr(xlErrRef)
        Exit Function
    End If
    
    ReadRangeValue = rng.Value
    Exit Function
    
ErrorHandler:
    ReadRangeValue = CVErr(xlErrRef)
End Function

Private Function ResolveDynamicRef(refStr As String, defaultWb As String, defaultWs As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim wbName As String
    Dim wsName As String
    Dim cellAddr As String
    
    ParseReferenceString refStr, wbName, wsName, cellAddr, defaultWb, defaultWs
    
    ResolveDynamicRef = ReadRangeValue(wbName, wsName, cellAddr)
    Exit Function
    
ErrorHandler:
    ResolveDynamicRef = "#REF: " & refStr
End Function

Private Sub ParseReferenceString(refStr As String, _
                                  ByRef wbName As String, _
                                  ByRef wsName As String, _
                                  ByRef cellAddr As String, _
                                  defaultWb As String, _
                                  defaultWs As String)
    
    wbName = defaultWb
    wsName = defaultWs
    cellAddr = refStr
    
    Dim workStr As String
    workStr = refStr
    
    If Len(workStr) > 0 And Left(workStr, 1) = "[" Then
        Dim bracketEnd As Long
        bracketEnd = InStr(workStr, "]")
        If bracketEnd > 1 Then
            wbName = Mid(workStr, 2, bracketEnd - 2)
            workStr = Mid(workStr, bracketEnd + 1)
        End If
    End If
    
    Dim exclamPos As Long
    exclamPos = InStr(workStr, "!")
    If exclamPos > 0 Then
        wsName = Left(workStr, exclamPos - 1)
        wsName = Replace(wsName, "'", "")
        cellAddr = Mid(workStr, exclamPos + 1)
    Else
        cellAddr = workStr
    End If
End Sub
