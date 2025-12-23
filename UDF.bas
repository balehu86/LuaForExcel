' ============================================
' UDF.bas - 用户自定义函数
' ============================================
' 设计原因：
' 1. 提供 Excel 公式接口
' 2. 通过 CoreRegistry 路由到对应 WorkbookRuntime
' 3. 不直接访问 TaskTable
' 4. LuaEval 和 LuaCall 用于同步 Lua 调用
' ============================================

Option Explicit

' ====LuaTask - 定义协程任务====
' 用法：=LuaTask("funcName", arg1, arg2, "|", resumeArg1, resumeArg2)
' 返回：taskId（字符串）
Public Function LuaTask(ParamArray params() As Variant) As String
    On Error GoTo ErrorHandler
    
    ' 获取调用单元格所在的 Workbook
    Dim wb As Workbook
    Set wb = GetCallerWorkbook()

    If wb Is Nothing Then
        LuaTask = "#ERROR: 无法获取工作簿"
        Exit Function
    End If
    
    ' 获取对应的 WorkbookRuntime
    Dim rt As WorkbookRuntime
    Set rt = CoreRegistry.EnsureRuntimeForWorkbook(wb)
    
    If rt Is Nothing Then
        LuaTask = "#ERROR: 未找到运行时"
        Exit Function
    End If
    
    ' 解析参数
    If UBound(params) < 0 Then
        LuaTask = "#ERROR: 需要函数名"
        Exit Function
    End If
    
    Dim funcName As String
    funcName = CStr(params(0))
    
    Dim cellAddr As String
    cellAddr = Application.Caller.Address(External:=True)
    
    ' 检查是否已有任务（通过 CoreRegistry）
    Dim existingTaskId As String
    existingTaskId = CoreRegistry.FindTaskByCell(cellAddr)
    If existingTaskId <> "" Then
        LuaTask = existingTaskId
        Exit Function
    End If
    
    ' 分离 startArgs 和 resumeSpec
    Dim startList As Object, resumeList As Object
    Set startList = CreateObject("System.Collections.ArrayList")
    Set resumeList = CreateObject("System.Collections.ArrayList")
    
    Dim phase As Long
    phase = 0  ' 0=start, 1=resume
    
    Dim i As Long
    For i = 1 To UBound(params)
        If VarType(params(i)) = vbString And params(i) = "|" Then
            phase = 1
        Else
            If phase = 0 Then
                startList.Add params(i)
            Else
                resumeList.Add params(i)
            End If
        End If
    Next
    
    Dim startArgs As Variant, resumeSpec As Variant
    If startList.Count > 0 Then
        startArgs = startList.ToArray()
    Else
        startArgs = Array()
    End If
    
    If resumeList.Count > 0 Then
        resumeSpec = resumeList.ToArray()
    Else
        resumeSpec = Array()
    End If
    
    ' 创建任务
    Dim taskId As String
    taskId = rt.CreateTask(cellAddr, funcName, startArgs, resumeSpec)
    
    LuaTask = taskId
    Exit Function
    
ErrorHandler:
    LuaTask = "#ERROR: " & Err.Description
End Function
' ====LuaGet - 获取任务字段====
Public Function LuaGet(taskId As String, field As String) As Variant
    On Error GoTo ErrorHandler
    
    Application.Volatile True
    
    ' 通过全局路由表解析运行时
    Dim rt As WorkbookRuntime
    Set rt = CoreRegistry.ResolveRuntime(taskId)
    
    If rt Is Nothing Then
        LuaGet = "#ERROR: 任务不存在"
        Exit Function
    End If
    
    LuaGet = rt.GetTaskField(taskId, field)
    Exit Function
    
ErrorHandler:
    LuaGet = "#ERROR: " & Err.Description
End Function

' ====LuaEval - 执行 Lua 表达式（同步）====
' 用法: =LuaEval("1 + 1")
'      =LuaCall("funcName", arg1, arg2)
Public Function LuaEval(expression As String) As Variant
    On Error GoTo ErrorHandler
    
    Application.Volatile True
    
    ' 获取调用单元格所在的 Workbook
    Set wb = GetCallerWorkbook()  ' 修改此行
    
    If wb Is Nothing Then
        LuaEval = "#ERROR: 无法获取工作簿"
        Exit Function
    End If
    
    ' 获取对应的 WorkbookRuntime
    Dim rt As WorkbookRuntime
    Set rt = CoreRegistry.EnsureRuntimeForWorkbook(wb)
    
    If rt Is Nothing Then
        LuaEval = "#ERROR: 未找到运行时"
        Exit Function
    End If
    
    ' 调用 Runtime 的 Eval 方法
    LuaEval = rt.EvalExpression(expression)
    Exit Function
    
ErrorHandler:
    LuaEval = "#ERROR: " & Err.Description
End Function
Public Function LuaCall(funcName As String, ParamArray args() As Variant) As Variant
    On Error GoTo ErrorHandler
    
    Application.Volatile True
    
    ' 获取调用单元格所在的 Workbook
    Set wb = GetCallerWorkbook()  ' 修改此行
    
    If wb Is Nothing Then
        LuaEval = "#ERROR: 无法获取工作簿"
        Exit Function
    End If
    
    ' 获取对应的 WorkbookRuntime
    Dim rt As WorkbookRuntime
    Set rt = CoreRegistry.EnsureRuntimeForWorkbook(wb)
    
    If rt Is Nothing Then
        LuaCall = "#ERROR: 未找到运行时"
        Exit Function
    End If
    
    ' 将 ParamArray 转换为普通数组
    Dim argArray() As Variant
    If UBound(args) >= LBound(args) Then
        ReDim argArray(LBound(args) To UBound(args))
        Dim i As Long
        For i = LBound(args) To UBound(args)
            argArray(i) = args(i)
        Next
    Else
        argArray = Array()
    End If
    
    ' 调用 Runtime 的 CallFunction 方法
    LuaCall = rt.CallFunction(funcName, argArray)
    Exit Function
    
ErrorHandler:
    LuaCall = "#ERROR: " & Err.Description
End Function

' 在 UDF.bas 顶部添加辅助函数
Private Function GetCallerWorkbook() As Workbook
    On Error Resume Next
    
    ' 尝试从 Caller 获取工作簿
    If TypeName(Application.Caller) = "Range" Then
        Set GetCallerWorkbook = Application.Caller.Parent.Parent
    Else
        ' 如果 Caller 不是 Range，使用 ActiveWorkbook
        Set GetCallerWorkbook = ActiveWorkbook
    End If
    
    ' 确保不是加载宏本身
    If Not GetCallerWorkbook Is Nothing Then
        If GetCallerWorkbook.IsAddin Then
            Set GetCallerWorkbook = ActiveWorkbook
        End If
    End If
End Function