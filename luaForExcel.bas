' ============================================
' Excel-Lua 5.4 集成模块（完整版 + 协程支持）
' 使用纯 DLL 调用方式
' ============================================

Option Explicit

' ===== Lua 5.4 DLL 声明 =====
#If VBA7 Then
    ' 基础 API
    Private Declare PtrSafe Function luaL_newstate Lib "lua54.dll" () As LongPtr
    Private Declare PtrSafe Sub luaL_openlibs Lib "lua54.dll" (ByVal L As LongPtr)
    Private Declare PtrSafe Sub lua_close Lib "lua54.dll" (ByVal L As LongPtr)
    Private Declare PtrSafe Function luaL_loadstring Lib "lua54.dll" (ByVal L As LongPtr, ByVal s As String) As Long
    Private Declare PtrSafe Function lua_pcallk Lib "lua54.dll" (ByVal L As LongPtr, ByVal nargs As Long, ByVal nResults As Long, ByVal msgh As Long, ByVal ctx As LongPtr, ByVal k As LongPtr) As Long
    Private Declare PtrSafe Function lua_tonumberx Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long, ByVal isnum As LongPtr) As Double
    Private Declare PtrSafe Function lua_tolstring Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long, ByVal leng As LongPtr) As LongPtr
    Private Declare PtrSafe Function lua_toboolean Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long) As Long
    Private Declare PtrSafe Function lua_type Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long) As Long
    Private Declare PtrSafe Sub lua_pushnil Lib "lua54.dll" (ByVal L As LongPtr)
    Private Declare PtrSafe Sub lua_pushnumber Lib "lua54.dll" (ByVal L As LongPtr, ByVal n As Double)
    Private Declare PtrSafe Sub lua_pushstring Lib "lua54.dll" (ByVal L As LongPtr, ByVal s As String)
    Private Declare PtrSafe Sub lua_pushboolean Lib "lua54.dll" (ByVal L As LongPtr, ByVal b As Long)
    Private Declare PtrSafe Function lua_gettop Lib "lua54.dll" (ByVal L As LongPtr) As Long
    Private Declare PtrSafe Sub lua_settop Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long)
    Private Declare PtrSafe Function lua_getglobal Lib "lua54.dll" (ByVal L As LongPtr, ByVal name As String) As Long
    Private Declare PtrSafe Sub lua_createtable Lib "lua54.dll" (ByVal L As LongPtr, ByVal narr As Long, ByVal nrec As Long)
    Private Declare PtrSafe Sub lua_rawseti Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long, ByVal n As LongPtr)
    Private Declare PtrSafe Function lua_rawgeti Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long, ByVal n As LongPtr) As Long
    Private Declare PtrSafe Function lua_rawlen Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long) As LongPtr
    Private Declare PtrSafe Function luaL_loadfilex Lib "lua54.dll" (ByVal L As LongPtr, ByVal filename As String, ByVal mode As LongPtr) As Long
    Private Declare PtrSafe Sub lua_setglobal Lib "lua54.dll" (ByVal L As LongPtr, ByVal name As String)
    Private Declare PtrSafe Function lua_next Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long) As Long
    Private Declare PtrSafe Sub lua_pushvalue Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long)
    Private Declare PtrSafe Function lua_getfield Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long, ByVal k As String) As Long
    ' 协程 API
    Private Declare PtrSafe Function lua_newthread Lib "lua54.dll" (ByVal L As LongPtr) As LongPtr
    Private Declare PtrSafe Function lua_resume Lib "lua54.dll" (ByVal L As LongPtr, ByVal from As LongPtr, ByVal narg As Long, ByVal nres As LongPtr) As Long
    Private Declare PtrSafe Function lua_status Lib "lua54.dll" (ByVal L As LongPtr) As Long
    Private Declare PtrSafe Sub lua_xmove Lib "lua54.dll" (ByVal fromL As LongPtr, ByVal toL As LongPtr, ByVal n As Long)
    Private Declare PtrSafe Function luaL_ref Lib "lua54.dll" (ByVal L As LongPtr, ByVal t As Long) As Long
    Private Declare PtrSafe Sub luaL_unref Lib "lua54.dll" (ByVal L As LongPtr, ByVal t As Long, ByVal ref As Long)
    ' 系统 API
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As LongPtr)
    Private Declare PtrSafe Function lstrlenA Lib "kernel32" (ByVal ptr As LongPtr) As Long
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
#Else
    ' 32位版本声明（暂不提供）
#End If
' ===== Lua 类型常量 =====
Private Const LUA_TNIL = 0
Private Const LUA_TBOOLEAN = 1
Private Const LUA_TNUMBER = 3
Private Const LUA_TSTRING = 4
Private Const LUA_TTABLE = 5
Private Const LUA_TFUNCTION = 6
' ===== Lua 状态常量 =====
Private Const LUA_OK = 0
Private Const LUA_YIELD = 1
Private Const LUA_ERRRUN = 2
' ===== 全局变量 =====
Private g_LogLevel As Byte ' 日志等级：0=错误，1=信息，2=调试
Private g_LuaState As LongPtr
Private g_Initialized As Boolean
Private g_HotReloadEnabled As Boolean
Private g_FunctionsPath As String  ' 固定为加载项目录
Private g_LastModified As Date
' ===== 协程全局变量 =====
Public Enum CoStatus
    CO_DEFINED
    CO_YIELD
    CO_PAUSED
    CO_DONE
    CO_ERROR
    CO_TERMINATED
End Enum
Public g_Tasks As Object       ' task Id -> task Instance
Public g_Workbooks As Object    ' Dictionary: wbName -> WorkbookInfo
Public g_TaskQueue As Collection     ' taskId 列表，按调度顺序排列
Public g_Watches As Object          ' Dictionary: watchCell -> WatchInfo
' ===== 调度全局变量 =====
Private g_SchedulerRunning As Boolean   ' 调度器是否运行中
Private g_StateDirty As Boolean         ' 本 tick 是否有状态变化，用来检测是否需要刷新单元格
Public g_NextTaskId As Integer         ' 新建下一个任务ID计数器
Private g_SchedulerIntervalMilliSec As Long ' 调度间隔(ms)
Private g_NextScheduleTime As Date     '标记记下一次调度时间
Private g_CFS_minVruntime As Double       ' 队列中最小的 vruntime（用于新任务初始化）
Private g_CFS_targetLatency As Double     ' 目标延迟周期（ms），默认 100ms
Private g_CFS_minGranularity As Double    ' 最小执行粒度（ms），默认 10ms
Private g_CFS_niceToWeight(0 To 39) As Double  ' nice 值到权重的映射表
' ===== 配置常量 =====
Private Const CP_UTF8 As Long = 65001
Private Const LOG_LEVEL As Byte = 1  ' 默认日志等级：0=错误，1=信息，2=调试
Private Const DEFAULT_HOT_RELOAD_ENABLED As Boolean = True
Private Const SCHEDULER_INTERVAL_Milli_SEC As Long = 1000  ' 调度间隔，默认1000ms

Private Const CFS_DEFAULT_WEIGHT As Double = 1024 ' 默认权重（对应 nice=0）
Private Const CFS_TARGET_LATENCY As Double = 100  ' 目标最小延迟周期（ms）
Private Const CFS_MIN_GRANULARITY As Double = 10  ' 最小执行粒度（ms）

Private Const LUA_REGISTRYINDEX As Long = -1001000
' ===== 性能统计全局变量 =====
Private Type SchedulerStats
    TotalTime As Double      ' 调度器总运行时间(ms)
    LastTime As Double       ' 上次调度花费时间(ms)
    TotalCount As Long       ' 总调度次数
    StartTime As Date        ' 调度器启动时间
End Type
Private g_SchedulerStats As SchedulerStats
' ============================================
' 第一部分：核心初始化和清理
' ============================================
' 主初始化函数：创建空白 Lua 状态机
Public Function InitLuaState() As Boolean
    On Error GoTo ErrorHandler

    If g_Initialized Then
        InitLuaState = True
        Exit Function
    End If

    g_LogLevel = LOG_LEVEL

    ' 创建Lua状态机
    g_LuaState = luaL_newstate()
    If g_LuaState = 0 Then
        MsgBox "无法创建 Lua 状态机。" & vbCrLf & _
               "请确保 lua54.dll 在系统路径中。", vbCritical, "初始化失败"
        InitLuaState = False
        Exit Function
    End If

    luaL_openlibs g_LuaState

    ' 固定为加载项目录下的functions.lua
    g_FunctionsPath = ThisWorkbook.Path & "\functions.lua"
    g_LastModified = #1/1/1900#
    g_HotReloadEnabled = DEFAULT_HOT_RELOAD_ENABLED

    InitCoroutineSystem

    ' 尝试加载functions.lua
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(g_FunctionsPath) Then
        If Not TryLoadFunctionsFile() Then MsgBox "functions.lua 加载失败。" & vbCrLf & "Lua 引擎已启动,但自定义函数不可用。", vbExclamation, "InitLuaState_Warning"
    End If

    InitLuaState = True
    g_Initialized = True
    ' MsgBox "Lua栈初始化完成",vbInformation,"InitLuaState_Info" 
    Exit Function
ErrorHandler:
    MsgBox "初始化 Lua 失败: " & Err.Description, vbCritical, "严重错误"
    InitLuaState = False
End Function

' 初始化协程系统
Private Sub InitCoroutineSystem()
    If g_SchedulerIntervalMilliSec = 0 Then g_SchedulerIntervalMilliSec = SCHEDULER_INTERVAL_Milli_SEC
    ' CFS 参数初始化
    If g_CFS_minVruntime = 0 Then g_CFS_minVruntime = 0
    If g_CFS_targetLatency = 0 Then g_CFS_targetLatency = CFS_TARGET_LATENCY
    If g_CFS_minGranularity = 0 Then g_CFS_minGranularity = CFS_MIN_GRANULARITY
    ' 初始化 nice 到权重的映射表（简化版，只用 0-39 对应 nice -20 到 +19）
    ' 权重公式: weight = 1024 / 1.25^nice  (nice=0 时 weight=1024)
    Dim i As Integer
    For i = 0 To 39
        g_CFS_niceToWeight(i) = 1024 / (1.25 ^ (i - 20))
    Next
    ' 初始化调度器性能统计
    g_SchedulerStats.TotalTime = 0
    g_SchedulerStats.LastTime = 0
    g_SchedulerStats.TotalCount = 0
    g_SchedulerStats.StartTime = Now

    If g_Workbooks Is Nothing Then Set g_Workbooks = CreateObject("Scripting.Dictionary")
    If g_Tasks Is Nothing Then Set g_Tasks = CreateObject("Scripting.Dictionary")
    If g_TaskQueue Is Nothing Then Set g_TaskQueue = New Collection
    If g_Watches Is Nothing Then Set g_Watches = CreateObject("Scripting.Dictionary")

    If g_NextTaskId = 0 Then g_NextTaskId = 1
    g_StateDirty = False
End Sub

' 清理 Lua 状态机
Public Sub CleanupLua()
    If g_Initialized Then
        StopScheduler

        ' 先释放所有协程
        If Not g_Tasks Is Nothing Then
            Dim taskId As Variant
            For Each taskId In g_Tasks.Keys
                ReleaseTaskCoroutine g_Tasks(taskId)
            Next
            g_Tasks.RemoveAll
        End If

        ' 然后关闭 Lua 状态机
        If g_LuaState <> 0 Then
            lua_close g_LuaState
            g_LuaState = 0
        End If

        g_Initialized = False
    End If
End Sub
' ============================================
' 第二部分：functions.lua 加载和热重载
' ============================================
' 在临时状态中验证 functions.lua 语法
Private Function ValidateFunctionsFile() As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(g_FunctionsPath) Then
        ValidateFunctionsFile = False
        Exit Function
    End If

    Dim tempL As LongPtr
    tempL = luaL_newstate()
    If tempL = 0 Then
        ValidateFunctionsFile = False
        Exit Function
    End If

    luaL_openlibs tempL

    Dim stackTop As Long
    stackTop = lua_gettop(tempL)  ' 入口保存

    Dim result As Long
    result = luaL_loadfilex(tempL, g_FunctionsPath, 0)
    If result = 0 Then result = lua_pcallk(tempL, 0, 0, 0, 0, 0)

    If result <> 0 Then
        Dim errMsg As String
        errMsg = GetStringFromState(tempL, -1)
        lua_settop tempL, stackTop
        lua_close tempL
        
        MsgBox "functions.lua 存在语法错误:" & vbCrLf & vbCrLf & _
               errMsg, vbCritical, "文件验证失败"
        ValidateFunctionsFile = False
        Exit Function
    End If

    lua_settop tempL, stackTop
    lua_close tempL
    ValidateFunctionsFile = True
    Exit Function

ErrorHandler:
    If tempL <> 0 Then
        lua_settop tempL, stackTop
        lua_close tempL
    End If
    ValidateFunctionsFile = False
End Function

' 在主状态中加载 functions.lua
Private Function LoadFunctionsIntoMainState() As Boolean
    On Error GoTo ErrorHandler

    Dim stackTop As Long
    stackTop = lua_gettop(g_LuaState)  ' 入口保存

    Dim result As Long
    result = luaL_loadfilex(g_LuaState, g_FunctionsPath, 0)
    If result = 0 Then result = lua_pcallk(g_LuaState, 0, 0, 0, 0, 0)

    If result <> 0 Then
        Dim errMsg As String
        errMsg = GetStringFromState(g_LuaState, -1)
        lua_settop g_LuaState, stackTop  ' 统一恢复

        MsgBox "主状态加载 functions.lua 失败:" & vbCrLf & vbCrLf & _
               errMsg, vbCritical, "加载失败"
        LoadFunctionsIntoMainState = False
        Exit Function
    End If

    g_LastModified = FileDateTime(g_FunctionsPath)
    lua_settop g_LuaState, stackTop  ' 统一恢复
    LoadFunctionsIntoMainState = True
    Exit Function

ErrorHandler:
    MsgBox "加载过程发生 VBA 错误: " & Err.Description, vbCritical, "严重错误"
    If g_Initialized Then lua_settop g_LuaState, stackTop
    LoadFunctionsIntoMainState = False
End Function

' 尝试加载 functions.lua（先验证，再加载）
Private Function TryLoadFunctionsFile() As Boolean
    ' 第一步：验证语法
    If Not ValidateFunctionsFile() Then
        TryLoadFunctionsFile = False
        Exit Function
    End If
    
    ' 第二步：加载到主状态
    TryLoadFunctionsFile = LoadFunctionsIntoMainState()
End Function

' 自动热重载检查（如果启用）
Private Sub CheckAutoReload()
    If Not g_HotReloadEnabled Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(g_FunctionsPath) Then Exit Sub

    Dim currentModified As Date
    currentModified = FileDateTime(g_FunctionsPath)
    If Not (currentModified <> g_LastModified) Then Exit Sub

    Call TryLoadFunctionsFile
End Sub
' ============================================
' 第三部分：公共UDF接口（基础功能）
' ============================================
' 执行 Lua 表达式
Public Function LuaEval(expression As String) As Variant
    On Error GoTo ErrorHandler

    If Not InitLuaState() Then
        LuaEval = CVErr(xlErrValue)
        Exit Function
    End If

    CheckAutoReload

    Dim stackTop As Long
    stackTop = lua_gettop(g_LuaState)  ' 入口保存

    Dim fullCode As String
    fullCode = "return " & expression

    Dim result As Long
    result = luaL_loadstring(g_LuaState, fullCode)
    If result <> 0 Then
        LuaEval = "语法错误: " & GetStringFromState(g_LuaState, -1)
        lua_settop g_LuaState, stackTop  ' 统一恢复
        Exit Function
    End If

    result = lua_pcallk(g_LuaState, 0, 1, 0, 0, 0)
    If result <> 0 Then
        LuaEval = "运行错误: " & GetStringFromState(g_LuaState, -1)
        lua_settop g_LuaState, stackTop  ' 统一恢复
        Exit Function
    End If

    LuaEval = GetValue(g_LuaState, -1)
    lua_settop g_LuaState, stackTop  ' 统一恢复
    Exit Function
ErrorHandler:
    LuaEval = "VBA错误: " & Err.Description
    If g_Initialized Then lua_settop g_LuaState, stackTop  ' 统一恢复
End Function

' 调用 functions.lua 中的函数
Public Function LuaCall(funcName As String, ParamArray args() As Variant) As Variant
    On Error GoTo ErrorHandler

    If Not InitLuaState() Then
        LuaCall = CVErr(xlErrValue)
        Exit Function
    End If

    CheckAutoReload

    Dim stackTop As Long
    stackTop = lua_gettop(g_LuaState)  ' 入口保存

    lua_getglobal g_LuaState, funcName
    If lua_type(g_LuaState, -1) <> LUA_TFUNCTION Then
        LuaCall = "错误: 函数 '" & funcName & "' 不存在"
        lua_settop g_LuaState, stackTop  ' 统一恢复
        Exit Function
    End If

    Dim i As Long, argCount As Long
    argCount = 0
    For i = LBound(args) To UBound(args)
        PushValue g_LuaState, args(i)
        argCount = argCount + 1
    Next i

    Dim result As Long
    result = lua_pcallk(g_LuaState, argCount, -1, 0, 0, 0)
    If result <> 0 Then
        LuaCall = "运行错误: " & GetStringFromState(g_LuaState, -1)
        lua_settop g_LuaState, stackTop  ' 统一恢复
        Exit Function
    End If

    Dim nResults As Long
    nResults = lua_gettop(g_LuaState) - stackTop  ' 相对计算结果数

    If nResults = 0 Then
        LuaCall = Empty
    ElseIf nResults = 1 Then
        LuaCall = GetValue(g_LuaState, -1)
    Else
        Dim results() As Variant
        ReDim results(1 To 1, 1 To nResults)
        For i = 1 To nResults
            results(1, i) = GetValue(g_LuaState, stackTop + i)
        Next i
        LuaCall = results
    End If

    lua_settop g_LuaState, stackTop  ' 统一恢复
    Exit Function
ErrorHandler:
    LuaCall = "VBA错误: " & Err.Description
    If g_Initialized Then lua_settop g_LuaState, stackTop  ' 统一恢复
End Function
' ============================================
' 第四部分：协程UDF接口
' ============================================
' 任务定义函数
Public Function LuaTask(ParamArray params() As Variant) As String
    On Error GoTo ErrorHandler

    If Not InitLuaState() Then
        LuaTask = "#ERROR: Lua未初始化"
        Exit Function
    End If

    If UBound(params) < 0 Then
        LuaTask = "#ERROR: 需要函数名"
        Exit Function
    End If

    ' 获取调用单元格地址和工作簿名
    Dim taskCell As String
    Dim wbName As String
    Dim callerWb As Workbook

    On Error Resume Next
    taskCell = Application.Caller.Address(External:=True)
    Set callerWb = Application.Caller.Worksheet.Parent

    ' 检查调用者工作簿是否存在
    If callerWb Is Nothing Then
        LuaTask = "#ERROR: 无法获取调用工作簿"
        Exit Function
    End If

    ' 防止在xlam文件中创建任务
    If callerWb.FileFormat = xlAddIn Then
        LuaTask = "#ERROR: 不能在宏文件中创建任务"
        MsgBox "禁止在宏文件里创建任务", vbCritical, "LuaTask:Warning"
        Exit Function
    End If

    wbName = callerWb.Name
    On Error GoTo ErrorHandler
    ' 自动注册工作簿
    If Not g_Workbooks.Exists(wbName) Then
        Dim wbInfo As New WorkbookInfo
        wbInfo.wbName = wbName
        g_Workbooks.Add wbName, wbInfo
        Debug.Print "LuaTask自动注册工作簿: " & wbName
    End If

    ' 检查是否已存在任务
    Dim existingTaskId As String
    existingTaskId = FindTaskByCell(taskCell)

    If existingTaskId <> vbNullString Then
        LuaTask = existingTaskId
        Exit Function
    End If

    ' 解析参数
    Dim funcName As String
    funcName = CStr(params(0))

    Dim startArgs As Variant, resumeSpec As Variant
    startArgs = Array()
    resumeSpec = Array()

    Dim phase As Long
    phase = 0

    Dim startList As Object, resumeList As Object
    Set startList = CreateObject("System.Collections.ArrayList")
    Set resumeList = CreateObject("System.Collections.ArrayList")

    Dim i As Long
    For i = 1 To UBound(params)
        If VarType(params(i)) = vbString Then
            If params(i) = "|" Then
                phase = 1
            Else
                Select Case phase
                    Case 0: startList.Add params(i)
                    Case 1: resumeList.Add params(i)
                End Select
            End If
        Else
            Select Case phase
                Case 0: startList.Add params(i)
                Case 1: resumeList.Add params(i)
            End Select
        End If
    Next i

    If startList.Count > 0 Then startArgs = startList.ToArray()
    If resumeList.Count > 0 Then resumeSpec = resumeList.ToArray()

    Dim taskId As String
    taskId = "Task_" & CStr(g_NextTaskId)

    ' 注册任务
    Dim task As New TaskUnit
    With task
        .taskId = g_NextTaskId
        .taskFunc = funcName
        .taskWorkbook = wbName
        .taskStartArgs = startArgs
        .taskResumeSpec = resumeSpec
        .taskCell = taskCell
        .taskStatus = CO_DEFINED
        .taskProgress = 0
        .taskMessage = vbNullString
        .taskValue = vbNull
        .taskError = vbNullString
        .taskCoThread = 0
        .taskLastTime = 0
        .taskTotalTime = 0
        .taskTickCount = 0
        .CFS_weight = CFS_DEFAULT_WEIGHT
        .CFS_vruntime = g_CFS_minVruntime  ' 从当前最小值开始
    End With
    g_Tasks.Add taskId, task

    LuaTask = taskId
    g_NextTaskId = g_NextTaskId + 1
    Exit Function
ErrorHandler:
    Dim errorDetails As String
    errorDetails = "Task错误:" & vbCrLf
    errorDetails = errorDetails & "错误号: " & Err.Number & vbCrLf
    errorDetails = errorDetails & "描述: " & Err.Description & vbCrLf
    ' ' 输出到立即窗口便于调试
    Debug.Print "=== Task错误详情 ==="
    Debug.Print errorDetails
    Debug.Print "======================="
    LuaTask = "#ERROR: Task" & Err.Description
End Function

' 读取任务状态（静态读取，不自动刷新）
' 如需实时监控，请使用 LuaWatch 函数
Public Function LuaGet(taskId As String, field As String) As Variant
    On Error GoTo ErrorHandler

    ' 完全移除 Volatile
    ' 此函数只在输入公式或手动刷新(F9)时计算
    ' 实时监控请使用 LuaWatch

    If Not InitLuaState() Then
        LuaGet = CVErr(xlErrValue)
        Exit Function
    End If

    If Not g_Tasks.Exists(taskId) Then
        LuaGet = "#ERROR: 任务不存在"
        Exit Function
    End If

    Dim task As TaskUnit
    Set task = g_Tasks(taskId)
    Dim taskstatus As String
    Select Case task.taskStatus
        Case CO_DEFINED: taskstatus = "defined"
        Case CO_YIELD: taskstatus = "yield"
        Case CO_PAUSED: taskstatus = "paused"
        Case CO_DONE: taskstatus = "done"
        Case CO_ERROR: taskstatus = "error"
    End Select

    Select Case LCase(field)
        Case "status"
            LuaGet = taskstatus
        Case "progress"
            LuaGet = task.taskProgress
        Case "message"
            LuaGet = task.taskMessage
        Case "value"
            LuaGet = task.taskValue
        Case "error"
            LuaGet = task.taskError
        Case "summary"
            Dim summary As String
            summary = "状态:" & taskstatus
            summary = summary & " | 进度:" & Format(task.taskProgress, "0.0") & "%"
            If taskstatus = "error" Then
                summary = summary & " | 错误:" & Left(task.taskError, 30)
            End If
            LuaGet = summary
        Case Else
            LuaGet = "#ERROR: 未知字段"
    End Select

    Exit Function
ErrorHandler:
    LuaGet = "#ERROR: " & Err.Description
End Function

' 监控任务字段变化（注册监控点，由调度器统一刷新）
Public Function LuaWatch(taskIdOrCell As Variant, field As String, _
                         Optional targetCell As Variant, _
                         Optional direction As Integer = 0) As Variant
    On Error GoTo ErrorHandler

    If Not InitLuaState() Then
        LuaWatch = CVErr(xlErrValue)
        Exit Function
    End If

    ' 初始化监控字典
    If g_Watches Is Nothing Then Set g_Watches = CreateObject("Scripting.Dictionary")
    ' 获取调用单元格信息
    Dim callerCell As Range
    Dim callerAddr As String
    Dim callerWb As Workbook

    On Error Resume Next
    Set callerCell = Application.Caller
    If callerCell Is Nothing Then
        LuaWatch = "#ERROR: 只能在单元格中使用"
        Exit Function
    End If
    callerAddr = callerCell.Address(External:=True)
    Set callerWb = callerCell.Worksheet.Parent
    On Error GoTo ErrorHandler

    ' 解析 taskId
    Dim taskId As String
    If TypeName(taskIdOrCell) = "Range" Then
        taskId = CStr(taskIdOrCell.Value)
    Else
        taskId = CStr(taskIdOrCell)
    End If

    ' 验证任务存在
    If Not g_Tasks.Exists(taskId) Then
        LuaWatch = "#ERROR: 任务不存在"
        Exit Function
    End If

    ' 计算目标单元格地址
    Dim targetAddr As String
    Dim targetRange As Range

    If IsMissing(targetCell) Or IsEmpty(targetCell) Then
        Select Case direction
            Case 0: Set targetRange = callerCell.Offset(0, 1)  ' 右
            Case 1: Set targetRange = callerCell.Offset(-1, 0) ' 上
            Case 2: Set targetRange = callerCell.Offset(0, -1) ' 左
            Case 3: Set targetRange = callerCell.Offset(1, 0)  ' 下
            Case Else: Set targetRange = callerCell.Offset(0, 1)
        End Select
        targetAddr = targetRange.Address(External:=True)
    Else
        If TypeName(targetCell) = "Range" Then
            targetAddr = targetCell.Address(External:=True)
        Else
            On Error Resume Next
            Set targetRange = callerWb.Sheets(callerCell.Worksheet.Name).Range(CStr(targetCell))
            If targetRange Is Nothing Then
                targetAddr = callerCell.Offset(0, 1).Address(External:=True)
            Else
                targetAddr = targetRange.Address(External:=True)
            End If
            On Error GoTo ErrorHandler
        End If
    End If

    ' 检查是否已存在相同的监控
    Dim wi As WatchInfo
    Dim needUpdateIndex As Boolean
    needUpdateIndex = False

    If g_Watches.Exists(callerAddr) Then
        ' 已存在监控：检查参数是否变化
        Set wi = g_Watches(callerAddr)

        ' 检查关键参数是否变化
        Dim paramsChanged As Boolean
        paramsChanged = False

        ' 检查 taskId
        If wi.watchTaskId <> taskId Then
            paramsChanged = True
            needUpdateIndex = True
        End If
        ' 检查 field
        If wi.watchField <> LCase(Trim(field)) Then
            paramsChanged = True
        End If
        ' 检查 targetAddr
        If wi.watchTargetCell <> targetAddr Then
            paramsChanged = True
        End If

        ' 只有参数变化时才更新
        If paramsChanged Then
            ' 更新二级索引（如果 taskId 变化）
            If needUpdateIndex Then
                RemoveFromTaskWatches wi.watchTask, callerAddr  ' 使用旧的 task 对象
                AddToTaskWatches g_Tasks(taskId), callerAddr    ' 使用新的 taskId
            End If

            ' 更新监控属性
            With wi
                .watchTaskId = taskId
                .watchField = LCase(Trim(field))
                .watchTargetCell = targetAddr
                .watchDirection = direction
                ' 参数变化，清空上次值，标记为脏
                .watchLastValue = Empty
                .watchDirty = True
            End With
        End If
        ' 参数未变化时，不修改任何状态

    Else
        ' 新建监控
        Set wi = New WatchInfo

        With wi
            .watchCell = callerAddr
            .watchTaskId = taskId
            .watchField = LCase(Trim(field))
            .watchTargetCell = targetAddr
            .watchDirection = direction
            .watchWorkbook = callerWb.Name
            .watchLastValue = Empty
            .watchDirty = True  ' 新监控需要首次写入
        End With
        Set wi.watchTask = g_Tasks(taskId)

        ' 添加到主索引
        g_Watches.Add callerAddr, wi
        ' 添加到二级索引
        AddToTaskWatches g_Tasks(taskId), callerAddr
    End If

    ' 返回静态描述文本
    LuaWatch = "监控: " & taskId & "." & field & " -> " & targetAddr

    Exit Function
ErrorHandler:
    LuaWatch = "#ERROR: " & Err.Description
End Function
' 刷新所有脏的监控（批量写入，不触发重算）
Private Sub RefreshWatches()
    On Error Resume Next

    If g_Watches Is Nothing Then Exit Sub
    If g_Watches.Count = 0 Then Exit Sub

    Dim watchCell As Variant
    Dim watchInfo As WatchInfo
    Dim task As TaskUnit
    Dim currentValue As Variant
    Dim writeCount As Long
    writeCount = 0

    ' 收集所有需要写入的工作表，最后统一处理
    Dim sheetsToRefresh As Object
    Set sheetsToRefresh = CreateObject("Scripting.Dictionary")

    For Each watchCell In g_Watches.Keys
        Set watchInfo = g_Watches(watchCell)
        ' 只处理脏的监控
        If Not watchInfo.watchDirty Then GoTo NextWatch

        ' 检查任务是否存在
        If Not g_Tasks.Exists(watchInfo.watchTaskId) Then
            ' 任务已删除，移除监控和二级索引
            RemoveFromTaskWatches watchInfo.watchTask, CStr(watchCell)
            g_Watches.Remove watchCell
            GoTo NextWatch
        End If

        Set task = watchInfo.watchTask

        ' 获取当前字段值
        Select Case watchInfo.watchField
            Case "status"
                currentValue = task.taskStatus
            Case "progress"
                currentValue = task.taskProgress
            Case "message"
                currentValue = task.taskMessage
            Case "value"
                currentValue = task.taskValue
            Case "error"
                currentValue = task.taskError
            Case Else
                currentValue = "#未知字段"
        End Select

        ' 检查值是否真的变化了（避免不必要的写入）
        Dim needWrite As Boolean
        needWrite = False
        If IsEmpty(watchInfo.watchLastValue) Then
            needWrite = True
        ElseIf IsArray(currentValue) Or IsArray(watchInfo.watchLastValue) Then
            needWrite = True  ' 数组总是写入
        ElseIf currentValue <> watchInfo.watchLastValue Then
            needWrite = True
        End If

        ' 写入目标单元格（直接写值，不触发计算）
        If needWrite Then
            WriteToTargetCellDirect watchInfo.watchTargetCell, currentValue, watchInfo.watchWorkbook, sheetsToRefresh
            watchInfo.watchLastValue = currentValue
            writeCount = writeCount + 1
        End If
        ' 清除脏标记
        watchInfo.watchDirty = False
NextWatch:
    Next watchCell
    ' 只在有实际写入时，统一刷新一次
    ' 这里不再调用 Calculate，因为直接写值不需要重算
End Sub
' 直接写入目标单元格（不触发 Calculate）
Private Sub WriteToTargetCellDirect(targetAddr As String, value As Variant, wbName As String, sheetsDict As Object)
    On Error Resume Next
    Dim targetRange As Range
    Dim wb As Workbook

    Set wb = Application.Workbooks(wbName)
    If wb Is Nothing Then Exit Sub
    ' 解析地址
    Dim sheetName As String
    Dim cellAddr As String
    Dim exclamPos As Long
    Dim bracketEnd As Long
    ' 处理外部引用格式 [Book1.xlsx]Sheet1!$A$1
    If Left(targetAddr, 1) = "[" Then
        bracketEnd = InStr(targetAddr, "]")
        exclamPos = InStr(targetAddr, "!")
        If exclamPos > 0 Then
            sheetName = Mid(targetAddr, bracketEnd + 1, exclamPos - bracketEnd - 1)
            cellAddr = Mid(targetAddr, exclamPos + 1)
        Else
            cellAddr = Mid(targetAddr, bracketEnd + 1)
        End If
    ElseIf InStr(targetAddr, "!") > 0 Then
        exclamPos = InStr(targetAddr, "!")
        sheetName = Left(targetAddr, exclamPos - 1)
        cellAddr = Mid(targetAddr, exclamPos + 1)
        sheetName = Replace(sheetName, "'", "")
    Else
        cellAddr = targetAddr
    End If
    ' 获取目标范围
    If sheetName <> "" Then
        Set targetRange = wb.Sheets(sheetName).Range(cellAddr)
    Else
        Set targetRange = wb.ActiveSheet.Range(cellAddr)
    End If
    If targetRange Is Nothing Then Exit Sub
    ' 直接写值，不触发事件和计算
    If IsArray(value) Then
        targetRange.Resize(UBound(value, 1) - LBound(value, 1) + 1, _
                          UBound(value, 2) - LBound(value, 2) + 1).value = value
    Else
        targetRange.value = value
    End If
End Sub
' 辅助函数：添加到二级索引（防重复）
Private Sub AddToTaskWatches(task As TaskUnit, watchCell As String)
    If task Is Nothing Then Exit Sub

    ' 检查是否已存在
    Dim exists As Boolean
    exists = False
    On Error Resume Next
    Dim dummy As Variant
    Dim i As Long
    For i = 1 To task.taskWatches.Count
        If task.taskWatches(i) = watchCell Then
            exists = True
            Exit For
        End If
    Next
    On Error GoTo 0
    ' 不存在才添加
    If Not exists Then
        task.taskWatches.Add watchCell
    End If
End Sub
' 辅助函数：从二级索引移除
Private Sub RemoveFromTaskWatches(task As TaskUnit, watchCell As String)
    On Error Resume Next

    If task Is Nothing Then Exit Sub

    Dim i As Long
    For i = task.taskWatches.Count To 1 Step -1
        If task.taskWatches(i) = watchCell Then
            task.taskWatches.Remove i
            Exit For
        End If
    Next
End Sub
' 优化后的 MarkWatchesDirty - O(m) 复杂度
Private Sub MarkWatchesDirty(task As TaskUnit)
    On Error Resume Next
    If task Is Nothing Then Exit Sub

    Dim wc As Variant
    For Each wc In task.taskWatches
        If g_Watches.Exists(CStr(wc)) Then
            g_Watches(CStr(wc)).watchDirty = True
        End If
    Next
End Sub
' ============================================
' 第五部分：协程执行和调度
' ============================================
' 启动协程
Public Sub StartLuaCoroutine(taskId As String)
    On Error GoTo ErrorHandler
    If Not InitLuaState() Then
        MsgBox "Lua状态初始化失败", vbCritical
        Exit Sub
    End If
    If Not g_Tasks.Exists(taskId) Then
        MsgBox "错误：任务 " & taskId & " 不存在", vbCritical
        Exit Sub
    End If
    Dim task As TaskUnit
    Set task = g_Tasks(taskId)
    If task.taskStatus <> CO_DEFINED Then
        MsgBox "错误：任务已启动或已完成，当前状态: " & task.taskStatus, vbExclamation
        Exit Sub
    End If

    ' 统一释放旧协程（防止泄漏）
    ReleaseTaskCoroutine task

    ' 创建协程并锚定到注册表
    Dim coThread As LongPtr
    coThread = lua_newthread(g_LuaState)
    If coThread = 0 Then
        SetTaskStatus task, CO_ERROR
        task.taskError = "无法创建协程线程"
        Exit Sub
    End If
    ' 锚定协程
    task.taskCoRef = luaL_ref(g_LuaState, LUA_REGISTRYINDEX)
    task.taskCoThread = coThread

    ' 获取函数并移动到协程栈
    lua_getglobal g_LuaState, task.taskFunc
    If lua_type(g_LuaState, -1) <> LUA_TFUNCTION Then
        SetTaskStatus task, CO_ERROR
        task.taskError = "函数 '" & task.taskFunc & "' 不存在"
        lua_settop g_LuaState, 0
        Exit Sub
    End If
    lua_xmove g_LuaState, coThread, 1
    lua_pushstring coThread, task.taskCell
    Dim nargs As Long
    nargs = 1
    Dim startArgs As Variant
    startArgs = task.taskStartArgs
    If IsArray(startArgs) Then
        Dim i As Long
        For i = LBound(startArgs) To UBound(startArgs)
            PushValue coThread, startArgs(i)
            nargs = nargs + 1
        Next i
    End If
    Dim nres As LongPtr
    Dim result As Long
    result = lua_resume(coThread, g_LuaState, nargs, VarPtr(nres))
    ' 处理结果（内部会调用 SetTaskStatus）
    HandleCoroutineResult task, result, CLng(nres)
    ' 如果是 yield 状态，加入队列
    If task.taskStatus = CO_YIELD Then
        task.CFS_vruntime = g_CFS_minVruntime
        task.CFS_lastScheduled = GetTickCount()
        TaskQueueAdd task
        StartSchedulerIfNeeded
    End If
    Exit Sub
ErrorHandler:
    SetTaskStatus task, CO_ERROR
    task.taskError = "VBA错误: " & Err.Description & " (行 " & Erl & ")"
    MsgBox "启动协程失败: " & Err.Description, vbCritical
End Sub

' 启动调度器
Private Sub StartSchedulerIfNeeded()
    If g_SchedulerRunning Then Exit Sub
    If g_TaskQueue.Count = 0 Then Exit Sub
    g_SchedulerRunning = True
    g_NextScheduleTime = Now + g_SchedulerIntervalMilliSec / 86400000#
    Application.OnTime g_NextScheduleTime, "SchedulerTick" 
End Sub

' 调度器心跳 - 主入口 （添加定期清理）
Public Sub SchedulerTick()
    On Error GoTo ErrorHandler
    If Not g_SchedulerRunning Then Exit Sub
    If g_TaskQueue.Count = 0 Then
        g_SchedulerRunning = False
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Dim schedulerStart As Double
    schedulerStart = GetTickCount()
    ' 使用 CFS 调度算法 
    Call ScheduleByCFS
    ' 性能计时统计
    Dim schedulerElapsed As Double
    schedulerElapsed = GetTickCount() - schedulerStart
    g_SchedulerStats.LastTime = schedulerElapsed
    g_SchedulerStats.TotalTime = g_SchedulerStats.TotalTime + schedulerElapsed
    g_SchedulerStats.TotalCount = g_SchedulerStats.TotalCount + 1
    ' ★ 修复：每100次调度清理一次孤儿Watch
    If g_SchedulerStats.TotalCount Mod 100 = 0 Then
        CleanupOrphanWatches
    End If
    ' 刷新监控
    If g_StateDirty Then
        RefreshWatches
        g_StateDirty = False
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ' 重新安排下一次调度
    If g_TaskQueue.Count > 0 Then
        g_NextScheduleTime = Now + g_SchedulerIntervalMilliSec / 86400000#
        Application.OnTime g_NextScheduleTime, "SchedulerTick"
    Else
        g_SchedulerRunning = False
    End If
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Debug.Print "SchedulerTick Error: " & Err.Description
End Sub

' CFS 调度核心算法
Private Sub ScheduleByCFS()
    If g_TaskQueue.Count = 0 Then Exit Sub

    Dim tickBudget As Double
    Dim taskStart As Double, taskElapsed As Double
    Dim selectedTask As TaskUnit

    ' 计算本次 tick 的时间预算
    tickBudget = g_SchedulerIntervalMilliSec * 0.8  ' 留 20% 余量给 VBA 开销

    Dim totalElapsed As Double
    totalElapsed = 0

    Do While totalElapsed < tickBudget And g_TaskQueue.Count > 0
        ' 1. 选择 vruntime 最小的任务
        Set selectedTask = CFS_PickNextTask()

        If selectedTask Is Nothing Then Exit Do

        If Not g_Tasks.Exists("Task_" & selectedTask.taskId) Then
            ' 僵尸任务，移除
            TaskQueueRemove selectedTask
            GoTo ContinueLoop
        End If

        ' 只调度 yielded 状态
        If selectedTask.taskStatus <> CO_YIELD Then
            GoTo ContinueLoop
        End If

        ' 2. 执行任务
        taskStart = GetTickCount()

        ResumeCoroutine selectedTask

        taskElapsed = GetTickCount() - taskStart
        If taskElapsed < g_CFS_minGranularity Then taskElapsed = g_CFS_minGranularity

        totalElapsed = totalElapsed + taskElapsed

        ' 3. 更新 vruntime
        CFS_UpdateVruntime selectedTask, taskElapsed

        ' 4. 更新工作簿统计（保留此功能）
        If g_Workbooks.Exists(selectedTask.taskWorkbook) Then
            With g_Workbooks(selectedTask.taskWorkbook)
                .wbLastTime = taskElapsed
                .wbTotalTime = .wbTotalTime + taskElapsed
                .wbTickCount = .wbTickCount + 1
            End With
        End If

        ' 5. 终止态清理
        Select Case selectedTask.taskStatus
            Case CO_DONE, CO_ERROR
                TaskQueueRemove selectedTask
        End Select

ContinueLoop:
    Loop
End Sub
' 选择 vruntime 最小的任务（CFS 红黑树的简化实现）
Private Function CFS_PickNextTask() As TaskUnit
    Dim minVruntime As Double
    Dim t As Variant
    Dim task As TaskUnit
    minVruntime = 1E+308  ' Double 最大值
    Set CFS_PickNextTask = Nothing
    For Each t In g_TaskQueue
        Set task = t
        ' 只考虑 yielded 状态的任务
        If task.taskStatus = CO_YIELD Then
            If task.CFS_vruntime < minVruntime Then
                minVruntime = task.CFS_vruntime
                Set CFS_PickNextTask = task
            End If
        End If
    Next
End Function
' 更新任务的 vruntime
Private Sub CFS_UpdateVruntime(task As TaskUnit, actualRuntime As Double)
    Dim vruntimeDelta As Double
    vruntimeDelta = actualRuntime * (CFS_DEFAULT_WEIGHT / task.CFS_weight)

    task.CFS_vruntime = task.CFS_vruntime + vruntimeDelta
    task.CFS_lastScheduled = GetTickCount()

    ' 更新全局最小 vruntime
    If task.CFS_vruntime > g_CFS_minVruntime Then
        Dim minV As Double
        Dim t As Variant
        Dim tsk As TaskUnit  ' 改名避免与参数冲突

        minV = task.CFS_vruntime
        For Each t In g_TaskQueue
            Set tsk = t
            If tsk.taskStatus = CO_YIELD And tsk.CFS_vruntime < minV Then
                minV = tsk.CFS_vruntime
            End If
        Next
        g_CFS_minVruntime = minV
    End If
End Sub

' Resume 协程
Private Sub ResumeCoroutine(task As TaskUnit)
    On Error GoTo ErrorHandler

    If task.taskStatus <> CO_YIELD Then Exit Sub

    ' 检查协程是否有效
    If Not task.HasValidCoroutine() Then
        SetTaskStatus task, CO_ERROR
        task.taskError = "协程引用无效"
        Exit Sub
    End If

    Dim taskStart As Long
    taskStart = GetTickCount()

    Dim coThread As LongPtr
    coThread = task.taskCoThread

    ' 检查协程状态
    Dim status As Long
    status = lua_status(coThread)
    If status <> LUA_OK And status <> LUA_YIELD Then
        SetTaskStatus task, CO_ERROR
        task.taskError = "协程状态异常: " & status
        Exit Sub
    End If

    ' 检查工作簿是否仍然打开
    Dim wbName As String
    wbName = task.taskWorkbook
    Dim wb As Workbook

    If Not IsEmpty(wbName) And wbName <> vbNullString And wbName <> "" Then
        Dim wbExists As Boolean
        wbExists = False

        On Error Resume Next
        Set wb = Application.Workbooks(wbName)
        wbExists = Not (wb Is Nothing)
        On Error GoTo ErrorHandler

        If Not wbExists Then
            SetTaskStatus task, CO_ERROR
            task.taskError = "工作簿已关闭: " & wbName
            Exit Sub
        End If
    Else
        wbName = vbNullString
        Set wb = Nothing
    End If

    ' 清空协程栈
    lua_settop coThread, 0

    ' 准备 resume 参数
    Dim resumeSpec As Variant
    resumeSpec = task.taskResumeSpec

    Dim nargs As Long
    nargs = 0

    If IsArray(resumeSpec) Then
        Dim i As Long
        For i = LBound(resumeSpec) To UBound(resumeSpec)
            Dim param As Variant
            param = resumeSpec(i)

            If VarType(param) = vbString Then
                Dim paramStr As String
                paramStr = Trim(CStr(param))

                If Len(paramStr) > 0 And Not wb Is Nothing Then
                    On Error Resume Next
                    Dim rng As Range
                    Dim ws As Worksheet
                    Set ws = GetWorksheetFromAddress(task.taskCell, wb)
                    Set rng = ws.Range(paramStr)

                    If Err.Number = 0 And Not rng Is Nothing Then
                        If rng.Cells.Count = 1 Then
                            PushValue coThread, rng.Value
                        Else
                            PushValue coThread, rng.Value
                        End If
                    Else
                        PushValue coThread, paramStr
                    End If
                    Err.Clear
                    On Error GoTo ErrorHandler
                Else
                    PushValue coThread, paramStr
                End If
            Else
                PushValue coThread, param
            End If

            nargs = nargs + 1
        Next
    End If

    ' 执行 resume
    Dim nres As LongPtr
    Dim result As Long
    result = lua_resume(coThread, g_LuaState, nargs, VarPtr(nres))

    ' 处理结果（内部会调用 SetTaskStatus 处理终态）
    HandleCoroutineResult task, result, CLng(nres)
    ' 每次 Resume 后都标记该任务的监控为脏
    MarkWatchesDirty task
    g_StateDirty = True

    ' 性能统计
    Dim taskElapsed As Double
    taskElapsed = GetTickCount() - taskStart
    task.taskLastTime = taskElapsed
    task.taskTotalTime = task.taskTotalTime + taskElapsed
    task.taskTickCount = task.taskTickCount + 1

    Exit Sub

ErrorHandler:
    Dim errorDetails As String
    errorDetails = "Resume错误: " & Err.Description & " (行 " & Erl & ")"

    SetTaskStatus task, CO_ERROR
    task.taskError = errorDetails

    Debug.Print "=== Resume 错误 ===" & vbCrLf & errorDetails
End Sub

' 手动停止调度器
Public Sub StopScheduler()
    ' 停止调度标志
    If g_SchedulerRunning Then
        g_SchedulerRunning = False
        ' 尝试取消所有 OnTime 调度
        On Error Resume Next
        Application.OnTime g_NextScheduleTime, "SchedulerTick", , False
        MsgBox "调度器已停止。" & vbCrLf & _
            "活跃任务将不会继续执行。" & vbCrLf & vbCrLf
    End If
End Sub

' 恢复调度器
Private Sub ResumeScheduler()
    If g_TaskQueue Is Nothing Or g_TaskQueue.Count = 0 Then
        MsgBox "队列中没有任务，无需启动调度器。", vbExclamation, "无任务"
        Exit Sub
    End If

    If g_SchedulerRunning Then
        MsgBox "调度器已在运行中。", vbInformation, "调度器状态"
        Exit Sub
    End If

    g_SchedulerRunning = True
    g_NextScheduleTime = Now + g_SchedulerIntervalMilliSec / 86400000#
    Application.OnTime g_NextScheduleTime, "SchedulerTick"

    MsgBox "调度器已启动。" & vbCrLf & _
           "当前队列任务数: " & g_TaskQueue.Count, vbInformation, "调度器已启动"
End Sub
' ============================================
' 第六部分：辅助函数（内部使用）
' ============================================
' 统一压栈函数 - 支持主状态机和协程线程
Private Sub PushValue(ByVal L As LongPtr, ByVal value As Variant)
    ' 处理 Range 对象
    If TypeName(value) = "Range" Then
        Dim rng As Range
        Set rng = value
        If rng.Cells.Count = 1 Then
            ' 单个单元格，递归调用处理其值
            PushValue L, rng.value
        Else
            ' 多个单元格，获取数组后递归调用
            PushValue L, rng.value
        End If
        Exit Sub
    End If
    
    ' 处理数组
    If IsArray(value) Then
        PushArray L, value
        Exit Sub
    End If
    
    ' 处理基本类型
    If IsEmpty(value) Or IsNull(value) Then
        lua_pushnil L
    ElseIf IsNumeric(value) Then
        lua_pushnumber L, CDbl(value)
    ElseIf VarType(value) = vbBoolean Then
        lua_pushboolean L, IIf(value, 1, 0)
    Else
        lua_pushstring L, CStr(value)
    End If
End Sub

' 统一数组压栈函数 - 支持主状态机和协程线程
Private Sub PushArray(ByVal L As LongPtr, arr As Variant)
    Dim i As Long, j As Long
    Dim rows As Long, cols As Long

    ' 处理一维数组
    On Error Resume Next
    rows = UBound(arr, 1) - LBound(arr, 1) + 1
    cols = UBound(arr, 2) - LBound(arr, 2) + 1

    If Err.Number <> 0 Then
        ' 一维数组
        Err.Clear
        On Error GoTo 0
        rows = UBound(arr) - LBound(arr) + 1

        lua_createtable L, rows, 0
        For i = LBound(arr) To UBound(arr)
            PushValue L, arr(i)  ' 递归调用 PushValue
            lua_rawseti L, -2, i - LBound(arr) + 1
        Next i
        Exit Sub
    End If
    On Error GoTo 0

    ' 二维数组
    lua_createtable L, rows, 0
    For i = LBound(arr, 1) To UBound(arr, 1)
        lua_createtable L, cols, 0
        For j = LBound(arr, 2) To UBound(arr, 2)
            PushValue L, arr(i, j)  ' 递归调用 PushValue
            lua_rawseti L, -2, j - LBound(arr, 2) + 1
        Next j
        lua_rawseti L, -2, i - LBound(arr, 1) + 1
    Next i
End Sub

' 处理协程返回结果
Private Sub HandleCoroutineResult(task As TaskUnit, result As Long, nres As Long)
    On Error GoTo ErrorHandler

    Dim coThread As LongPtr
    coThread = task.taskCoThread

    ' 处理前记录栈顶（此时栈上是 nres 个返回值）
    Dim stackTopBefore As Long
    stackTopBefore = lua_gettop(coThread)

    Select Case result
        Case LUA_OK
            If nres > 0 And stackTopBefore > 0 Then
                Dim retData As Variant
                retData = GetValue(coThread, -1)
                ParseYieldReturn task, retData, True
            End If
            SetTaskStatus task, CO_DONE
            task.taskProgress = 100

        Case LUA_YIELD
            If nres > 0 And stackTopBefore > 0 Then
                Dim yieldData As Variant
                yieldData = GetValue(coThread, -1)
                ParseYieldReturn task, yieldData, False
            End If
            If task.taskStatus <> CO_DONE And task.taskStatus <> CO_ERROR Then
                SetTaskStatus task, CO_YIELD
            End If

        Case Else
            If nres > 0 And stackTopBefore > 0 Then
                task.taskError = GetStringFromState(coThread, -1)
            Else
                task.taskError = "协程错误: 代码 " & result
            End If
            SetTaskStatus task, CO_ERROR
    End Select

    ' 清空协程栈上的返回值，为下次 resume 做准备
    lua_settop coThread, 0  ' 更明确：直接清空而不是恢复

    Exit Sub

ErrorHandler:
    task.taskError = "处理结果错误: " & Err.Description
    SetTaskStatus task, CO_ERROR
    If coThread <> 0 Then lua_settop coThread, 0
End Sub


' 从 Lua 栈获取字符串
Private Function GetStringFromState(ByVal L As LongPtr, ByVal idx As Long) As String
    Dim ptr As LongPtr
    Dim length As Long

    ptr = lua_tolstring(L, idx, VarPtr(length))
    If ptr = 0 Then
        GetStringFromState = vbNullString
        Exit Function
    End If

    If length = 0 Then
        GetStringFromState = vbNullString
        Exit Function
    End If

    ' 使用Windows API转换UTF-8
    ' 计算需要的缓冲区大小
    Dim nChars As Long
    nChars = MultiByteToWideChar(CP_UTF8, 0, ptr, length, 0, 0)

    If nChars > 0 Then
        ' 分配字符串缓冲区
        GetStringFromState = String$(nChars, 0)
        ' 执行转换
        MultiByteToWideChar CP_UTF8, 0, ptr, length, StrPtr(GetStringFromState), nChars
    End If
End Function

' 从 Lua 栈获取值
Private Function GetValue(ByVal L As LongPtr, ByVal idx As Long) As Variant
    Dim luaType As Long
    luaType = lua_type(L, idx)

    Select Case luaType
        Case LUA_TNIL
            GetValue = Empty
        Case LUA_TBOOLEAN
            GetValue = (lua_toboolean(L, idx) <> 0)
        Case LUA_TNUMBER
            GetValue = lua_tonumberx(L, idx, 0)
        Case LUA_TSTRING
            GetValue = GetStringFromState(L, idx)
        Case LUA_TTABLE
            GetValue = TableToVariant(L, idx)
        Case Else
            GetValue = "#LUA_TYPE_" & luaType
    End Select
End Function

' 将 Lua table 转换为 VBA Variant (字典或数组)
Private Function TableToVariant(ByVal L As LongPtr, ByVal idx As Long) As Variant
    On Error GoTo ErrorHandler

    ' 标准化索引为正数
    If idx < 0 Then
        idx = lua_gettop(L) + idx + 1
    End If

    ' 检查数组长度
    Dim length As LongPtr
    length = lua_rawlen(L, idx)

    ' 如果长度为0，尝试判断是否为字典
    If length = 0 Then
        ' 尝试获取第一个键值对
        Dim topBefore As Long
        topBefore = lua_gettop(L)

        lua_pushnil L
        If lua_next(L, idx) <> 0 Then
            ' 有内容，是字典
            lua_settop L, topBefore  ' 恢复栈
            TableToVariant = TableToDictArray(L, idx)
        Else
            ' 空表
            TableToVariant = Empty
        End If
        Exit Function
    End If

    ' 检查是否为纯数组（所有键都是1到length的连续整数）
    Dim isPureArray As Boolean
    isPureArray = True

    Dim testTop As Long
    testTop = lua_gettop(L)

    lua_pushnil L
    Do While lua_next(L, idx) <> 0
        Dim keyType As Long
        keyType = lua_type(L, -2)

        If keyType <> LUA_TNUMBER Then
            isPureArray = False
            lua_settop L, testTop  ' 立即恢复栈
            Exit Do
        End If

        Dim keyNum As Double
        keyNum = lua_tonumberx(L, -2, 0)

        ' 检查是否为整数且在范围内
        If keyNum <> CLng(keyNum) Or keyNum < 1 Or keyNum > length Then
            isPureArray = False
            lua_settop L, testTop
            Exit Do
        End If

        lua_settop L, -2  ' 只弹出value，保留key
    Loop

    lua_settop L, testTop  ' 确保栈恢复

    ' 如果不是纯数组，按字典处理
    If Not isPureArray Then
        TableToVariant = TableToDictArray(L, idx)
        Exit Function
    End If

    ' 纯数组处理
    ' 检查第一个元素
    lua_rawgeti L, idx, 1
    Dim firstIsTable As Boolean
    firstIsTable = (lua_type(L, -1) = LUA_TTABLE)
    lua_settop L, -2

    If firstIsTable Then
        ' 二维数组
        lua_rawgeti L, idx, 1
        Dim cols As LongPtr
        cols = lua_rawlen(L, -1)
        lua_settop L, -2

        If cols = 0 Then cols = 1  ' 防止空子表

        Dim arr2D() As Variant
        ReDim arr2D(1 To CLng(length), 1 To CLng(cols))

        Dim i As Long, j As Long
        For i = 1 To CLng(length)
            lua_rawgeti L, idx, CLng(i)

            If lua_type(L, -1) = LUA_TTABLE Then
                Dim subLen As LongPtr
                subLen = lua_rawlen(L, -1)

                For j = 1 To CLng(cols)
                    If j <= subLen Then
                        lua_rawgeti L, -1, CLng(j)
                        arr2D(i, j) = GetValue(L, -1)
                        lua_settop L, -2
                    Else
                        arr2D(i, j) = Empty
                    End If
                Next j
            Else
                arr2D(i, 1) = GetValue(L, -1)
            End If

            lua_settop L, -2
        Next i

        TableToVariant = arr2D
    Else
        ' 一维数组（转为单行二维）
        Dim arr1D() As Variant
        ReDim arr1D(1 To 1, 1 To CLng(length))

        For i = 1 To CLng(length)
            lua_rawgeti L, idx, CLng(i)
            arr1D(1, i) = GetValue(L, -1)
            lua_settop L, -2
        Next i

        TableToVariant = arr1D
    End If

    Exit Function

ErrorHandler:
    TableToVariant = "#TABLE_ERROR: " & Err.Description
End Function

' 辅助函数：将表转换为字典数组
Private Function TableToDictArray(ByVal L As LongPtr, ByVal idx As Long) As Variant
    On Error GoTo ErrorHandler

    ' 标准化索引
    If idx < 0 Then
        idx = lua_gettop(L) + idx + 1
    End If

    ' 第一遍：计数
    Dim count As Long
    count = 0

    Dim topBefore As Long
    topBefore = lua_gettop(L)

    lua_pushnil L
    Do While lua_next(L, idx) <> 0
        count = count + 1
        lua_settop L, -2  ' 弹出value，保留key
    Loop

    lua_settop L, topBefore

    If count = 0 Then
        TableToDictArray = Empty
        Exit Function
    End If

    ' 第二遍：提取数据
    Dim result() As Variant
    ReDim result(1 To count, 1 To 2)

    Dim i As Long
    i = 1

    lua_pushnil L
    Do While lua_next(L, idx) <> 0
        ' 获取键（在栈顶-1）
        Dim keyType As Long
        keyType = lua_type(L, -2)

        Select Case keyType
            Case LUA_TSTRING
                result(i, 1) = GetStringFromState(L, -2)
            Case LUA_TNUMBER
                result(i, 1) = lua_tonumberx(L, -2, 0)
            Case LUA_TBOOLEAN
                result(i, 1) = (lua_toboolean(L, -2) <> 0)
            Case Else
                result(i, 1) = "#KEY_TYPE_" & keyType
        End Select

        ' 获取值（在栈顶）
        result(i, 2) = GetValue(L, -1)

        i = i + 1
        lua_settop L, -2  ' 弹出value，保留key用于下次迭代
    Loop

    lua_settop L, topBefore

    TableToDictArray = result
    Exit Function

ErrorHandler:
    TableToDictArray = "#DICT_ERROR: " & Err.Description
End Function

' 解析 yield/return 字典
Private Sub ParseYieldReturn(task As TaskUnit, data As Variant, isFinal As Boolean)
    On Error Resume Next
    ' 如果不是数组,直接作为value处理
    If Not IsArray(data) Then
        task.taskValue = data
        Exit Sub
    End If

    ' 检查是否为字典格式(二维数组,第二维为2列)
    Dim isDictionary As Boolean
    isDictionary = False

    On Error Resume Next
    Dim cols As Long
    cols = UBound(data, 2) - LBound(data, 2) + 1
    If Err.Number = 0 And cols = 2 Then
        isDictionary = True
    End If
    On Error GoTo 0

    ' 如果是字典格式,解析键值对
    If isDictionary Then
        Dim i As Long
        For i = LBound(data, 1) To UBound(data, 1)
            Dim key As String
            Dim value As Variant

            key = LCase(Trim(CStr(data(i, 1))))
            value = data(i, 2)

            Select Case key
                Case "status"
                    ' 只有在非final或者值不是CO_DONE时才更新status
                    Dim statusVal As String
                    statusVal = LCase(Trim(CStr(value)))
                    If Not isFinal Then
                        ' yield时,根据返回的status字段决定协程状态
                        Select Case statusVal
                            Case "yield"
                                SetTaskStatus task, CO_YIELD
                            Case "done"
                                SetTaskStatus task, CO_DONE
                            Case "error"
                                SetTaskStatus task, CO_ERROR
                            Case Else
                                SetTaskStatus task, CO_YIELD ' 默认为yielded
                        End Select
                    End If
                Case "progress"
                    On Error Resume Next
                    task.taskProgress = CDbl(value)
                    On Error GoTo 0
                Case "message"
                    task.taskMessage = value
                Case "value"
                    task.taskValue = value
                Case "write"
                    ' 动态写入目标会在写入函数中处理
            End Select
        Next
    Else
        ' 如果不是字典格式,整个数组作为value
        task.taskValue = data
    End If
End Sub

' 根据调用单元格地址查找已存在的任务
Private Function FindTaskByCell(taskCell As String) As String
    Dim taskId As Variant
    If g_Tasks Is Nothing Then Exit Function
    For Each taskId In g_Tasks.Keys
        If g_Tasks(taskId).taskCell = taskCell Then
            FindTaskByCell = CStr(taskId)
            Exit Function
        End If
    Next
    FindTaskByCell = vbNullString
End Function

' ===== 辅助函数（新增）=====

' 检查 Collection 中是否存在指定任务
Private Function TaskQueueExists(task As TaskUnit) As Boolean
    On Error Resume Next
    Dim t As Variant
    For Each t In g_TaskQueue
        If t Is task Then
            TaskQueueExists = True
            Exit Function
        End If
    Next
    TaskQueueExists = False
End Function
' 安全地从 Collection 中移除元素
Private Sub TaskQueueRemove(task As TaskUnit)
    On Error Resume Next
    Dim i As Long
    For i = g_TaskQueue.Count To 1 Step -1
        If g_TaskQueue(i) Is task Then
            g_TaskQueue.Remove i
            Exit For
        End If
    Next
End Sub
' 安全地向 Collection 添加元素（避免重复）
Private Sub TaskQueueAdd(task As TaskUnit)
    If Not TaskQueueExists(task) Then
        g_TaskQueue.Add task
    End If
End Sub
' 状态转换验证函数
Private Function CanTransition(fromStatus As CoStatus, toStatus As CoStatus) As Boolean
    Select Case fromStatus
        Case CO_DEFINED
            CanTransition = (toStatus = CO_YIELD Or toStatus = CO_ERROR Or toStatus = CO_TERMINATED)
        Case CO_YIELD
            CanTransition = (toStatus = CO_DONE Or toStatus = CO_ERROR Or _
                           toStatus = CO_PAUSED Or toStatus = CO_TERMINATED)
        Case CO_PAUSED
            CanTransition = (toStatus = CO_YIELD Or toStatus = CO_TERMINATED)
        Case CO_DONE
            CanTransition = (toStatus = CO_DEFINED Or toStatus = CO_TERMINATED)
        Case CO_ERROR
            CanTransition = (toStatus = CO_DEFINED Or toStatus = CO_TERMINATED)
        Case CO_TERMINATED
            CanTransition = False
        Case Else
            CanTransition = False
    End Select
End Function
' 统一设置任务状态并处理副作用
Private Sub SetTaskStatus(task As TaskUnit, newStatus As CoStatus)
    If task Is Nothing Then Exit Sub

    Dim oldStatus As CoStatus
    oldStatus = task.taskStatus

    ' 相同状态不处理
    If oldStatus = newStatus Then Exit Sub

    Dim taskIdStr As String
    taskIdStr = "Task_" & task.taskId

    ' 状态转换验证
    If Not CanTransition(oldStatus, newStatus) Then
        LogError "非法状态转换: " & StatusToString(oldStatus) & " -> " & StatusToString(newStatus) & " (Task: " & taskIdStr & ")"
        Exit Sub
    End If

    ' 更新状态
    task.taskStatus = newStatus

    ' 根据新状态处理
    Select Case newStatus
        Case CO_DONE, CO_ERROR
            ' 完成或错误：释放协程但保留任务数据
            ReleaseTaskCoroutine task
            TaskQueueRemove task

        Case CO_TERMINATED
            ' 终止：释放协程并从队列移除
            ReleaseTaskCoroutine task
            TaskQueueRemove task

        Case CO_PAUSED
            ' 暂停：从队列移除但保留协程
            TaskQueueRemove task

        Case CO_YIELD
            ' 从 PAUSED 恢复时，重置 CFS 参数
            If oldStatus = CO_PAUSED Then
                task.CFS_vruntime = g_CFS_minVruntime
                task.CFS_lastScheduled = GetTickCount()
            End If

        Case CO_DEFINED
            ' ★ 修复：重置任务到初始状态
            ReleaseTaskCoroutine task
            TaskQueueRemove task
            ' 重置任务属性
            task.taskProgress = 0
            task.taskMessage = vbNullString
            task.taskValue = Empty
            task.taskError = vbNullString
            task.taskLastTime = 0
            task.taskTotalTime = 0
            task.taskTickCount = 0
            task.CFS_vruntime = 0
            task.CFS_weight = CFS_DEFAULT_WEIGHT
            task.CFS_lastScheduled = 0
    End Select

    ' 标记监控为脏
    g_StateDirty = True
    MarkWatchesDirty task

    LogDebug "状态转换: " & taskIdStr & " " & StatusToString(oldStatus) & " -> " & StatusToString(newStatus)
End Sub
' 统一释放任务的协程资源
Public Sub ReleaseTaskCoroutine(task As TaskUnit)
    On Error Resume Next

    If task Is Nothing Then Exit Sub
    If task.taskCoRef = 0 Then Exit Sub

    ' 确保 Lua 状态机有效
    If g_LuaState = 0 Or Not g_Initialized Then
        ' Debug.Print "ReleaseTaskCoroutine: Lua 状态机无效，跳过释放"
        task.ClearCoroutineRef
        Exit Sub
    End If

    ' 执行释放
    ' Debug.Print "ReleaseTaskCoroutine: Task_" & task.taskId & " 释放协程 Ref=" & task.taskCoRef
    luaL_unref g_LuaState, LUA_REGISTRYINDEX, task.taskCoRef

    ' 清除任务中的引用
    task.ClearCoroutineRef
End Sub
' 检查 Collection 中是否存在指定键
Private Function CollectionExists(col As Collection, key As String) As Boolean
    On Error Resume Next
    Dim dummy As Variant
    dummy = col(key)
    CollectionExists = (Err.Number = 0)
    Err.Clear
End Function
' 安全地从 Collection 中移除元素（按键）
Private Sub CollectionRemove(col As Collection, key As String)
    On Error Resume Next
    col.Remove key
    Err.Clear
End Sub
' 辅助函数：状态转字符串
Private Function StatusToString(status As CoStatus) As String
    Select Case status
        Case CO_DEFINED: StatusToString = "DEFINED"
        Case CO_YIELD: StatusToString = "YIELD"
        Case CO_PAUSED: StatusToString = "PAUSED"
        Case CO_DONE: StatusToString = "DONE"
        Case CO_ERROR: StatusToString = "ERROR"
        Case CO_TERMINATED: StatusToString = "TERMINATED"
        Case Else: StatusToString = "UNKNOWN(" & status & ")"
    End Select
End Function
' 清理孤儿 Watch（任务已删除但 Watch 仍存在）
Public Sub CleanupOrphanWatches()
    On Error Resume Next
    
    If g_Watches Is Nothing Then Exit Sub
    If g_Watches.Count = 0 Then Exit Sub
    
    Dim toRemove As Collection
    Set toRemove = New Collection
    
    Dim watchCell As Variant
    Dim wi As WatchInfo
    
    ' 第一遍：收集孤儿
    For Each watchCell In g_Watches.Keys
        Set wi = g_Watches(watchCell)
        
        ' 检查任务是否存在
        If Not g_Tasks.Exists(wi.watchTaskId) Then
            toRemove.Add CStr(watchCell)
        ' 检查 watchTask 引用是否有效
        ElseIf wi.watchTask Is Nothing Then
            toRemove.Add CStr(watchCell)
        End If
    Next
    
    ' 第二遍：移除孤儿
    Dim removeKey As Variant
    For Each removeKey In toRemove
        g_Watches.Remove CStr(removeKey)
    Next
    
    If toRemove.Count > 0 Then
        LogInfo "清理了 " & toRemove.Count & " 个孤儿Watch"
    End If
End Sub
' 安全删除任务（统一入口）
Public Sub SafeDeleteTask(taskId As String)
    On Error GoTo ErrorHandler
    
    If Not g_Tasks.Exists(taskId) Then
        LogDebug "任务不存在，无需删除: " & taskId
        Exit Sub
    End If
    
    Dim task As TaskUnit
    Set task = g_Tasks(taskId)
    
    ' 1. 先清理该任务的所有 Watch
    CleanupTaskWatches task
    
    ' 2. 释放协程
    ReleaseTaskCoroutine task
    
    ' 3. 从队列移除
    TaskQueueRemove task
    
    ' 4. 释放任务对象
    Set g_Tasks(taskId) = Nothing
    g_Tasks.Remove taskId
    
    LogDebug "任务已安全删除: " & taskId
    Exit Sub
    
ErrorHandler:
    LogError "删除任务失败: " & taskId & " - " & Err.Description
End Sub
' ★ 新增辅助函数：从外部地址解析工作表
Private Function GetWorksheetFromAddress(addr As String, wb As Workbook) As Worksheet
    On Error Resume Next

    ' addr 格式示例: "[Book1.xlsx]Sheet1!$A$1" 或 "'Sheet Name'!$A$1"
    Dim sheetName As String
    Dim exclamPos As Long
    Dim bracketEnd As Long

    Set GetWorksheetFromAddress = Nothing

    ' 处理外部引用格式 [Book1.xlsx]Sheet1!$A$1
    If Left(addr, 1) = "[" Then
        bracketEnd = InStr(addr, "]")
        exclamPos = InStr(addr, "!")
        If exclamPos > bracketEnd Then
            sheetName = Mid(addr, bracketEnd + 1, exclamPos - bracketEnd - 1)
        End If
    ElseIf InStr(addr, "!") > 0 Then
        exclamPos = InStr(addr, "!")
        sheetName = Left(addr, exclamPos - 1)
    End If

    ' 移除引号
    sheetName = Replace(sheetName, "'", "")

    If Len(sheetName) > 0 Then
        Set GetWorksheetFromAddress = wb.Sheets(sheetName)
    End If

    ' 如果解析失败，使用活动工作表
    If GetWorksheetFromAddress Is Nothing Then
        Set GetWorksheetFromAddress = wb.ActiveSheet
    End If
End Function
' ============================================
' 清理特定工作簿的监视点
' ============================================
Public Sub CleanupWorkbookWatches(wbName As String)
    On Error Resume Next
    
    If g_Watches Is Nothing Then Exit Sub
    If g_Watches.Count = 0 Then Exit Sub
    
    ' 收集要移除的 watch
    Dim toRemove As Collection
    Set toRemove = New Collection
    
    Dim watchCell As Variant
    Dim watchInfo As WatchInfo
    
    For Each watchCell In g_Watches.Keys
        Set watchInfo = g_Watches(watchCell)
        If watchInfo.watchWorkbook = wbName Then
            toRemove.Add CStr(watchCell)
        End If
    Next
    
    ' 移除收集到的 watch
    Dim removeKey As Variant
    For Each removeKey In toRemove
        If g_Watches.Exists(CStr(removeKey)) Then
            Set watchInfo = g_Watches(CStr(removeKey))
            ' 从任务的 taskWatches 集合中移除
            If Not watchInfo.watchTask Is Nothing Then
                RemoveFromTaskWatches watchInfo.watchTask, CStr(removeKey)
            End If
            ' 从主索引移除
            g_Watches.Remove CStr(removeKey)
        End If
    Next
    
    If toRemove.Count > 0 Then
        LogDebug "已清理工作簿 " & wbName & " 的 " & toRemove.Count & " 个监控"
    End If
End Sub