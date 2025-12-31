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
Private g_LuaState As LongPtr
Private g_Initialized As Boolean
Private g_HotReloadEnabled As Boolean
Private g_FunctionsPath As String  ' 固定为加载项目录
Private g_LastModified As Date
' ===== 协程全局变量 =====
Private g_Tasks As Object       ' task Id -> task Instance
Public g_Workbooks As Object    ' Dictionary: wbName -> WorkbookInfo
Private g_TaskQueue As Object     ' taskId -> True (active tasks)
' ===== 调度全局变量 =====
Private g_SchedulerRunning As Boolean   ' 调度器是否运行中
Private g_SchedulerCursorByTask As Long ' Round-Robin 游标
Private g_StateDirty As Boolean         ' 本 tick 是否有状态变化，用来检测是否需要刷新单元格
Private g_NextTaskId As Long            ' 新建下一个任务ID计数器
Private g_SchedulerIntervalMilliSec As Long ' 调度间隔(ms)
Private g_NextScheduleTime As Date     '标记记下一次调度时间

Private g_ScheduleMode As Integer         ' 0=按任务顺序, 1=按工作簿
Private g_MaxIterationsPerTick As Integer ' 按任务时调度：每次调度迭代次数
Private g_WorkbookTicks As Integer        ' 按工作簿调度：默认每个工作簿的tick数
' ===== 配置常量 =====
Private Const CP_UTF8 As Long = 65001
Private Const DEFAULT_HOT_RELOAD_ENABLED As Boolean = True
Private Const SCHEDULER_INTERVAL_Milli_SEC As Long = 1000  ' 调度间隔，默认1000ms
Private Const DEFAULT_MAX_ITERATIONS_PER_TICK As Long = 1  ' 每次调度迭代次数，默认1
Private Const DEFAULT_SCHEDULER_MODE As Integer = 0  ' 调度模式：0=按任务顺序, 1=按工作簿
Private Const DEFAULT_WORKBOOK_TICKS As Integer = 1  ' 每个工作簿的默认tick数
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

    g_Initialized = True
    g_HotReloadEnabled = DEFAULT_HOT_RELOAD_ENABLED

    InitCoroutineSystem

    ' 尝试加载functions.lua
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(g_FunctionsPath) Then
        If Not TryLoadFunctionsFile() Then
            MsgBox "functions.lua 加载失败。" & vbCrLf & _
                   "Lua 引擎已启动,但自定义函数不可用。", _
                   vbExclamation, "InitLuaState_Warning"
        End If
    End If

    InitLuaState = True
    ' MsgBox "Lua栈初始化完成",vbInformation,"InitLuaState_Info" 
    Exit Function

ErrorHandler:
    MsgBox "初始化 Lua 失败: " & Err.Description, vbCritical, "严重错误"
    InitLuaState = False
End Function

' 初始化协程系统
Private Sub InitCoroutineSystem()
    g_MaxIterationsPerTick = DEFAULT_MAX_ITERATIONS_PER_TICK
    g_SchedulerIntervalMilliSec = SCHEDULER_INTERVAL_Milli_SEC
    g_ScheduleMode = DEFAULT_SCHEDULER_MODE ' 默认按任务顺序调度
    g_WorkbookTicks = DEFAULT_WORKBOOK_TICKS ' 按工作簿调度，默认每个工作簿1个tick

    ' 初始化性能统计
    g_SchedulerStats.TotalTime = 0
    g_SchedulerStats.LastTime = 0
    g_SchedulerStats.TotalCount = 0
    g_SchedulerStats.StartTime = Now
    If g_Workbooks Is Nothing Then Set g_Workbooks = CreateObject("Scripting.Dictionary")
    If g_Tasks Is Nothing Then Set g_Tasks = CreateObject("Scripting.Dictionary")
    If g_TaskQueue Is Nothing Then Set g_TaskQueue = CreateObject("Scripting.Dictionary")

    If g_NextTaskId = 1 Then g_NextTaskId = 1
    g_SchedulerCursorByTask = 0
    g_StateDirty = False
End Sub

' 清理 Lua 状态机
Public Sub CleanupLua()
    If g_Initialized Then
        g_SchedulerRunning = False
        StopScheduler

        If Not g_Tasks Is Nothing Then
            g_Tasks.RemoveAll
            g_TaskQueue.RemoveAll
        End If

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

    ' 关键修复:检查调用者工作簿是否为宏文件
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
        wbInfo.WbName = wbName
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
    Dim task As New TaskInfo
    task.taskId = g_NextTaskId
    task.taskFunc = funcName
    task.taskWorkbook = wbName
    task.taskStartArgs = startArgs
    task.taskResumeSpec = resumeSpec
    task.taskCell = taskCell
    task.taskStatus = "defined"
    task.taskProgress = 0
    task.taskMessage = vbNullString
    task.taskValue = vbNull
    task.taskError = vbNullString
    task.taskCoThread = 0
    task.taskLastTime = 0
    task.taskTotalTime = 0
    task.taskTickCount = 0
    g_Tasks.Add taskId, task
    g_Workbooks(wbName).AddTask taskId, task

    LuaTask = taskId
    g_NextTaskId = g_NextTaskId + 1
    Exit Function
ErrorHandler:
    Dim errorDetails As String
    errorDetails = "Task错误:" & vbCrLf
    errorDetails = errorDetails & "错误号: " & Err.Number & vbCrLf
    errorDetails = errorDetails & "描述: " & Err.Description & vbCrLf
    ' If Err.Erl <> 0 Then
    '     errorDetails = errorDetails & "行号: " & Err.Erl & vbCrLf
    ' End If
    ' LuaTask = "#ERROR: " & errorDetails
    ' 'errorDetails = errorDetails & "工作簿: " & wbName & vbCrLf
    ' 'errorDetails = errorDetails & "对比名称: " & g_Workbooks(wbName).Name & vbCrLf
    ' ' 输出到立即窗口便于调试
    Debug.Print "=== Task错误详情 ==="
    Debug.Print errorDetails
    Debug.Print "======================="
    LuaTask = "#ERROR: Task" & Err.Description
End Function

' 读取任务状态
Public Function LuaGet(taskId As String, field As String) As Variant
    On Error GoTo ErrorHandler

    ' 标记为 volatile，每次计算都会刷新
    Application.Volatile True

    If Not InitLuaState() Then
        LuaGet = CVErr(xlErrValue)
        Exit Function
    End If

    If Not g_Tasks.Exists(taskId) Then
        LuaGet = "#ERROR: 任务不存在"
        Exit Function
    End If
    Dim task As TaskInfo
    Set task = g_Tasks(taskId)

    Select Case LCase(field)
        Case "status"
            LuaGet = task.taskStatus
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
            summary = "状态:" & task.taskStatus
            summary = summary & " | 进度:" & Format(task.taskProgress, "0.0") & "%"
            If task.taskStatus = "error" Then
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
' ============================================
' 第五部分：协程执行和调度
' ============================================
' 启动协程
Public Sub StartLuaCoroutine(taskId As String)
    On Error GoTo ErrorHandler
    If g_Tasks Is Nothing Then
        InitCoroutineSystem
    End If

    If Not g_Tasks.Exists(taskId) Then
        MsgBox "错误：任务 " & taskId & " 不存在", vbCritical
        Exit Sub
    End If

    Dim task As TaskInfo
    Set task = g_Tasks(taskId)

    If task.taskStatus <> "defined" Then
        MsgBox "错误：任务已启动或已完成", vbExclamation
        Exit Sub
    End If
    If g_LuaState = 0 Then
        MsgBox "Lua主状态未初始化", vbCritical
        Exit Sub
    End If

    Dim coThread As LongPtr
    coThread = lua_newthread(g_LuaState)
    If coThread = 0 Then
        task.taskStatus = "error"
        task.taskError = "无法创建协程线程"
        Exit Sub
    End If
    task.taskCoThread = coThread

    lua_getglobal g_LuaState, task.taskFunc

    If lua_type(g_LuaState, -1) <> LUA_TFUNCTION Then
        task.taskStatus = "error"
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

    HandleCoroutineResult task, result, CLng(nres)
    If task.taskStatus = "yielded" Then
        g_TaskQueue(taskId) = True
    End If

    Exit Sub
ErrorHandler:
    task.taskStatus = "error"
    task.taskError = "VBA错误: " & Err.Description & " (行 " & Erl & ")"
    MsgBox "启动协程失败: " & Err.Description, vbCritical
End Sub

' 启动调度器
Private Sub StartSchedulerIfNeeded()
    If g_SchedulerRunning Then Exit Sub
    If g_TaskQueue Is Nothing Then Exit Sub
    If g_TaskQueue.Count = 0 Then Exit Sub
    g_SchedulerRunning = True
    g_NextScheduleTime = Now + g_SchedulerIntervalMilliSec / 86400000#
    Application.OnTime g_NextScheduleTime, "SchedulerTick" 
End Sub

' 调度器心跳 - 主入口
Public Sub SchedulerTick()
    On Error Resume Next
    If Not g_SchedulerRunning Then Exit Sub
    If g_TaskQueue Is Nothing Or g_TaskQueue.Count = 0 Then
        g_SchedulerRunning = False
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim schedulerStart As Double
    schedulerStart = GetTickCount()

    ' 根据调度模式分流执行
    If g_ScheduleMode = 0 Then
        Call ScheduleByTask  ' 按任务顺序调度
    Else
        Call ScheduleByWorkbook  ' 按工作簿调度
    End If

    ' 性能计时统计
    Dim schedulerElapsed As Double
    schedulerElapsed = GetTickCount() - schedulerStart
    g_SchedulerStats.LastTime = schedulerElapsed
    g_SchedulerStats.TotalTime = g_SchedulerStats.TotalTime + schedulerElapsed
    g_SchedulerStats.TotalCount = g_SchedulerStats.TotalCount + 1

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If g_StateDirty Then
        g_StateDirty = False
        ActiveSheet.Calculate
    End If

    ' 重新安排下一次调度
    If g_TaskQueue.Count > 0 Then
        g_NextScheduleTime = Now + g_SchedulerIntervalMilliSec / 86400000#
        Application.OnTime g_NextScheduleTime, "SchedulerTick"
    Else
        g_SchedulerRunning = False
    End If
End Sub

' 按任务顺序调度 - 直接执行
Private Sub ScheduleByTask()
    ' 构建任务ID数组
    Dim taskIds() As Variant
    ReDim taskIds(0 To g_TaskQueue.Count - 1)

    Dim idx As Long, taskId As Variant
    idx = 0
    For Each taskId In g_TaskQueue.Keys
        taskIds(idx) = taskId
        idx = idx + 1
    Next

    Dim total As Long
    total = UBound(taskIds) + 1
    If total = 0 Then Exit Sub

    ' Round-Robin 直接执行
    Dim executed As Long, cur As Long
    Dim task As TaskInfo
    Dim taskStart As Double, taskElapsed As Double
    Dim wb As String

    cur = g_SchedulerCursorByTask Mod total

    Do While executed < g_MaxIterationsPerTick And executed < total
        taskId = CStr(taskIds(cur))

        If g_Tasks.Exists(taskId) Then
            Set task = g_Tasks(taskId)

            ' 只调度 yielded 状态的任务
            If task.taskStatus = "yielded" Then
                ' 任务级别计时开始
                taskStart = GetTickCount()
                wb = task.taskWorkbook

                ' 直接执行任务
                ResumeCoroutine task

                ' 任务级别计时结束
                taskElapsed = GetTickCount() - taskStart

                ' 更新工作簿统计
                g_Workbooks(wb).LastTime = taskElapsed
                g_Workbooks(wb).TotalTime = g_Workbooks(wb).TotalTime + taskElapsed
                g_Workbooks(wb).TickCount = g_Workbooks(wb).TickCount + 1
                ' 检查任务是否完成
                Dim status As String
                status = task.taskStatus
                If status = "done" Or status = "error" Or status = "terminated" Then
                    g_TaskQueue.Remove taskId
                End If
            End If
        Else
            g_TaskQueue.Remove taskId
        End If

        executed = executed + 1
        cur = (cur + 1) Mod total
    Loop

    ' 更新游标
    g_SchedulerCursorByTask = cur

End Sub

' 按工作簿调度 - 直接执行
Private Sub ScheduleByWorkbook()
    On Error GoTo ErrorHandler

    If g_Workbooks Is Nothing Then Exit Sub
    If g_Workbooks.Count = 0 Then Exit Sub

    Dim wb As WorkbookInfo
    Dim task As TaskInfo
    Dim taskId As String
    Dim taskStart As Double, taskElapsed As Double

    ' 遍历所有工作簿
    For Each wb In g_Workbooks.Items

        ' 确定此工作簿的tick数
        Dim tickCount As Integer
        If wb.WbAllowedTickCount > -1 Then
            tickCount = wb.WbAllowedTickCount
        Else
            tickCount = g_WorkbookTicks
        End If

        ' 如果没有任务或tick数为0，跳过
        If wb.Tasks.Count = 0 Or tickCount <= 0 Then GoTo NextWorkbook

        ' 获取任务列表
        Dim taskList As Variant
        taskList = wb.Tasks.Keys
        Dim totalTasks As Long
        totalTasks = UBound(taskList) + 1

        ' Round-Robin 调度
        Dim cursor As Integer
        cursor = wb.WbCursor

        Dim executedCount As Long
        Dim stepCount As Long
        executedCount = 0
        stepCount = 0

        Do While executedCount < tickCount And stepCount < totalTasks
            taskId = CStr(taskList(cursor))

            ' 检查任务是否仍在全局队列中
            If g_Tasks.Exists(taskId) Then
                Set task = g_Tasks(taskId)

                ' 只执行 yielded 状态的任务
                If task.taskStatus = "yielded" Then
                    ' 性能计时开始
                    taskStart = GetTickCount()

                    ' 执行任务
                    ResumeCoroutine task

                    ' 性能计时结束
                    taskElapsed = GetTickCount() - taskStart

                    ' 更新工作簿统计
                    wb.WbLastTime = taskElapsed
                    wb.WbTotalTime = wb.WbTotalTime + taskElapsed
                    wb.WbTickCount = wb.WbTickCount + 1

                    ' 更新任务统计
                    task.taskLastTime = taskElapsed
                    task.taskTotalTime = task.taskTotalTime + taskElapsed
                    task.taskTickCount = task.taskTickCount + 1

                    executedCount = executedCount + 1

                    ' 检查任务是否完成
                    Dim status As String
                    status = task.taskStatus
                    If status = "done" Or status = "error" Or status = "terminated" Then
                        g_TaskQueue.Remove taskId
                        wb.Tasks.Remove taskId  ' 从工作簿任务列表移除
                    End If
                End If
            Else
                ' 任务不存在，从队列和工作簿移除
                If g_TaskQueue.Exists(taskId) Then g_TaskQueue.Remove taskId
                If wb.Tasks.Exists(taskId) Then wb.Tasks.Remove taskId
            End If

            cursor = (cursor + 1) Mod totalTasks
            stepCount = stepCount + 1
        Loop

        ' 保存游标
        wb.WbCursor = cursor

NextWorkbook:
    Next wb
    Exit Sub
ErrorHandler:
    Debug.Print "ScheduleByWorkbook错误: " & Err.Description
End Sub

' Resume 协程
Private Sub ResumeCoroutine(task As TaskInfo)
    On Error GoTo ErrorHandler

    If task.taskStatus <> "yielded" Then
        Exit Sub
    End If

    ' 性能计时开始
    Dim taskStart As Long
    taskStart = GetTickCount()

    ' 检查协程线程是否有效
    Dim coThread As LongPtr
    coThread = task.taskCoThread

    If coThread = 0 Then
        task.taskStatus = "error"
        task.taskError = "协程线程无效(coThread=0)"
        Exit Sub
    End If

    ' 检查协程状态
    Dim coStatus As Long
    coStatus = lua_status(coThread)
    If coStatus <> LUA_OK And coStatus <> LUA_YIELD Then
        task.taskStatus = "error"
        task.taskError = "协程状态异常: " & coStatus & " (期望: " & LUA_YIELD & ")"
        Exit Sub
    End If

    ' 检查工作簿是否仍然打开
    Dim wbName As String
    wbName = task.taskWorkbook

    If Not IsEmpty(wbName) And wbName <> vbNullString And wbName <> "" Then
        Dim wb As Workbook
        Dim wbExists As Boolean
        wbExists = False

        On Error Resume Next
        Set wb = Application.Workbooks(wbName)
        wbExists = Not (wb Is Nothing)
        On Error GoTo ErrorHandler

        If Not wbExists Then
            task.taskStatus = "error"
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

            ' 处理字符串参数(可能是单元格地址)
            If VarType(param) = vbString Then
                Dim paramStr As String
                paramStr = Trim(CStr(param))

                ' 尝试作为单元格地址解析
                If Len(paramStr) > 0 And Not wb Is Nothing Then
                    On Error Resume Next
                    Dim rng As Range
                    Set rng = Nothing
                    Set rng = wb.Range(paramStr)

                    If Err.Number = 0 And Not rng Is Nothing Then
                        ' 成功解析为单元格,传递值
                        If rng.Cells.Count = 1 Then
                            PushValue coThread, rng.Value
                        Else
                            PushValue coThread, rng.Value
                        End If
                    Else
                        ' 不是有效的单元格地址,作为普通字符串
                        PushValue coThread, paramStr
                    End If
                    Err.Clear
                    On Error GoTo ErrorHandler
                Else
                    ' 空字符串或无工作簿
                    PushValue coThread, paramStr
                End If
            Else
                ' 非字符串参数直接传递
                PushValue coThread, param
            End If

            nargs = nargs + 1
        Next
    End If

    ' 执行 resume
    Dim nres As LongPtr
    Dim result As Long
    result = lua_resume(coThread, g_LuaState, nargs, VarPtr(nres))

    ' 处理结果
    HandleCoroutineResult task, result, CLng(nres)

    ' 性能计时结束并统计
    Dim taskElapsed As Double
    taskElapsed = GetTickCount() - taskStart

    ' 更新任务统计
    task.taskLastTime = taskElapsed
    task.taskTotalTime = task.taskTotalTime + taskElapsed
    task.taskTickCount = task.taskTickCount + 1
    Exit Sub
ErrorHandler:
    Dim errorDetails As String
    errorDetails = "Resume错误:" & vbCrLf
    errorDetails = errorDetails & "错误号: " & Err.Number & vbCrLf
    errorDetails = errorDetails & "描述: " & Err.Description & vbCrLf

    If Err.Erl <> 0 Then
        errorDetails = errorDetails & "行号: " & Err.Erl & vbCrLf
    End If

    errorDetails = errorDetails & "协程线程: " & coThread & vbCrLf
    errorDetails = errorDetails & "工作簿: " & wbName & vbCrLf
    errorDetails = errorDetails & "参数数量: " & nargs

    task.taskStatus = "error"
    task.taskError = errorDetails

    ' 输出到立即窗口便于调试
    Debug.Print "=== Resume 错误详情 ==="
    Debug.Print errorDetails
    Debug.Print "======================="
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
Private Sub HandleCoroutineResult(task As TaskInfo, result As Long, nres As Long)
    On Error GoTo ErrorHandler

    Dim coThread As LongPtr
    coThread = task.taskCoThread

    Dim stackTop As Long
    stackTop = lua_gettop(coThread)  ' 入口保存

    Select Case result
        Case LUA_OK
            task.taskStatus = "done"
            g_StateDirty = True
            task.taskProgress = 100

            If nres > 0 And stackTop > 0 Then
                Dim retData As Variant
                retData = GetValue(coThread, -1)
                ParseYieldReturn task, retData, True
            End If

        Case LUA_YIELD
            If nres > 0 And stackTop > 0 Then
                Dim yieldData As Variant
                yieldData = GetValue(coThread, -1)
                ParseYieldReturn task, yieldData, False
            End If
            If task.taskStatus <> "done" And task.taskStatus <> "error" Then
                task.taskStatus = "yielded"
            End If
            g_StateDirty = True

        Case Else
            task.taskStatus = "error"
            g_StateDirty = True

            If nres > 0 And stackTop > 0 Then
                task.taskError = GetStringFromState(coThread, -1)
            Else
                task.taskError = "协程错误: 代码 " & result
            End If
    End Select

    lua_settop coThread, stackTop  ' 统一恢复
    Exit Sub

ErrorHandler:
    task.taskStatus = "error"
    task.taskError = "处理结果错误: " & Err.Description
    If coThread <> 0 Then lua_settop coThread, stackTop
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
Private Sub ParseYieldReturn(task As TaskInfo, data As Variant, isFinal As Boolean)
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
                    ' 只有在非final或者值不是"done"时才更新status
                    Dim statusVal As String
                    statusVal = LCase(Trim(CStr(value)))
                    If Not isFinal Then
                        ' yield时,根据返回的status字段决定协程状态
                        Select Case statusVal
                            Case "yielded", "done", "error"
                                task.taskStatus = statusVal
                            Case Else
                                task.taskStatus = "yielded" ' 默认为yielded
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
        If g_Tasks(taskId).Cell = taskCell Then
            FindTaskByCell = CStr(taskId)
            Exit Function
        End If
    Next
    FindTaskByCell = vbNullString
End Function
