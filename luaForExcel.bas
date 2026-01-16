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
    Private Declare PtrSafe Sub lua_pushlstring Lib "lua54.dll" (ByVal L As LongPtr, ByVal s As LongPtr, ByVal strLen As LongPtr)
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
    Private Declare PtrSafe Function luaL_loadbufferx Lib "lua54.dll" (ByVal L As LongPtr, ByVal buff As LongPtr, ByVal sz As LongPtr, ByVal name As LongPtr, ByVal mode As LongPtr) As Long
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
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
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
Private Const LUA_REFNIL As Long = -1 ' luaL_ref 对 nil 值的返回
Private Const LUA_NOREF As Long = -2 ' 无效引用
' ===== 参数规格类型常量（供 TaskUnit 使用）=====
Public Const PARAM_LITERAL As Long = 0    ' 字面量（数值、布尔、普通字符串）
Public Const PARAM_REF As Long = 1        ' 单元格区域引用（Range 对象传入）,引用类型（level字段决定解引用层数）
' ===== 全局变量 =====
Private g_LuaState As LongPtr         ' lua 栈
Private g_Initialized As Boolean      ' 是否初始化
Private g_HotReloadEnabled As Boolean ' 是否启用 hot-reload，默认开启
Private g_FunctionsPath As String     ' 固定为加载项目录
Private g_LastModified As Date        ' functions.lua 上次修改，用于热重载
Private g_CFS_autoWeight As Boolean   ' 自动调整权重开关
' 调度模式：0=单轮模式（每tick只运行一个任务），1=时间片模式（每tick运行x ms）
Private g_ScheduleMode As Long
Private Const SCHEDULE_MODE_SINGLE As Long = 0    ' 单轮模式
Private Const SCHEDULE_MODE_TIMESLICE As Long = 1 ' 时间片模式
' ===== 协程全局变量 =====
Public Enum CoStatus ' task 的状态枚举
    CO_DEFINED
    CO_YIELD
    CO_PAUSED
    CO_DONE
    CO_ERROR
    CO_TERMINATED
End Enum
Public g_Tasks As Object         ' task Id -> task Instance
Public g_Workbooks As Object     ' Dictionary: wbName -> WorkbookInfo
Public g_TaskQueue As Collection ' taskId 列表，按调度顺序排列
Public g_Watches As Object       ' Dictionary: watchCell -> WatchInfo
' ===== 调度全局变量 =====
Private g_SchedulerRunning As Boolean   ' 调度器是否运行中
Private g_StateDirty As Boolean         ' 本 tick 是否有状态变化，用来检测是否需要刷新单元格
Public g_NextTaskId As Integer          ' 新建下一个任务ID计数器
Private g_SchedulerIntervalMilliSec As Long ' 调度间隔(ms)，默认1000ms
Private g_NextScheduleTime As Date     '标记记下一次调度时间
Private g_CFS_minVruntime As Double       ' 队列中最小的 vruntime（用于新任务初始化）
Private g_CFS_targetLatency As Double     ' 目标延迟周期（ms），默认为调度间隔的十分之一
Private g_CFS_minGranularity As Double    ' 最小执行粒度（ms），默认 5ms
Private g_CFS_niceToWeight(0 To 39) As Double  ' nice 值到权重的映射表
Private g_ActiveTaskCount As Long  ' 当前队列中活跃任务数
' ===== 配置常量 =====
Private Const CP_UTF8 As Long = 65001
Private Const LOG_LEVEL As Byte = 1  ' 默认日志等级：0=错误，1=信息，2=调试
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
    g_HotReloadEnabled = True

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
    If g_SchedulerIntervalMilliSec = 0 Then g_SchedulerIntervalMilliSec = 1000
    ' CFS 参数初始化
    If g_CFS_minVruntime = 0 Then g_CFS_minVruntime = 0
    If g_CFS_targetLatency = 0 Then g_CFS_targetLatency = g_SchedulerIntervalMilliSec / 10  ' 默认为调度间隔的十分之一
    If g_CFS_minGranularity = 0 Then g_CFS_minGranularity = 5
    g_CFS_autoWeight = False  ' 默认关闭自动权重调整
    ' 调度模式初始化（默认单轮模式）
    If g_ScheduleMode = 0 Then g_ScheduleMode = SCHEDULE_MODE_SINGLE
    ' 初始化 nice 到权重的映射表（简化版，只用 0-39 对应 nice -20 到 +19）
    ' 权重公式: weight = 1024 / 1.25^nice  (nice=0 时 weight=1024)
    Dim i As Byte
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
    g_ActiveTaskCount = 0
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

    ' 检查是否为协程 yield 错误
    If result = LUA_YIELD Then
        LuaCall = "提示: 函数 '" & funcName & "' 是协程函数，请使用 LuaTask 调用而非 LuaCall"
        lua_settop g_LuaState, stackTop  ' 统一恢复
        Exit Function
    End If

    If result <> 0 Then
        Dim errMsg As String
        errMsg = GetStringFromState(g_LuaState, -1)

        ' 检查错误信息是否包含 yield 相关内容
        If InStr(1, errMsg, "yield", vbTextCompare) > 0 Or _
           InStr(1, errMsg, "coroutine", vbTextCompare) > 0 Then
            LuaCall = "提示: 函数 '" & funcName & "' 包含协程操作(yield)，请使用 LuaTask 调用" & vbCrLf & _
                      "用法: =LuaTask(""" & funcName & """, 参数...)"
        Else
            LuaCall = "运行错误: " & errMsg
        End If
        lua_settop g_LuaState, stackTop  ' 统一恢复
        Exit Function
    End If

    Dim nResults As Long
    nResults = lua_gettop(g_LuaState) - stackTop  ' 相对计算结果数

    If nResults = 0 Then
        LuaCall = Empty
    Else
        LuaCall = GetValue(g_LuaState, -nResults)
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
    ' 获取调用单元格地址和工作簿/工作表
    Dim taskCell As String
    Dim wbName As String
    Dim callerWb As Workbook
    Dim callerWs As Worksheet
    On Error Resume Next
    taskCell = Application.Caller.Address(External:=True)
    Set callerWb = Application.Caller.Worksheet.Parent
    Set callerWs = Application.Caller.Worksheet
    On Error GoTo ErrorHandler
    ' 检查调用者工作簿是否存在
    If callerWb Is Nothing Then
        LuaTask = "#ERROR: 无法获取调用工作簿"
        Exit Function
    End If
    ' 防止在xlam文件中创建任务
    If callerWb.FileFormat = xlAddIn Then
        LuaTask = "#ERROR: 不能在宏文件中创建任务"
        Exit Function
    End If
    wbName = callerWb.Name
    ' ========== 解析参数 ==========
    Dim funcName As String
    funcName = CStr(params(0))
    Dim startList As Object
    Dim resumeList As Object
    Set startList = CreateObject("System.Collections.ArrayList")
    Set resumeList = CreateObject("System.Collections.ArrayList")
    ' 遍历参数，用布尔标记区分阶段
    Dim isResumePhase As Boolean
    isResumePhase = False
    Dim i As Long
    Dim isSeparator As Boolean
    Dim rangeInfo As Object
    For i = 1 To UBound(params)
        If IsMissing(params(i)) Then
            ' 空参数位置，作为 nil 处理
            If Not isResumePhase Then
                startList.Add Empty  ' 启动参数阶段：添加 Empty（会被转换为 nil）
            Else
                resumeList.Add Empty
            End If
        Else
            ' 先检查是否是分隔符（必须在检查 Range 之前）
            isSeparator = False
            If Not IsObject(params(i)) Then
                If VarType(params(i)) = vbString Then
                    If CStr(params(i)) = "|" Then
                        isSeparator = True
                    End If
                End If
            End If
            If isSeparator Then
                isResumePhase = True
            ElseIf Not isResumePhase Then
                ' 启动参数：Range 取值，其他直接存储
                If TypeName(params(i)) = "Range" Then
                    startList.Add params(i).Value
                Else
                    startList.Add params(i)
                End If
            Else
                ' Resume 参数：Range 存储地址信息，其他直接存储
                If TypeName(params(i)) = "Range" Then
                    Set rangeInfo = CreateObject("Scripting.Dictionary")
                    rangeInfo("isRange") = True
                    rangeInfo("address") = params(i).Address(False, False)
                    rangeInfo("workbook") = params(i).Worksheet.Parent.Name
                    rangeInfo("worksheet") = params(i).Worksheet.Name
                    rangeInfo("cellCount") = params(i).Cells.Count
                    resumeList.Add rangeInfo
                Else
                    resumeList.Add params(i)
                End If
            End If
        End If
    Next i
    ' 转换为数组
    Dim startArgs As Variant
    Dim resumeSpec As Variant
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
    ' ========== 检查是否已存在任务 ==========
    Dim existingTaskId As String
    existingTaskId = FindTaskByCell(taskCell)
    If existingTaskId <> vbNullString Then
        Dim existingTask As TaskUnit
        Set existingTask = g_Tasks(existingTaskId)
        Dim paramsChanged As Boolean
        paramsChanged = False
        ' 校验函数名
        If existingTask.taskFunc <> funcName Then
            paramsChanged = True
        End If
        ' 校验启动参数
        If Not paramsChanged Then
            paramsChanged = Not CompareVariantArrays(existingTask.taskStartArgs, startArgs)
        End If
        ' 校验 Resume 参数规格（比较地址，不比较值）
        If Not paramsChanged Then
            paramsChanged = Not CompareResumeSpecs(existingTask.taskResumeSpec, resumeSpec)
        End If
        ' 如果参数未变化，直接返回现有任务ID
        If Not paramsChanged Then
            LuaTask = existingTaskId
            Exit Function
        End If
        ' 参数变化，删除旧任务
        LogDebug "LuaTask: 参数变化，删除旧任务 " & existingTaskId
        SetTaskStatus existingTask, CO_TERMINATED
    End If
    ' ========== 创建新任务 ==========
    Dim taskIdStr As String
    taskIdStr = "Task_" & CStr(g_NextTaskId)
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
        .CFS_weight = 1024
        .CFS_vruntime = g_CFS_minVruntime
    End With
    task.ParseResumeSpecs resumeSpec, callerWb, callerWs
    g_Tasks.Add taskIdStr, task
    LuaTask = taskIdStr
    g_NextTaskId = g_NextTaskId + 1
    Exit Function
ErrorHandler:
    Debug.Print "LuaTask Error: " & Err.Number & " - " & Err.Description
    LuaTask = "#ERROR: " & Err.Description
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
    taskstatus = StatusToString(task.taskStatus)

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
        If g_Watches.Exists(callerAddr) Then g_Watches.Remove callerAddr  ' ← 主动清理孤儿
        LuaWatch = "#ERROR: 任务不存在"
        Exit Function
    End If

    ' 计算目标单元格地址
    Dim targetAddr As String
    Dim targetRange As Range

    ' 修复：使用 TypeName 判断是否为 Range，而不是 IsEmpty
    ' IsEmpty 对空单元格的 Range 会返回 True（检查的是 .Value）
    Dim hasTargetCell As Boolean
    hasTargetCell = False

    If Not IsMissing(targetCell) Then
        ' 检查是否为 Range 对象（即使是空单元格也算传入了目标）
        If TypeName(targetCell) = "Range" Then
            hasTargetCell = True
        ElseIf Not IsEmpty(targetCell) Then
            ' 非 Range 且非空（可能是字符串地址）
            hasTargetCell = True
        End If
    End If

    If Not hasTargetCell Then
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
                wi.watchTask.RemoveWatch callerAddr  ' 使用旧的 task 对象
                g_Tasks(taskId).AddWatch callerAddr    ' 使用新的 taskId
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
        g_Tasks(taskId).AddWatch callerAddr
    End If

    ' 返回静态描述文本
    LuaWatch = "监控: " & taskId & "." & field & " -> " & targetAddr

    Exit Function
ErrorHandler:
    LuaWatch = "#ERROR: " & Err.Description
End Function
Private Sub RefreshWatches()
    On Error GoTo ErrorHandler
    ' 快速路径：无脏数据时直接返回
    If Not g_StateDirty Then Exit Sub
    If g_Watches Is Nothing Then Exit Sub
    If g_Watches.Count = 0 Then Exit Sub
    ' 缓存已查找的工作簿
    Dim wbCache As Object
    Set wbCache = CreateObject("Scripting.Dictionary")
    Dim watchCell As Variant
    Dim watchInfo As WatchInfo
    Dim task As TaskUnit
    Dim currentValue As Variant
    For Each watchCell In g_Watches.Keys
        Set watchInfo = g_Watches(watchCell)
        ' 跳过非脏的监控（优化：减少不必要的处理）
        If Not watchInfo.watchDirty Then GoTo NextWatch
        ' 验证任务引用有效性
        If Not g_Tasks.Exists(watchInfo.watchTaskId) Then
            watchInfo.watchDirty = False  ' 清除脏标记，避免重复检查
            GoTo NextWatch
        End If
        ' 验证对象引用一致性
        If watchInfo.watchTask Is Nothing Then
            Set watchInfo.watchTask = g_Tasks(watchInfo.watchTaskId)
        ElseIf Not (watchInfo.watchTask Is g_Tasks(watchInfo.watchTaskId)) Then
            Set watchInfo.watchTask = g_Tasks(watchInfo.watchTaskId)
        End If
        Set task = watchInfo.watchTask
        ' 获取当前字段值
        Select Case watchInfo.watchField
            Case "status"
                currentValue = StatusToString(task.taskStatus)
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
        ' 因为已经通过 watchDirty 标记过滤，这里可以简化
        ' 只有被标记为脏的监控才会到达这里
        Dim needWrite As Boolean
        needWrite = True  ' 脏监控默认需要写入
        ' 可选：额外检查值是否真的变化（进一步优化）
        If Not IsEmpty(watchInfo.watchLastValue) Then
            On Error Resume Next
            If Not IsArray(currentValue) And Not IsArray(watchInfo.watchLastValue) Then
                If Not IsNull(currentValue) And Not IsNull(watchInfo.watchLastValue) Then
                    If Not IsObject(currentValue) And Not IsObject(watchInfo.watchLastValue) Then
                        If currentValue = watchInfo.watchLastValue Then
                            needWrite = False
                        End If
                    End If
                End If
            End If
            On Error GoTo ErrorHandler
        End If
        ' 写入目标单元格
        If needWrite Then
            WriteToTargetCellCached watchInfo.watchTargetCell, currentValue, watchInfo.watchWorkbook, wbCache
            watchInfo.watchLastValue = currentValue
        End If
        ' 清除脏标记
        watchInfo.watchDirty = False
NextWatch:
    Next watchCell
    ' 清理缓存
    Set wbCache = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "RefreshWatches Error: " & Err.Description & " at " & Erl
    Set wbCache = Nothing
End Sub
Private Sub WriteToTargetCellCached(targetAddr As String, value As Variant, wbName As String, wbCache As Object)
    On Error GoTo WriteError

    Dim targetRange As Range
    Dim wb As Workbook
    Dim sheetName As String
    Dim cellAddr As String
    Dim workAddr As String
    
    ' 先从缓存查找工作簿
    If wbCache.Exists(wbName) Then
        Set wb = wbCache(wbName)
    Else
        On Error Resume Next
        Set wb = Application.Workbooks(wbName)
        On Error GoTo WriteError
        
        If wb Is Nothing Then
            Debug.Print "WriteToTargetCellCached: 工作簿不存在 - " & wbName
            Exit Sub
        End If
        wbCache.Add wbName, wb
    End If
    
    workAddr = targetAddr
    sheetName = ""
    cellAddr = ""
    
    ' 处理外层单引号: '[Book.xlsx]Sheet-Name'!$A$1 -> [Book.xlsx]Sheet-Name!$A$1
    ' 查找 '! 模式（单引号后紧跟感叹号）
    Dim quoteExclamPos As Long
    quoteExclamPos = InStr(workAddr, "'!")
    
    If quoteExclamPos > 0 And Left(workAddr, 1) = "'" Then
        ' 格式: '[...]SheetName'!CellAddr
        ' 提取工作表部分（去掉首尾引号）
        Dim sheetPart As String
        sheetPart = Mid(workAddr, 2, quoteExclamPos - 2)  ' 去掉开头的 ' 和结尾的 '
        cellAddr = Mid(workAddr, quoteExclamPos + 2)       ' '! 之后的部分
        
        ' 从 sheetPart 中提取工作表名（可能包含 [Book.xlsx]）
        Dim bracketEnd As Long
        bracketEnd = InStr(sheetPart, "]")
        If bracketEnd > 0 Then
            sheetName = Mid(sheetPart, bracketEnd + 1)
        Else
            sheetName = sheetPart
        End If
    ElseIf Left(workAddr, 1) = "[" Then
        ' 格式: [Book.xlsx]SheetName!CellAddr（无外层引号）
        Dim bracketEnd2 As Long
        bracketEnd2 = InStr(workAddr, "]")
        If bracketEnd2 > 0 Then
            Dim exclamPos As Long
            exclamPos = InStr(bracketEnd2, workAddr, "!")
            If exclamPos > 0 Then
                sheetName = Mid(workAddr, bracketEnd2 + 1, exclamPos - bracketEnd2 - 1)
                cellAddr = Mid(workAddr, exclamPos + 1)
            Else
                cellAddr = Mid(workAddr, bracketEnd2 + 1)
            End If
        End If
    ElseIf InStr(workAddr, "!") > 0 Then
        ' 格式: SheetName!CellAddr 或 'Sheet-Name'!CellAddr
        Dim exclamPos2 As Long
        exclamPos2 = InStr(workAddr, "!")
        sheetName = Left(workAddr, exclamPos2 - 1)
        cellAddr = Mid(workAddr, exclamPos2 + 1)
        ' 去掉工作表名的引号
        If Len(sheetName) >= 2 Then
            If Left(sheetName, 1) = "'" And Right(sheetName, 1) = "'" Then
                sheetName = Mid(sheetName, 2, Len(sheetName) - 2)
            End If
        End If
    Else
        cellAddr = workAddr
    End If
    
    ' 获取目标范围
    On Error Resume Next
    If sheetName <> "" Then
        Set targetRange = wb.Sheets(sheetName).Range(cellAddr)
    Else
        Set targetRange = wb.ActiveSheet.Range(cellAddr)
    End If
    
    If Err.Number <> 0 Or targetRange Is Nothing Then
        Debug.Print "WriteToTargetCellCached: 无法获取范围 - Sheet:[" & sheetName & "] Cell:[" & cellAddr & "] Err:" & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo WriteError
    
    ' 直接写值
    If IsArray(value) Then
        Dim rows As Long, cols As Long
        On Error Resume Next
        rows = UBound(value, 1) - LBound(value, 1) + 1
        cols = UBound(value, 2) - LBound(value, 2) + 1
        If Err.Number = 0 Then
            targetRange.Resize(rows, cols).value = value
        Else
            Err.Clear
            rows = UBound(value) - LBound(value) + 1
            targetRange.Resize(1, rows).value = value
        End If
        On Error GoTo WriteError
    Else
        targetRange.value = value
    End If
    
    Debug.Print "WriteToTargetCellCached: 成功写入 " & targetAddr & " = " & IIf(IsArray(value), "(数组)", CStr(value))
    Exit Sub
    
WriteError:
    Debug.Print "WriteToTargetCellCached Error: " & Err.Description & " - 目标: " & targetAddr
End Sub
' 优化后的 MarkWatchesDirty - O(m) 复杂度
Private Sub MarkWatchesDirty(task As TaskUnit)
    On Error Resume Next

    Dim wc As Variant
    For Each wc In task.taskWatches
        If g_Watches.Exists(CStr(wc)) Then
            g_Watches(CStr(wc)).watchDirty = True
            g_StateDirty = True
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
        MsgBox "错误：任务已启动或已完成，当前状态: " & StatusToString(task.taskStatus), vbExclamation
        Exit Sub
    End If
    ' 保存主状态机栈顶
    Dim mainStackTop As Long
    mainStackTop = lua_gettop(g_LuaState)
    ' 创建协程线程（压入主栈）
    Dim coThread As LongPtr
    coThread = lua_newthread(g_LuaState)
    If coThread = 0 Then
        task.taskError = "无法创建协程线程"
        SetTaskStatus task, CO_ERROR
        lua_settop g_LuaState, mainStackTop
        Exit Sub
    End If
    ' luaL_ref 会弹出栈顶的 thread 并返回引用号
    Dim refResult As Long
    refResult = luaL_ref(g_LuaState, LUA_REGISTRYINDEX)
    ' 检查引用是否成功
    If refResult = LUA_REFNIL Or refResult = LUA_NOREF Then
        task.taskError = "无法创建协程引用 (ref=" & refResult & ")"
        SetTaskStatus task, CO_ERROR
        lua_settop g_LuaState, mainStackTop
        Exit Sub
    End If
    task.taskCoRef = refResult
    task.taskCoThread = coThread
    ' 获取函数并检查
    lua_getglobal g_LuaState, task.taskFunc
    If lua_type(g_LuaState, -1) <> LUA_TFUNCTION Then
        task.taskError = "函数 '" & task.taskFunc & "' 不存在"
        SetTaskStatus task, CO_ERROR
        lua_settop g_LuaState, mainStackTop
        Exit Sub
    End If
    ' 将函数移动到协程栈
    lua_xmove g_LuaState, coThread, 1
    ' 参数计数从 0 开始
    Dim nargs As Long
    nargs = 0
    ' 压入启动参数
    Dim startArgs As Variant
    startArgs = task.taskStartArgs
    If IsArray(startArgs) Then
        Dim i As Long
        Dim lb As Long, ub As Long
        On Error Resume Next
        lb = LBound(startArgs)
        ub = UBound(startArgs)
        If Err.Number = 0 Then
            On Error GoTo ErrorHandler
            For i = lb To ub
                PushValue coThread, startArgs(i)
                nargs = nargs + 1
            Next i
        Else
            Err.Clear
            On Error GoTo ErrorHandler
        End If
    End If
    ' 执行协程
    Dim nres As LongPtr
    Dim result As Long
    result = lua_resume(coThread, g_LuaState, nargs, VarPtr(nres))
    ' 处理结果（这里会设置任务状态和值）
    HandleCoroutineResult task, result, CLng(nres)
    ' 关键：首次启动必须立即刷新，确保第一个返回值被写入
    ' 无论任务状态如何，只要有监控就标记为脏并刷新
    MarkWatchesDirty task
    ' 立即刷新（不等待调度器）
    ' 这确保了：
    ' 1. 第一个 yield 的值能立即显示
    ' 2. 如果任务直接完成(CO_DONE)，结果也能显示
    ' 3. 如果任务出错(CO_ERROR)，错误信息也能显示
    RefreshWatches
    ' 如果是 yield 状态，启动调度器
    If task.taskStatus = CO_YIELD Then
        StartSchedulerIfNeeded
    End If
    Exit Sub
ErrorHandler:
    Dim errMsg As String
    errMsg = "VBA错误: " & Err.Description
    ' 安全地设置错误状态
    If Not task Is Nothing Then
        task.taskError = errMsg
        SetTaskStatus task, CO_ERROR
        MarkWatchesDirty task
        RefreshWatches  ' 确保错误状态也被刷新
    End If
    ' 安全地恢复主栈
    If g_LuaState <> 0 And g_Initialized Then
        On Error Resume Next
        lua_settop g_LuaState, mainStackTop
        On Error GoTo 0
    End If
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

' 手动停止调度器
Public Sub StopScheduler()
    ' 停止调度标志
    If g_SchedulerRunning Then
        g_SchedulerRunning = False
        ' 尝试取消所有 OnTime 调度
        On Error Resume Next
        Application.OnTime g_NextScheduleTime, "SchedulerTick", , False
    End If
End Sub

' 调度器心跳 - 主入口
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

    ' 每个 tick 结束时统一刷新所有脏监控
    ' RefreshWatches 内部会检查每个监控的 watchDirty 标记
    RefreshWatches
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

' CFS 调度核心算法 - 支持两种模式
Private Sub ScheduleByCFS()
    If g_TaskQueue.Count = 0 Or g_ActiveTaskCount = 0 Then Exit Sub

    Dim tickBudget As Double: tickBudget = g_CFS_targetLatency
    Dim idealSlice As Double: idealSlice = IIf(tickBudget / g_ActiveTaskCount < g_CFS_minGranularity, _
                                                g_CFS_minGranularity, tickBudget / g_ActiveTaskCount)
    Dim totalElapsed As Double: totalElapsed = 0

    Do
        Dim selectedTask As TaskUnit: Set selectedTask = CFS_PickNextTask()
        If selectedTask Is Nothing Then Exit Do

        Dim taskStart As Double: taskStart = GetTickCount()
        ResumeCoroutine selectedTask
        Dim taskElapsed As Double: taskElapsed = GetTickCount() - taskStart
        If taskElapsed < g_CFS_minGranularity Then taskElapsed = g_CFS_minGranularity
        totalElapsed = totalElapsed + taskElapsed

        If selectedTask.taskStatus = CO_YIELD Then CFS_UpdateVruntime selectedTask, taskElapsed
        If g_CFS_autoWeight Then CFS_AutoAdjustWeight selectedTask, taskElapsed, idealSlice

        ' 更新工作簿统计
        If g_Workbooks.Exists(selectedTask.taskWorkbook) Then
            With g_Workbooks(selectedTask.taskWorkbook)
                .wbLastTime = taskElapsed
                .wbTotalTime = .wbTotalTime + taskElapsed
                .wbTickCount = .wbTickCount + 1
            End With
        End If
    Loop While (g_ScheduleMode = SCHEDULE_MODE_TIMESLICE) And (totalElapsed < tickBudget) And (g_TaskQueue.Count > 0)
End Sub
' 自动调整任务权重
Private Sub CFS_AutoAdjustWeight(task As TaskUnit, actualTime As Double, idealSlice As Double)
    ' 基于实际执行时间与理想时间片的比较调整权重
    ' 如果任务执行时间远超理想时间片，降低权重（让其他任务有机会）
    ' 如果任务执行时间远低于理想时间片，提高权重（它可以跑更多）

    Dim ratio As Double
    Dim adjustment As Double
    Dim newWeight As Double

    If idealSlice <= 0 Then Exit Sub

    ratio = actualTime / idealSlice

    ' 根据比例计算调整系数
    If ratio > 2 Then
        ' 执行时间过长，降低权重
        adjustment = 0.95
    ElseIf ratio > 1.5 Then
        adjustment = 0.98
    ElseIf ratio < 0.5 Then
        ' 执行时间很短，提高权重
        adjustment = 1.05
    ElseIf ratio < 0.75 Then
        adjustment = 1.02
    Else
        ' 在合理范围内，不调整
        Exit Sub
    End If

    newWeight = task.CFS_weight * adjustment

    ' 限制权重范围（使用 nice 值映射表的范围）
    If newWeight < g_CFS_niceToWeight(39) Then
        newWeight = g_CFS_niceToWeight(39)  ' 最低权重 ~12
    ElseIf newWeight > g_CFS_niceToWeight(0) Then
        newWeight = g_CFS_niceToWeight(0)   ' 最高权重 ~90000
    End If

    task.CFS_weight = newWeight
End Sub
' 选择 vruntime 最小的任务 - O(1)，有序队列保证队首就是最小的
Private Function CFS_PickNextTask() As TaskUnit
    On Error Resume Next
    Set CFS_PickNextTask = Nothing

    If g_TaskQueue.Count = 0 Then Exit Function

    ' 收集需要移除的无效任务
    Dim toRemove As Collection
    Set toRemove = New Collection

    Dim i As Long
    Dim task As TaskUnit
    Dim foundTask As TaskUnit
    Set foundTask = Nothing

    For i = 1 To g_TaskQueue.Count
        Set task = g_TaskQueue(i)

        ' 检查任务状态
        If task.taskStatus = CO_YIELD Then
            ' 找到第一个有效任务
            If foundTask Is Nothing Then
                Set foundTask = task
            End If
        Else
            ' 非 YIELD 状态的任务应该被移除
            toRemove.Add i
        End If
    Next

    ' 从后向前移除无效任务（避免索引错乱）
    Dim removeIdx As Variant
    For i = toRemove.Count To 1 Step -1
        g_TaskQueue.Remove CLng(toRemove(i))
    Next

    ' 如果移除了任务，更新活跃计数
    If toRemove.Count > 0 Then
        LogDebug "CFS_PickNextTask: 清理了 " & toRemove.Count & " 个无效任务"
    End If

    Set CFS_PickNextTask = foundTask
End Function
' 更新任务的 vruntime 并重新排序
Private Sub CFS_UpdateVruntime(task As TaskUnit, actualRuntime As Double)
    ' 确保最小执行粒度
    If actualRuntime < g_CFS_minGranularity Then actualRuntime = g_CFS_minGranularity

    ' 计算加权虚拟运行时间
    Dim vruntimeDelta As Double
    vruntimeDelta = actualRuntime * (1024 / task.CFS_weight)

    task.CFS_vruntime = task.CFS_vruntime + vruntimeDelta
    task.CFS_lastScheduled = GetTickCount()

    ' 重新调整队列位置（维护有序性）
    TaskQueueReposition task

    ' 更新全局最小 vruntime（队首就是最小的）
    If g_TaskQueue.Count > 0 Then
        Dim firstTask As TaskUnit
        Set firstTask = g_TaskQueue(1)
        g_CFS_minVruntime = firstTask.CFS_vruntime
    End If
End Sub

' Resume 协程
Private Sub ResumeCoroutine(task As TaskUnit)
    On Error GoTo ErrorHandler
    Dim taskStart As Long
    taskStart = GetTickCount()
    Dim coThread As LongPtr
    coThread = task.taskCoThread
    ' 检查工作簿是否仍然打开
    If Len(task.taskWorkbook) > 0 Then
        Dim wb As Workbook
        On Error Resume Next
        Set wb = Application.Workbooks(task.taskWorkbook)
        On Error GoTo ErrorHandler
        If wb Is Nothing Then
            task.taskError = "工作簿已关闭: " & task.taskWorkbook
            SetTaskStatus task, CO_ERROR
            MarkWatchesDirty task  ' 确保错误状态被标记
            Exit Sub
        End If
    End If
    ' 清空协程栈
    lua_settop coThread, 0
    ' 使用新的参数解析机制获取动态参数
    Dim resumeArgs As Variant
    resumeArgs = task.GetResumeArgs()
    Dim nargs As Long
    nargs = 0
    ' 压入参数
    If IsArray(resumeArgs) Then
        Dim lb As Long, ub As Long
        On Error Resume Next
        lb = LBound(resumeArgs)
        ub = UBound(resumeArgs)
        If Err.Number = 0 Then
            On Error GoTo ErrorHandler
            Dim i As Long
            For i = lb To ub
                PushValue coThread, resumeArgs(i)
                nargs = nargs + 1
            Next i
        End If
        On Error GoTo ErrorHandler
    End If
    ' 执行 resume
    Dim nres As LongPtr
    Dim result As Long
    result = lua_resume(coThread, g_LuaState, nargs, VarPtr(nres))
    
    ' 处理结果
    HandleCoroutineResult task, result, CLng(nres)
    ' 在调度器循环中，每个任务执行后只标记脏
    ' 统一在 SchedulerTick 结束时刷新
    MarkWatchesDirty task
    ' 性能统计
    Dim taskElapsed As Double
    taskElapsed = GetTickCount() - taskStart
    With task
        .taskLastTime = taskElapsed
        .taskTotalTime = .taskTotalTime + taskElapsed
        .taskTickCount = .taskTickCount + 1
    End With
    Exit Sub
ErrorHandler:
    Dim errorDetails As String
    errorDetails = "Resume错误: " & Err.Description
    SetTaskStatus task, CO_ERROR
    task.taskError = errorDetails
    MarkWatchesDirty task  ' 确保错误状态被标记
    Debug.Print "ResumeCoroutine Error: " & errorDetails
End Sub
' ============================================
' 第六部分：推栈压栈函数
' ============================================
' 处理协程返回结果
Private Sub HandleCoroutineResult(task As TaskUnit, result As Long, nres As Long)
    On Error GoTo ErrorHandler
    Dim coThread As LongPtr: coThread = task.taskCoThread
    Dim hasResult As Boolean: hasResult = (nres > 0 And lua_gettop(coThread) > 0)

    If result = LUA_OK Or result = LUA_YIELD Then
        If hasResult Then
            Dim statusSet As Boolean
            ' 先检查是否为控制表，再决定如何获取值
            Dim rawValue As Variant
            rawValue = GetValueForYieldReturn(coThread, -1)
            statusSet = ParseYieldReturn(task, rawValue, (result = LUA_OK))
        End If
        If Not statusSet Then
            SetTaskStatus task, IIf(result = LUA_OK, CO_DONE, CO_YIELD)
        End If
    Else
        task.taskError = IIf(hasResult, GetStringFromState(coThread, -1), "协程错误: 代码 " & result)
        SetTaskStatus task, CO_ERROR
    End If

    lua_settop coThread, 0
    Exit Sub
ErrorHandler:
    task.taskError = "处理结果错误: " & Err.Description
    SetTaskStatus task, CO_ERROR
    If coThread <> 0 Then lua_settop coThread, 0
End Sub

' 解析 yield/return 字典
Private Function ParseYieldReturn(task As TaskUnit, data As Variant, isFinal As Boolean) As Boolean
    On Error Resume Next
    ParseYieldReturn = False

    ' 非数组直接作为 value
    If Not IsArray(data) Then
        task.taskValue = data
        Exit Function
    End If

    ' 检查是否为字典格式（二维数组，第二维为2列）
    Dim cols As Long
    cols = UBound(data, 2) - LBound(data, 2) + 1
    If Err.Number <> 0 Or cols <> 2 Then
        Err.Clear
        task.taskValue = data
        Exit Function
    End If

    ' 已知的字典键名（小写）
    Const KNOWN_KEYS As String = "|progress|message|value|error|status|"

    Dim hasKnownKey As Boolean
    hasKnownKey = False

    Dim i As Long
    Dim testKey As String

    ' 遍历所有行，检查是否至少有一个已知键
    For i = LBound(data, 1) To UBound(data, 1)
        testKey = LCase(Trim(CStr(data(i, 1))))
        If InStr(1, KNOWN_KEYS, "|" & testKey & "|", vbTextCompare) > 0 Then
            hasKnownKey = True
            Exit For
        End If
    Next

    ' 如果没有任何已知键，作为普通数据处理
    If Not hasKnownKey Then
        task.taskValue = data
        Exit Function
    End If

    ' 解析字典键值对
    Dim key As String, value As Variant
    For i = LBound(data, 1) To UBound(data, 1)
        key = LCase(Trim(CStr(data(i, 1))))
        value = data(i, 2)

        Select Case key
            Case "progress"
                On Error Resume Next
                task.taskProgress = CDbl(value)
                If Err.Number <> 0 Then Err.Clear
                On Error Resume Next
            Case "message"
                task.taskMessage = CStr(value)
            Case "value"
                ' value 字段的值直接赋给 taskValue
                ' 注意：此时 value 已经通过 GetValue 处理过了
                ' 如果 value 是带 __size 的表，已经被正确转换
                task.taskValue = value
            Case "error"
                task.taskError = CStr(value)
            Case "status"
                ParseYieldReturn = True  ' 标记 Lua 显式设置了状态
                Select Case LCase(Trim(CStr(value)))
                    Case "yield"
                        If Not isFinal Then SetTaskStatus task, CO_YIELD
                    Case "paused"
                        SetTaskStatus task, CO_PAUSED
                    Case "defined"
                        SetTaskStatus task, CO_DEFINED
                    Case "done"
                        SetTaskStatus task, CO_DONE
                    Case "error"
                        SetTaskStatus task, CO_ERROR
                    Case "terminated"
                        SetTaskStatus task, CO_TERMINATED
                    Case Else
                        ' 未知状态值，yield 时默认继续
                        If Not isFinal Then SetTaskStatus task, CO_YIELD
                        ParseYieldReturn = False
                End Select
            ' Case Else: 忽略未知键（允许用户添加自定义键）
        End Select
    Next i
End Function

' 专门用于 yield/return 的值获取（智能判断是否为控制表）
Private Function GetValueForYieldReturn(ByVal L As LongPtr, ByVal idx As Long) As Variant
    ' 如果不是表，直接返回
    If lua_type(L, idx) <> LUA_TTABLE Then
        GetValueForYieldReturn = GetValue(L, idx, False)
        Exit Function
    End If
    ' 标准化索引
    Dim absIdx As Long
    If idx < 0 Then
        absIdx = lua_gettop(L) + idx + 1
    Else
        absIdx = idx
    End If
    ' 检查是否为控制表（包含已知控制键）
    Dim isControlTable As Boolean
    isControlTable = IsLuaControlTable(L, absIdx)
    ' 根据判断结果获取值
    GetValueForYieldReturn = GetValue(L, idx, isControlTable)
End Function

' 统一压栈函数 - 简化版
Private Sub PushValue(ByVal L As LongPtr, ByVal value As Variant)
    If IsMissing(value) Then
        lua_pushnil L
        Exit Sub
    End If

    ' Range 对象直接取值（单次）
    If TypeName(value) = "Range" Then value = value.Value

    Select Case VarType(value)
        Case vbEmpty, vbNull
            lua_pushnil L
        Case vbBoolean
            lua_pushboolean L, IIf(value, 1, 0)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbLongLong
            lua_pushnumber L, CDbl(value)
        Case vbString
            PushUTF8ToState L, CStr(value)
        Case vbError
            lua_pushnil L
        Case Else
            If IsArray(value) Then
                PushArray L, value
            ElseIf IsNumeric(value) Then
                lua_pushnumber L, CDbl(value)
            Else
                PushUTF8ToState L, CStr(value)
            End If
    End Select
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

' 从 Lua 栈获取值
Private Function GetValue(ByVal L As LongPtr, ByVal idx As Long, Optional ByVal isControlTable As Boolean = False) As Variant
    Select Case lua_type(L, idx)
        Case LUA_TNIL
            GetValue = Empty
        Case LUA_TBOOLEAN
            GetValue = (lua_toboolean(L, idx) <> 0)
        Case LUA_TNUMBER
            GetValue = lua_tonumberx(L, idx, 0)
        Case LUA_TSTRING
            ' 使用 UTF-8 安全的字符串获取
            GetValue = GetStringFromState(L, idx)
        Case LUA_TTABLE
            GetValue = LuaTableToVariant(L, idx, isControlTable)
        Case Else
            GetValue = "#LUA_TYPE_" & lua_type(L, idx)
    End Select
End Function

' VBA 字符串 (Unicode) -> UTF-8 字节数组
Private Function StringToUTF8(ByVal str As String) As Byte()
    Dim utf8Bytes() As Byte
    Dim utf8Len As Long
    ' 空字符串返回空数组（长度为0）
    If Len(str) = 0 Then
        ReDim utf8Bytes(0 To -1)  ' 空数组
        StringToUTF8 = utf8Bytes
        Exit Function
    End If
    ' 获取所需缓冲区大小
    utf8Len = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), Len(str), 0, 0, 0, 0)
    If utf8Len = 0 Then
        ReDim utf8Bytes(0 To -1)  ' 转换失败返回空数组
        StringToUTF8 = utf8Bytes
        Exit Function
    End If

    ' 分配缓冲区
    ReDim utf8Bytes(0 To utf8Len - 1)

    ' 执行转换
    WideCharToMultiByte CP_UTF8, 0, StrPtr(str), Len(str), VarPtr(utf8Bytes(0)), utf8Len, 0, 0
    StringToUTF8 = utf8Bytes
End Function
' UTF-8 字节指针 -> VBA 字符串 (已有，优化版本)
Private Function UTF8ToString(ByVal ptr As LongPtr, ByVal byteLen As Long) As String
    If ptr = 0 Or byteLen <= 0 Then
        UTF8ToString = vbNullString
        Exit Function
    End If

    ' 复制 UTF-8 字节到 VBA 数组
    Dim utf8Bytes() As Byte
    ReDim utf8Bytes(0 To byteLen - 1)
    CopyMemory utf8Bytes(0), ByVal ptr, byteLen

    ' 计算 Unicode 字符数
    Dim nChars As Long
    nChars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8Bytes(0)), byteLen, 0, 0)

    If nChars = 0 Then
        ' 转换失败，尝试 ASCII 回退
        UTF8ToString = StrConv(utf8Bytes, vbUnicode)
        Exit Function
    End If

    ' 执行转换
    UTF8ToString = String$(nChars, vbNullChar)
    MultiByteToWideChar CP_UTF8, 0, VarPtr(utf8Bytes(0)), byteLen, StrPtr(UTF8ToString), nChars
End Function
' 安全的字符串压栈（自动处理 UTF-8 转换）
Private Sub PushUTF8ToState(ByVal L As LongPtr, ByVal str As String)
    ' 空字符串统一处理
    If Len(str) = 0 Then
        lua_pushlstring L, 0, 0  ' 压入空的 Lua 字符串
        Exit Sub
    End If
    Dim utf8Bytes() As Byte
    utf8Bytes = StringToUTF8(str)
    ' 检查转换结果
    Dim actualLen As LongPtr
    On Error Resume Next
    actualLen = UBound(utf8Bytes) - LBound(utf8Bytes) + 1
    If Err.Number <> 0 Then
        ' 空数组情况
        Err.Clear
        lua_pushlstring L, 0, 0
        Exit Sub
    End If
    On Error GoTo 0

    If actualLen <= 0 Then
        lua_pushlstring L, 0, 0
        Exit Sub
    End If
    lua_pushlstring L, VarPtr(utf8Bytes(0)), actualLen
End Sub
' 从 Lua 栈获取字符串
Private Function GetStringFromState(ByVal L As LongPtr, ByVal idx As Long) As String
    Dim ptr As LongPtr
    Dim strLen As LongPtr  ' 使用 LongPtr 接收长度

    strLen = 0
    ptr = lua_tolstring(L, idx, VarPtr(strLen))

    If ptr = 0 Or strLen = 0 Then
        GetStringFromState = vbNullString
        Exit Function
    End If

    ' 直接调用统一转换函数
    GetStringFromState = UTF8ToString(ptr, CLng(strLen))
End Function

' 将 Lua table 转换为 VBA Variant (字典或数组)
' 参数 isControlTable: 标记是否为协程控制表的顶层（控制表不处理 __size）
Private Function LuaTableToVariant(ByVal L As LongPtr, ByVal idx As Long, Optional ByVal isControlTable As Boolean = False) As Variant
    On Error GoTo ErrorHandler
    ' 标准化索引为正数
    If idx < 0 Then idx = lua_gettop(L) + idx + 1
    Dim sizeRows As Long, sizeCols As Long

    ' 只有非控制表才检查 __size
    If Not isControlTable Then
        If GetTableSizeField(L, idx, sizeRows, sizeCols) Then
            LuaTableToVariant = BuildSparseArray(L, idx, sizeRows, sizeCols)
            Exit Function
        End If
    End If

    ' 无 __size 时自动检测
    Dim length As LongPtr
    length = lua_rawlen(L, idx)

    ' 如果长度为0，尝试判断是否为字典
    If length = 0 Then
        Dim topBefore As Long
        topBefore = lua_gettop(L)

        lua_pushnil L
        If lua_next(L, idx) <> 0 Then
            ' 有内容，是字典
            lua_settop L, topBefore
            LuaTableToVariant = LuaTableToDictArray(L, idx)
        Else
            ' 空表，返回空字符串，而不是Empty(会显示0，而不是空白)
            LuaTableToVariant = vbNullString
        End If
        Exit Function
    End If

    ' 检查是否为纯数组（只有连续数字索引 1 到 length，没有其他键）
    Dim isPureArray As Boolean, testTop As Long
    Dim totalKeyCount As Long
    isPureArray = True
    totalKeyCount = 0
    testTop = lua_gettop(L)

    lua_pushnil L
    Do While lua_next(L, idx) <> 0
        totalKeyCount = totalKeyCount + 1

        Dim keyType As Long
        keyType = lua_type(L, -2)

        If keyType = LUA_TNUMBER Then
            Dim keyNum As Double
            keyNum = lua_tonumberx(L, -2, 0)
            ' 检查是否为有效的数组索引
            If keyNum <> CLng(keyNum) Or keyNum < 1 Or keyNum > length Then
                isPureArray = False
                lua_settop L, testTop
                Exit Do
            End If
        Else
            ' 有非数字键（字符串键等），不是纯数组
            isPureArray = False
            lua_settop L, testTop
            Exit Do
        End If
        lua_settop L, -2  ' 弹出 value，保留 key
    Loop

    ' 额外检查：键的数量必须等于 length（确保是连续的 1 到 length）
    If isPureArray And totalKeyCount <> CLng(length) Then
        isPureArray = False
    End If

    lua_settop L, testTop

    ' 如果不是纯数组，按字典处理
    If Not isPureArray Then
        LuaTableToVariant = LuaTableToDictArray(L, idx)
        Exit Function
    End If

    ' 纯数组处理
    lua_rawgeti L, idx, 1
    Dim firstIsTable As Boolean
    firstIsTable = (lua_type(L, -1) = LUA_TTABLE)
    lua_settop L, -2

    Dim i As Long, j As Long
    If firstIsTable Then
        ' 二维数组
        lua_rawgeti L, idx, 1
        Dim cols As LongPtr
        cols = lua_rawlen(L, -1)
        lua_settop L, -2
        If cols = 0 Then cols = 1

        Dim arr2D() As Variant
        ReDim arr2D(1 To CLng(length), 1 To CLng(cols))

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
                    End If
                Next j
            End If
            lua_settop L, -2
        Next i
        LuaTableToVariant = arr2D
    Else
        ' 一维数组
        Dim arr1D() As Variant
        ReDim arr1D(1 To 1, 1 To CLng(length))
        For i = 1 To CLng(length)
            lua_rawgeti L, idx, CLng(i)
            arr1D(1, i) = GetValue(L, -1)
            lua_settop L, -2
        Next i
        LuaTableToVariant = arr1D
    End If
    Exit Function

ErrorHandler:
    LuaTableToVariant = "#TABLE_ERROR: " & Err.Description
End Function

' 检查 Lua 表是否为协程控制表
Private Function IsLuaControlTable(ByVal L As LongPtr, ByVal idx As Long) As Boolean
    On Error Resume Next
    IsLuaControlTable = False

    ' 已知的控制表键名
    Const CONTROL_KEYS As String = "|progress|message|value|error|status|"

    ' 检查表中是否存在任何控制键
    Dim testTop As Long
    testTop = lua_gettop(L)

    ' 检查 progress
    lua_getfield L, idx, "progress"
    If lua_type(L, -1) <> LUA_TNIL Then
        lua_settop L, testTop
        IsLuaControlTable = True
        Exit Function
    End If
    lua_settop L, testTop

    ' 检查 message
    lua_getfield L, idx, "message"
    If lua_type(L, -1) <> LUA_TNIL Then
        lua_settop L, testTop
        IsLuaControlTable = True
        Exit Function
    End If
    lua_settop L, testTop

    ' 检查 value
    lua_getfield L, idx, "value"
    If lua_type(L, -1) <> LUA_TNIL Then
        lua_settop L, testTop
        IsLuaControlTable = True
        Exit Function
    End If
    lua_settop L, testTop

    ' 检查 error
    lua_getfield L, idx, "error"
    If lua_type(L, -1) <> LUA_TNIL Then
        lua_settop L, testTop
        IsLuaControlTable = True
        Exit Function
    End If
    lua_settop L, testTop

    ' 检查 status
    lua_getfield L, idx, "status"
    If lua_type(L, -1) <> LUA_TNIL Then
        lua_settop L, testTop
        IsLuaControlTable = True
        Exit Function
    End If
    lua_settop L, testTop

    ' 没有找到任何控制键，不是控制表
    IsLuaControlTable = False
End Function

' 获取 __size 字段并按尺寸构建数组
Private Function GetTableSizeField(ByVal L As LongPtr, ByVal idx As Long, ByRef rows As Long, ByRef cols As Long) As Boolean
    GetTableSizeField = False

    lua_getfield L, idx, "__size"
    If lua_type(L, -1) = LUA_TTABLE Then
        lua_rawgeti L, -1, 1
        lua_rawgeti L, -2, 2
        If lua_type(L, -2) = LUA_TNUMBER And lua_type(L, -1) = LUA_TNUMBER Then
            rows = CLng(lua_tonumberx(L, -2, 0))
            cols = CLng(lua_tonumberx(L, -1, 0))
            GetTableSizeField = (rows > 0 And cols > 0)
        End If
        lua_settop L, -3
    End If
    lua_settop L, -2
End Function

' 按指定尺寸构建数组
Private Function BuildSparseArray(ByVal L As LongPtr, ByVal idx As Long, ByVal rows As Long, ByVal cols As Long) As Variant
    On Error GoTo ErrorHandler

    Dim arr() As Variant
    ReDim arr(1 To rows, 1 To cols)

    ' 初始化为 Empty（Excel 显示为空白）
    Dim r As Long, c As Long
    For r = 1 To rows
        For c = 1 To cols
            arr(r, c) = Empty
        Next c
    Next r

    ' 检查表的结构类型
    Dim length As LongPtr
    length = lua_rawlen(L, idx)

    ' 判断是否为纯二维数组结构（第一个元素是表）
    Dim isNestedArray As Boolean
    isNestedArray = False

    If length > 0 Then
        lua_rawgeti L, idx, 1
        isNestedArray = (lua_type(L, -1) = LUA_TTABLE)
        lua_settop L, -2
    End If

    If isNestedArray Then
        ' 二维数组结构：{{1,2,3}, {4,5,6}, ...}
        For r = 1 To rows
            If r > length Then Exit For
            lua_rawgeti L, idx, CLng(r)
            If lua_type(L, -1) = LUA_TTABLE Then
                Dim subLen As LongPtr
                subLen = lua_rawlen(L, -1)
                For c = 1 To cols
                    If c > subLen Then Exit For
                    lua_rawgeti L, -1, CLng(c)
                    arr(r, c) = GetValue(L, -1)
                    lua_settop L, -2
                Next c
            End If
            lua_settop L, -2
        Next r
    Else
        ' 混合表/字典结构：按键值对填充为两列
        Dim topBefore As Long
        topBefore = lua_gettop(L)

        Dim rowIdx As Long
        rowIdx = 1

        Dim skipThisKey As Boolean
        Dim keyStr As String

        lua_pushnil L
        Do While lua_next(L, idx) <> 0
            skipThisKey = False

            ' 检查是否为 __size 键（需要跳过）
            If lua_type(L, -2) = LUA_TSTRING Then
                keyStr = GetStringFromState(L, -2)
                If keyStr = "__size" Then
                    skipThisKey = True
                End If
            End If

            If Not skipThisKey Then
                ' 检查是否超出行数限制
                If rowIdx > rows Then
                    ' 超出限制，清理栈并退出循环
                    lua_settop L, topBefore
                    Exit Do
                End If

                ' 填充键（第一列）
                If cols >= 1 Then
                    Select Case lua_type(L, -2)
                        Case LUA_TSTRING
                            arr(rowIdx, 1) = GetStringFromState(L, -2)
                        Case LUA_TNUMBER
                            arr(rowIdx, 1) = lua_tonumberx(L, -2, 0)
                        Case LUA_TBOOLEAN
                            arr(rowIdx, 1) = (lua_toboolean(L, -2) <> 0)
                        Case Else
                            arr(rowIdx, 1) = "#KEY"
                    End Select
                End If

                ' 填充值（第二列）
                If cols >= 2 Then
                    arr(rowIdx, 2) = GetValue(L, -1)
                End If

                rowIdx = rowIdx + 1
            End If

            ' 弹出 value，保留 key 用于下次迭代
            lua_settop L, -2
        Loop

        lua_settop L, topBefore
    End If

    BuildSparseArray = arr
    Exit Function

ErrorHandler:
    ReDim arr(1 To rows, 1 To cols)
    BuildSparseArray = arr
End Function

' 将 Lua table 转换为字典数组
Private Function LuaTableToDictArray(ByVal L As LongPtr, ByVal idx As Long) As Variant
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
        LuaTableToDictArray = Empty
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

    LuaTableToDictArray = result
    Exit Function

ErrorHandler:
    LuaTableToDictArray = "#DICT_ERROR: " & Err.Description
End Function
' ============================================
' 第七部分：辅助函数
' ============================================
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

' 检查 g_TaskQueue 中是否存在指定任务 - O(n)
' 注：有序队列仍需遍历检查存在性
Private Function TaskQueueExists(task As TaskUnit) As Boolean
    On Error Resume Next
    Dim i As Long
    For i = 1 To g_TaskQueue.Count
        If g_TaskQueue(i) Is task Then
            TaskQueueExists = True
            Exit Function
        End If
    Next
    TaskQueueExists = False
End Function

' 从 g_TaskQueue 中移除元素 - O(n)
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

' 有序插入任务到队列（按 vruntime 升序）- O(n)
' 队列头部是 vruntime 最小的任务
Private Sub TaskQueueAddSorted(task As TaskUnit)
    On Error Resume Next

    ' 先检查是否已存在
    If TaskQueueExists(task) Then Exit Sub

    ' 空队列直接添加
    If g_TaskQueue.Count = 0 Then
        g_TaskQueue.Add task
        Exit Sub
    End If

    ' 找到插入位置（按 vruntime 升序）
    Dim i As Long
    Dim currentTask As TaskUnit

    For i = 1 To g_TaskQueue.Count
        Set currentTask = g_TaskQueue(i)
        If task.CFS_vruntime < currentTask.CFS_vruntime Then
            ' 插入到位置 i 之前
            If i = 1 Then
                g_TaskQueue.Add task, Before:=1
            Else
                g_TaskQueue.Add task, Before:=i
            End If
            Exit Sub
        End If
    Next

    ' 如果 vruntime 最大，添加到末尾
    g_TaskQueue.Add task
End Sub

' 重新调整任务在队列中的位置（vruntime 更新后调用）- O(n)
Private Sub TaskQueueReposition(task As TaskUnit)
    ' 先移除
    TaskQueueRemove task
    ' 重新有序插入
    TaskQueueAddSorted task
End Sub

' 统一设置任务状态并处理所有副作用
' 这是任务状态管理的唯一入口点
' newStatus: 目标状态
' options: 可选参数，用于控制特殊行为
'   - "DELETE": 状态设置后从 g_Tasks 中删除任务
'   - "COROUTINE": 不释放协程（用于暂停等场景）
Private Sub SetTaskStatus(task As TaskUnit, newStatus As CoStatus, Optional options As String = vbNullString)
    If task Is Nothing Then Exit Sub

    Dim oldStatus As CoStatus
    oldStatus = task.taskStatus

    ' 相同状态不处理（除非有特殊选项）
    If oldStatus = newStatus And options = vbNullString Then Exit Sub

    Dim taskIdStr As String
    taskIdStr = "Task_" & task.taskId

    Dim shouldDelete As Boolean
    Dim keepCoroutine As Boolean
    shouldDelete = InStr(1, options, "DELETE", vbTextCompare) > 0
    keepCoroutine = InStr(1, options, "COROUTINE", vbTextCompare) > 0

    ' 更新活跃任务计数器（状态变更前）
    g_ActiveTaskCount = g_ActiveTaskCount - IIf(oldStatus = CO_YIELD, 1, 0)
    ' 更新状态
    task.taskStatus = newStatus
    ' 更新活跃任务计数器（状态变更后）
    g_ActiveTaskCount = g_ActiveTaskCount + IIf(newStatus = CO_YIELD, 1, 0)

    ' 根据新状态处理副作用
    Select Case newStatus
        Case CO_DONE, CO_ERROR
            ' 完成或错误：释放协程，从队列移除，但保留任务数据
            If Not keepCoroutine Then ReleaseTaskCoroutine task
            TaskQueueRemove task
            ' 如果指定删除，则完全清理
            If shouldDelete Then
                CleanupTaskWatches task
                If g_Tasks.Exists(taskIdStr) Then
                    Set g_Tasks(taskIdStr) = Nothing
                    g_Tasks.Remove taskIdStr
                End If
            End If

        Case CO_TERMINATED
            ' 终止：释放协程，从队列移除，清理 Watch
            If Not keepCoroutine Then ReleaseTaskCoroutine task
            TaskQueueRemove task
            CleanupTaskWatches task
            ' 终止状态默认从任务列表删除
            If g_Tasks.Exists(taskIdStr) Then
                Set g_Tasks(taskIdStr) = Nothing
                g_Tasks.Remove taskIdStr
            End If

        Case CO_PAUSED
            ' 暂停：从队列移除但保留协程
            TaskQueueRemove task

        Case CO_YIELD
            ' 从 PAUSED 恢复时，重置 CFS 参数并加入队列
            If oldStatus = CO_PAUSED Then
                task.CFS_vruntime = g_CFS_minVruntime
                task.CFS_lastScheduled = GetTickCount()
                TaskQueueAddSorted task  ' 使用有序插入
            End If
            ' 新任务首次进入 YIELD 状态，使用 nice 值初始化权重
            If oldStatus = CO_DEFINED Then
                ' 默认 nice=0，使用索引 20
                task.CFS_weight = g_CFS_niceToWeight(20)
                task.CFS_vruntime = g_CFS_minVruntime
                task.CFS_lastScheduled = GetTickCount()
                TaskQueueAddSorted task  ' 使用有序插入
            End If

        Case CO_DEFINED
            ' 重置任务到初始状态
            If Not keepCoroutine Then ReleaseTaskCoroutine task
            TaskQueueRemove task
            ' 重置所有运行时属性
            With task
                .taskProgress = 0
                .taskMessage = vbNullString
                .taskValue = Empty
                .taskError = vbNullString
                .taskLastTime = 0
                .taskTotalTime = 0
                .taskTickCount = 0
                .CFS_vruntime = 0
                .CFS_weight = 1024
                .CFS_lastScheduled = 0
                .taskCoRef = 0
                .taskCoThread = 0
            End With
    End Select

    ' 标记监控为脏（除非任务已被删除）
    If g_Tasks.Exists(taskIdStr) Then
        g_StateDirty = True
        MarkWatchesDirty task
    End If

    LogDebug "状态转换: " & taskIdStr & " " & StatusToString(oldStatus) & " -> " & StatusToString(newStatus) & IIf(options <> "", " [" & options & "]", "")
End Sub

' 统一释放任务的协程资源
Public Sub ReleaseTaskCoroutine(task As TaskUnit)
    ' 修复3：添加完整的安全检查
    If task Is Nothing Then Exit Sub
    If task.taskCoRef = 0 Then Exit Sub

    ' 检查 Lua 状态机是否有效
    If g_LuaState = 0 Then
        ' Lua 已关闭，只清除引用不调用 API
        task.ClearCoroutineRef
        Exit Sub
    End If

    ' 检查是否已初始化
    If Not g_Initialized Then
        task.ClearCoroutineRef
        Exit Sub
    End If

    ' 执行释放
    On Error Resume Next
    luaL_unref g_LuaState, LUA_REGISTRYINDEX, task.taskCoRef
    On Error GoTo 0

    ' 清除任务中的引用
    task.ClearCoroutineRef
End Sub
' 辅助函数：状态转字符串
Private Function StatusToString(status As CoStatus) As String
    Select Case status
        Case CO_DEFINED: StatusToString = "defined"
        Case CO_YIELD: StatusToString = "yield"
        Case CO_PAUSED: StatusToString = "paused"
        Case CO_DONE: StatusToString = "done"
        Case CO_ERROR: StatusToString = "error"
        Case CO_TERMINATED: StatusToString = "terminated"
        Case Else: StatusToString = "UNKNOWN(" & status & ")"
    End Select
End Function

' 辅助函数：比较两个 Variant 数组是否相等
Private Function CompareVariantArrays(arr1 As Variant, arr2 As Variant) As Boolean
    On Error GoTo NotEqual
    ' 处理非数组情况
    If Not IsArray(arr1) And Not IsArray(arr2) Then
        ' 两者都不是数组，直接比较
        If IsObject(arr1) Or IsObject(arr2) Then
            CompareVariantArrays = False
        Else
            CompareVariantArrays = (arr1 = arr2)
        End If
        Exit Function
    End If
    ' 一个是数组一个不是
    If IsArray(arr1) Xor IsArray(arr2) Then
        CompareVariantArrays = False
        Exit Function
    End If
    ' 两者都是数组
    Dim lb1 As Long, ub1 As Long
    Dim lb2 As Long, ub2 As Long
    On Error Resume Next
    lb1 = LBound(arr1)
    ub1 = UBound(arr1)
    lb2 = LBound(arr2)
    ub2 = UBound(arr2)
    If Err.Number <> 0 Then
        Err.Clear
        ' 空数组处理
        On Error GoTo NotEqual
        Dim test1 As Long, test2 As Long
        On Error Resume Next
        test1 = UBound(arr1)
        Dim isEmpty1 As Boolean: isEmpty1 = (Err.Number <> 0)
        Err.Clear
        test2 = UBound(arr2)
        Dim isEmpty2 As Boolean: isEmpty2 = (Err.Number <> 0)
        Err.Clear
        On Error GoTo NotEqual
        ' 两个都是空数组则相等
        CompareVariantArrays = (isEmpty1 And isEmpty2)
        Exit Function
    End If
    On Error GoTo NotEqual
    ' 长度不同
    If (ub1 - lb1) <> (ub2 - lb2) Then
        CompareVariantArrays = False
        Exit Function
    End If
    ' 逐元素比较
    Dim i As Long
    For i = 0 To (ub1 - lb1)
        Dim v1 As Variant, v2 As Variant
        v1 = arr1(lb1 + i)
        v2 = arr2(lb2 + i)
        ' 递归比较（处理嵌套数组）
        If IsArray(v1) Or IsArray(v2) Then
            If Not CompareVariantArrays(v1, v2) Then
                CompareVariantArrays = False
                Exit Function
            End If
        ElseIf IsObject(v1) Or IsObject(v2) Then
            ' 对象比较：简单判断是否同一引用
            If Not (v1 Is v2) Then
                CompareVariantArrays = False
                Exit Function
            End If
        Else
            ' 值比较
            If VarType(v1) <> VarType(v2) Then
                CompareVariantArrays = False
                Exit Function
            End If
            If v1 <> v2 Then
                CompareVariantArrays = False
                Exit Function
            End If
        End If
    Next i

    CompareVariantArrays = True
    Exit Function
NotEqual:
    CompareVariantArrays = False
End Function

' 辅助函数：比较两个 Resume 参数规格是否相等
Private Function CompareResumeSpecs(spec1 As Variant, spec2 As Variant) As Boolean
    On Error GoTo NotEqual
    ' 处理非数组情况
    If Not IsArray(spec1) And Not IsArray(spec2) Then
        If IsObject(spec1) Or IsObject(spec2) Then
            CompareResumeSpecs = False
        Else
            CompareResumeSpecs = (spec1 = spec2)
        End If
        Exit Function
    End If

    If IsArray(spec1) Xor IsArray(spec2) Then
        CompareResumeSpecs = False
        Exit Function
    End If

    ' 获取边界
    Dim lb1 As Long, ub1 As Long
    Dim lb2 As Long, ub2 As Long
    On Error Resume Next
    lb1 = LBound(spec1): ub1 = UBound(spec1)
    lb2 = LBound(spec2): ub2 = UBound(spec2)
    If Err.Number <> 0 Then
        Err.Clear
        Dim test1 As Long, test2 As Long
        On Error Resume Next
        test1 = UBound(spec1)
        Dim isEmpty1 As Boolean: isEmpty1 = (Err.Number <> 0)
        Err.Clear
        test2 = UBound(spec2)
        Dim isEmpty2 As Boolean: isEmpty2 = (Err.Number <> 0)
        Err.Clear
        On Error GoTo NotEqual
        CompareResumeSpecs = (isEmpty1 And isEmpty2)
        Exit Function
    End If
    On Error GoTo NotEqual

    ' 长度不同
    If (ub1 - lb1) <> (ub2 - lb2) Then
        CompareResumeSpecs = False
        Exit Function
    End If

    ' 逐元素比较
    Dim i As Long
    For i = 0 To (ub1 - lb1)
        Dim idx1 As Long, idx2 As Long
        idx1 = lb1 + i
        idx2 = lb2 + i
        ' 检查两边是否都是 Dictionary
        Dim isDict1 As Boolean, isDict2 As Boolean
        isDict1 = (TypeName(spec1(idx1)) = "Dictionary")
        isDict2 = (TypeName(spec2(idx2)) = "Dictionary")
        If isDict1 And isDict2 Then
            Dim dict1 As Object, dict2 As Object
            Set dict1 = spec1(idx1)
            Set dict2 = spec2(idx2)
            ' 都是 Range 引用：只比较地址
            Dim isRange1 As Boolean, isRange2 As Boolean
            isRange1 = False: isRange2 = False
            If dict1.Exists("isRange") Then isRange1 = (dict1("isRange") = True)
            If dict2.Exists("isRange") Then isRange2 = (dict2("isRange") = True)
            If isRange1 And isRange2 Then
                If dict1("address") <> dict2("address") Then
                    CompareResumeSpecs = False
                    Exit Function
                End If
                If dict1("workbook") <> dict2("workbook") Then
                    CompareResumeSpecs = False
                    Exit Function
                End If
                If dict1("worksheet") <> dict2("worksheet") Then
                    CompareResumeSpecs = False
                    Exit Function
                End If
                GoTo NextParam
            End If
            ' 其他 Dictionary
            If dict1.Count <> dict2.Count Then
                CompareResumeSpecs = False
                Exit Function
            End If
            GoTo NextParam
        ElseIf isDict1 Or isDict2 Then
            ' 一个是 Dictionary 一个不是
            CompareResumeSpecs = False
            Exit Function
        End If
        ' 非 Dictionary 比较
        If TypeName(spec1(idx1)) <> TypeName(spec2(idx2)) Then
            CompareResumeSpecs = False
            Exit Function
        End If
        If Not IsObject(spec1(idx1)) And Not IsObject(spec2(idx2)) Then
            If spec1(idx1) <> spec2(idx2) Then
                CompareResumeSpecs = False
                Exit Function
            End If
        End If

NextParam:
    Next i
    CompareResumeSpecs = True
    Exit Function
NotEqual:
    CompareResumeSpecs = False
End Function

