' ============================================
' WorkbookRuntime.cls - 工作簿运行时
' ============================================
' 设计原因：
' 1. 每个 Workbook 一个实例
' 2. 封装 Lua VM、TaskTable、协程调度
' 3. 实现 IRuntime 接口，供调度器驱动
' 4. 处理热更新（模块级，非 VM 级）
' ============================================

Option Explicit
Implements IRuntime

' ===== Lua 5.4 API 声明 =====
#If VBA7 Then
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
    Private Declare PtrSafe Function lua_newthread Lib "lua54.dll" (ByVal L As LongPtr) As LongPtr
    Private Declare PtrSafe Function lua_resume Lib "lua54.dll" (ByVal L As LongPtr, ByVal from As LongPtr, ByVal narg As Long, ByVal nres As LongPtr) As Long
    Private Declare PtrSafe Sub lua_xmove Lib "lua54.dll" (ByVal fromL As LongPtr, ByVal toL As LongPtr, ByVal n As Long)
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As LongPtr)
#End If

' ===== Lua 类型常量 =====
Private Const LUA_TNIL = 0
Private Const LUA_TBOOLEAN = 1
Private Const LUA_TNUMBER = 3
Private Const LUA_TSTRING = 4
Private Const LUA_TTABLE = 5
Private Const LUA_TFUNCTION = 6
Private Const LUA_OK = 0
Private Const LUA_YIELD = 1

' ===== 运行时状态（私有，外部不可访问）=====
Private m_LuaState As LongPtr
Private m_Workbook As Workbook
Private m_WorkbookKey As String
Private m_FunctionsPath As String
Private m_LastModified As Date

' Task 管理（私有 Dictionary）
Private m_TaskTable As Object          ' taskId → Task 数据字典
Private m_ActiveQueue As Object        ' taskId → True（活跃任务）
Private m_RoundRobinCursor As Long

' 配置
Private m_MaxTasksPerTick As Long
Private m_HotReloadEnabled As Boolean

' ===== 配置常量 =====
Private Const DEFAULT_MAX_TASKS_PER_TICK As Long = 1
Private Const DEFAULT_HOT_RELOAD_ENABLED As Boolean = True

' ============================================
' 初始化与清理
' ============================================

Private Sub Class_Initialize()
    ' 创建 Lua VM
    m_LuaState = luaL_newstate()
    If m_LuaState = 0 Then
        Err.Raise vbObjectError + 1, , "无法创建 Lua VM"
    End If
    luaL_openlibs m_LuaState
    
    ' 初始化 Task 管理
    Set m_TaskTable = CreateObject("Scripting.Dictionary")
    Set m_ActiveQueue = CreateObject("Scripting.Dictionary")
    m_RoundRobinCursor = 0
    
    ' 配置
    m_MaxTasksPerTick = DEFAULT_MAX_TASKS_PER_TICK
    m_HotReloadEnabled = DEFAULT_HOT_RELOAD_ENABLED
    m_LastModified = #1/1/1900#
End Sub

Private Sub Class_Terminate()
    IRuntime_Dispose
End Sub

' 绑定工作簿
Public Sub BindWorkbook(wb As Workbook)
    Set m_Workbook = wb
    m_WorkbookKey = CoreRegistry.GetWorkbookKey(wb)
    m_FunctionsPath = wb.Path & "\functions.lua"
    
    ' 尝试加载 functions.lua
    TryLoadFunctions
End Sub

' ============================================
' IRuntime 接口实现
' ============================================

Private Function IRuntime_HasRunnable() As Boolean
    IRuntime_HasRunnable = (m_ActiveQueue.Count > 0)
End Function

Private Sub IRuntime_Tick()
    On Error Resume Next
    
    ' Round-robin 调度
    Dim taskIds() As Variant
    ReDim taskIds(0 To m_ActiveQueue.Count - 1)
    
    Dim idx As Long
    Dim taskId As Variant
    For Each taskId In m_ActiveQueue.Keys
        taskIds(idx) = taskId
        idx = idx + 1
    Next
    
    Dim total As Long
    total = UBound(taskIds) + 1
    If total = 0 Then Exit Sub
    
    ' 执行配额内的任务
    Dim executed As Long
    Dim cur As Long
    cur = m_RoundRobinCursor Mod total
    
    Dim tasksToRemove As Object
    Set tasksToRemove = CreateObject("System.Collections.ArrayList")
    
    Do While executed < m_MaxTasksPerTick And executed < total
        taskId = taskIds(cur)
        
        If m_TaskTable.Exists(CStr(taskId)) Then
            ResumeTask CStr(taskId)
            executed = executed + 1
            
            ' 检查是否完成
            Dim task As Object
            Set task = m_TaskTable(CStr(taskId))
            If task("status") = "done" Or task("status") = "error" Then
                tasksToRemove.Add taskId
            End If
        Else
            tasksToRemove.Add taskId
        End If
        
        cur = (cur + 1) Mod total
    Loop
    
    m_RoundRobinCursor = cur
    
    ' 清理完成任务
    Dim i As Long
    For i = 0 To tasksToRemove.Count - 1
        If m_ActiveQueue.Exists(tasksToRemove(i)) Then
            m_ActiveQueue.Remove tasksToRemove(i)
        End If
    Next
End Sub

Private Sub IRuntime_Dispose()
    ' 清理所有任务
    If Not m_TaskTable Is Nothing Then
        m_TaskTable.RemoveAll
        m_ActiveQueue.RemoveAll
    End If
    
    ' 关闭 Lua VM
    If m_LuaState <> 0 Then
        lua_close m_LuaState
        m_LuaState = 0
    End If
End Sub

' ============================================
' 公共方法（供 UDF/菜单调用）
' ============================================

' 创建任务
Public Function CreateTask(cellAddr As String, funcName As String, _
                          startArgs As Variant, resumeSpec As Variant) As String
    ' 生成 taskId
    Dim taskId As String
    taskId = Scheduler.GenerateTaskId(m_WorkbookKey, cellAddr)
    
    ' 创建 Task 数据字典
    Dim task As Object
    Set task = CreateObject("Scripting.Dictionary")
    task("funcName") = funcName
    task("cellAddr") = cellAddr
    task("startArgs") = startArgs
    task("resumeSpec") = resumeSpec
    task("status") = "defined"
    task("progress") = 0
    task("message") = Empty
    task("value") = Empty
    task("error") = ""
    task("coThread") = 0
    
    m_TaskTable(taskId) = task
    
    ' 注册到全局路由表
    CoreRegistry.RegisterTask taskId, Me
    
    CreateTask = taskId
End Function

' 启动任务
Public Sub StartTask(taskId As String)
    If Not m_TaskTable.Exists(taskId) Then
        Err.Raise vbObjectError + 2, , "任务不存在"
    End If
    
    Dim task As Object
    Set task = m_TaskTable(taskId)
    
    If task("status") <> "defined" Then
        Err.Raise vbObjectError + 3, , "任务状态不是 defined"
    End If
    
    ' 创建协程
    Dim coThread As LongPtr
    coThread = lua_newthread(m_LuaState)
    If coThread = 0 Then
        task("status") = "error"
        task("error") = "无法创建协程"
        Exit Sub
    End If
    task("coThread") = coThread
    
    ' 获取函数
    lua_getglobal m_LuaState, task("funcName")
    If lua_type(m_LuaState, -1) <> LUA_TFUNCTION Then
        task("status") = "error"
        task("error") = "函数不存在: " & task("funcName")
        lua_settop m_LuaState, 0
        Exit Sub
    End If
    
    ' 移动函数到协程
    lua_xmove m_LuaState, coThread, 1
    
    ' 推入 cellAddr
    lua_pushstring coThread, task("cellAddr")
    
    ' 推入启动参数
    Dim nargs As Long
    nargs = 1
    If IsArray(task("startArgs")) Then
        Dim i As Long
        For i = LBound(task("startArgs")) To UBound(task("startArgs"))
            PushValue coThread, task("startArgs")(i)
            nargs = nargs + 1
        Next
    End If
    
    ' 首次 resume
    Dim nres As LongPtr
    Dim result As Long
    result = lua_resume(coThread, m_LuaState, nargs, VarPtr(nres))
    
    HandleResumeResult taskId, result, CLng(nres)
    
    ' 加入活跃队列
    If task("status") = "yielded" Then
        m_ActiveQueue(taskId) = True
        Scheduler.StartScheduler
    End If
End Sub

' 获取任务字段
Public Function GetTaskField(taskId As String, field As String) As Variant
    If Not m_TaskTable.Exists(taskId) Then
        GetTaskField = "#ERROR: 任务不存在"
        Exit Function
    End If
    
    Dim task As Object
    Set task = m_TaskTable(taskId)
    
    Select Case LCase(field)
        Case "status": GetTaskField = task("status")
        Case "progress": GetTaskField = task("progress")
        Case "message": GetTaskField = task("message")
        Case "value": GetTaskField = task("value")
        Case "error": GetTaskField = task("error")
        Case Else: GetTaskField = "#ERROR: 未知字段"
    End Select
End Function

' 暂停任务
Public Sub PauseTask(taskId As String)
    If m_ActiveQueue.Exists(taskId) Then
        m_ActiveQueue.Remove taskId
        m_TaskTable(taskId)("status") = "paused"
    End If
End Sub

' 恢复任务
Public Sub ResumeTaskManual(taskId As String)
    If Not m_TaskTable.Exists(taskId) Then Exit Sub
    
    Dim task As Object
    Set task = m_TaskTable(taskId)
    
    If task("status") = "yielded" Or task("status") = "paused" Then
        If Not m_ActiveQueue.Exists(taskId) Then
            m_ActiveQueue(taskId) = True
            Scheduler.StartScheduler
        End If
    End If
End Sub

' 终止任务
Public Sub TerminateTask(taskId As String)
    If Not m_TaskTable.Exists(taskId) Then Exit Sub
    
    If m_ActiveQueue.Exists(taskId) Then
        m_ActiveQueue.Remove taskId
    End If
    
    m_TaskTable.Remove taskId
    CoreRegistry.UnregisterTask taskId
End Sub

' ============================================
' Lua 热更新（模块级）
' ============================================

' 手动重载 functions.lua
Public Function ReloadFunctions() As Boolean
    On Error GoTo ErrorHandler
    
    ' 设计关键：不重建 VM，只重载模块
    ' 1. package.loaded["functions"] = nil
    ' 2. require("functions")
    
    Dim code As String
    code = "package.loaded['functions'] = nil; require('functions')"
    
    Dim result As Long
    result = luaL_loadstring(m_LuaState, code)
    If result = 0 Then result = lua_pcallk(m_LuaState, 0, 0, 0, 0, 0)
    
    If result <> 0 Then
        Dim errMsg As String
        errMsg = GetStringFromState(m_LuaState, -1)
        lua_settop m_LuaState, 0
        CoreUI.LogError "热更新失败: " & errMsg
        ReloadFunctions = False
        Exit Function
    End If
    
    m_LastModified = FileDateTime(m_FunctionsPath)
    ReloadFunctions = True
    Exit Function
    
ErrorHandler:
    ReloadFunctions = False
End Function

' 自动检查热更新
Private Sub CheckHotReload()
    If Not m_HotReloadEnabled Then Exit Sub
    
    On Error Resume Next
    Dim currentModified As Date
    currentModified = FileDateTime(m_FunctionsPath)
    
    If Err.Number <> 0 Or currentModified = m_LastModified Then Exit Sub
    
    ReloadFunctions
End Sub

' 初次加载
Private Sub TryLoadFunctions()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(m_FunctionsPath) Then Exit Sub
    
    Dim code As String
    code = "require('functions')"
    
    Dim result As Long
    result = luaL_loadstring(m_LuaState, code)
    If result = 0 Then result = lua_pcallk(m_LuaState, 0, 0, 0, 0, 0)
    
    If result = 0 Then
        m_LastModified = FileDateTime(m_FunctionsPath)
    End If
End Sub

' ============================================
' 协程调度（私有）
' ============================================

Private Sub ResumeTask(taskId As String)
    On Error GoTo ErrorHandler
    
    Dim task As Object
    Set task = m_TaskTable(taskId)
    
    If task("status") <> "yielded" Then Exit Sub
    
    Dim coThread As LongPtr
    coThread = task("coThread")
    
    lua_settop coThread, 0
    
    ' 推入 resume 参数
    Dim nargs As Long
    nargs = 0
    If IsArray(task("resumeSpec")) Then
        Dim i As Long
        For i = LBound(task("resumeSpec")) To UBound(task("resumeSpec"))
            PushValue coThread, task("resumeSpec")(i)
            nargs = nargs + 1
        Next
    End If
    
    ' Resume
    Dim nres As LongPtr
    Dim result As Long
    result = lua_resume(coThread, m_LuaState, nargs, VarPtr(nres))
    
    HandleResumeResult taskId, result, CLng(nres)
    Exit Sub
    
ErrorHandler:
    task("status") = "error"
    task("error") = "Resume 错误: " & Err.Description
End Sub

Private Sub HandleResumeResult(taskId As String, result As Long, nres As Long)
    Dim task As Object
    Set task = m_TaskTable(taskId)
    
    Dim coThread As LongPtr
    coThread = task("coThread")
    
    Select Case result
        Case LUA_OK
            task("status") = "done"
            task("progress") = 100
            
            If nres > 0 Then
                Dim retData As Variant
                retData = GetValue(coThread, -1)
                ParseYieldReturn task, retData, True
            End If
            
        Case LUA_YIELD
            If nres > 0 Then
                Dim yieldData As Variant
                yieldData = GetValue(coThread, -1)
                ParseYieldReturn task, yieldData, False
            End If
            
            If task("status") <> "done" And task("status") <> "error" Then
                task("status") = "yielded"
            End If
            
        Case Else
            task("status") = "error"
            If nres > 0 Then
                task("error") = GetStringFromState(coThread, -1)
            Else
                task("error") = "协程错误: 代码 " & result
            End If
    End Select
    
    lua_settop coThread, 0
End Sub

' 解析 yield/return 字典
Private Sub ParseYieldReturn(task As Object, data As Variant, isFinal As Boolean)
    On Error Resume Next
    
    If Not IsArray(data) Then
        task("value") = data
        Exit Sub
    End If
    
    ' 检查是否为字典格式
    Dim isDictionary As Boolean
    isDictionary = False
    
    Dim cols As Long
    cols = UBound(data, 2) - LBound(data, 2) + 1
    If Err.Number = 0 And cols = 2 Then
        isDictionary = True
    End If
    On Error GoTo 0
    
    If isDictionary Then
        Dim i As Long
        For i = LBound(data, 1) To UBound(data, 1)
            Dim key As String
            Dim value As Variant
            
            key = LCase(Trim(CStr(data(i, 1))))
            value = data(i, 2)
            
            Select Case key
                Case "status"
                    If Not isFinal Then
                        Dim statusVal As String
                        statusVal = LCase(Trim(CStr(value)))
                        Select Case statusVal
                            Case "yielded", "done", "error"
                                task("status") = statusVal
                            Case Else
                                task("status") = "yielded"
                        End Select
                    End If
                Case "progress"
                    On Error Resume Next
                    task("progress") = CDbl(value)
                    On Error GoTo 0
                Case "message"
                    task("message") = value
                Case "value"
                    task("value") = value
            End Select
        Next
    Else
        task("value") = data
    End If
End Sub

' ============================================
' Lua 栈操作工具（私有）
' ============================================

Private Sub PushValue(ByVal L As LongPtr, ByVal value As Variant)
    If TypeName(value) = "Range" Then
        Dim rng As Range
        Set rng = value
        If rng.Cells.Count = 1 Then
            PushValue L, rng.value
        Else
            PushValue L, rng.value
        End If
        Exit Sub
    End If
    
    If IsArray(value) Then
        PushArray L, value
        Exit Sub
    End If
    
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

Private Sub PushArray(ByVal L As LongPtr, arr As Variant)
    Dim i As Long, j As Long
    Dim rows As Long, cols As Long
    
    On Error Resume Next
    rows = UBound(arr, 1) - LBound(arr, 1) + 1
    cols = UBound(arr, 2) - LBound(arr, 2) + 1
    
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        rows = UBound(arr) - LBound(arr) + 1
        
        lua_createtable L, rows, 0
        For i = LBound(arr) To UBound(arr)
            PushValue L, arr(i)
            lua_rawseti L, -2, i - LBound(arr) + 1
        Next
        Exit Sub
    End If
    On Error GoTo 0
    
    lua_createtable L, rows, 0
    For i = LBound(arr, 1) To UBound(arr, 1)
        lua_createtable L, cols, 0
        For j = LBound(arr, 2) To UBound(arr, 2)
            PushValue L, arr(i, j)
            lua_rawseti L, -2, j - LBound(arr, 2) + 1
        Next
        lua_rawseti L, -2, i - LBound(arr, 1) + 1
    Next
End Sub

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

Private Function GetStringFromState(ByVal L As LongPtr, ByVal idx As Long) As String
    Dim ptr As LongPtr
    Dim length As Long
    
    ptr = lua_tolstring(L, idx, VarPtr(length))
    If ptr = 0 Or length = 0 Then
        GetStringFromState = ""
        Exit Function
    End If
    
    Dim bytes() As Byte
    ReDim bytes(0 To length - 1)
    CopyMemory bytes(0), ByVal ptr, length
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "UTF-8"
    GetStringFromState = stream.ReadText
    stream.Close
End Function

Private Function TableToVariant(ByVal L As LongPtr, ByVal idx As Long) As Variant
    On Error Resume Next
    
    If idx < 0 Then idx = lua_gettop(L) + idx + 1
    
    Dim length As LongPtr
    length = lua_rawlen(L, idx)
    
    If length = 0 Then
        TableToVariant = Empty
        Exit Function
    End If
    
    ' 简化实现：仅处理一维数组
    Dim arr() As Variant
    ReDim arr(1 To 1, 1 To CLng(length))
    
    Dim i As Long
    For i = 1 To CLng(length)
        lua_rawgeti L, idx, CLng(i)
        arr(1, i) = GetValue(L, -1)
        lua_settop L, -2
    Next
    
    TableToVariant = arr
End Function