' ============================================
' WorkbookRuntime.cls - 工作簿运行时
' ============================================
' 设计原因：
' 1. 每个 Workbook 一个实例
' 2. 封装 Lua VM、TaskTable、协程调度
' 3. 实现 CanRun等接口，供调度器驱动
' 4. 处理热更新（模块级，非 VM 级）
' ============================================

Option Explicit

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
    Private Declare PtrSafe Function luaL_loadfilex Lib "lua54.dll" (ByVal L As LongPtr, ByVal filename As String, ByVal mode As LongPtr) As Long
    Private Declare PtrSafe Sub lua_setglobal Lib "lua54.dll" (ByVal L As LongPtr, ByVal name As String)
    Private Declare PtrSafe Function lua_next Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long) As Long
    Private Declare PtrSafe Sub lua_pushvalue Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long)
    Private Declare PtrSafe Function lua_getfield Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long, ByVal k As String) As Long
    Private Declare PtrSafe Sub lua_setfield Lib "lua54.dll" (ByVal L As LongPtr, ByVal idx As Long, ByVal k As String)

        ' 协程 API
    Private Declare PtrSafe Function lua_newthread Lib "lua54.dll" (ByVal L As LongPtr) As LongPtr
    Private Declare PtrSafe Function lua_resume Lib "lua54.dll" (ByVal L As LongPtr, ByVal from As LongPtr, ByVal narg As Long, ByVal nres As LongPtr) As Long
    Private Declare PtrSafe Function lua_status Lib "lua54.dll" (ByVal L As LongPtr) As Long
    Private Declare PtrSafe Sub lua_xmove Lib "lua54.dll" (ByVal fromL As LongPtr, ByVal toL As LongPtr, ByVal n As Long)
    ' 系统 API
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As LongPtr)
#End if

' Lua 类型常量
Private Const LUA_TNIL = 0
Private Const LUA_TBOOLEAN = 1
Private Const LUA_TNUMBER = 3
Private Const LUA_TSTRING = 4
Private Const LUA_TTABLE = 5
Private Const LUA_TFUNCTION = 6
Private Const LUA_OK = 0
Private Const LUA_YIELD = 1
Private Const LUA_ERRRUN = 2
' 运行时状态（私有，外部不可访问）
Private m_LuaState As LongPtr
Private m_Workbook As Workbook
Private m_WorkbookKey As String
Private m_LastModified As Date
' 路径管理
Private m_SystemMainPath As String      ' 系统 main.lua 路径
Private m_FunctionsPath As String       ' 工作簿 functions.lua 路径
Private m_LastMainModified As Date      ' main.lua 最后修改时间
Private m_LastFunctionsModified As Date ' functions.lua 最后修改时间
' Task 管理（私有 Dictionary）
Private m_TaskTable As Object          ' taskId → Task 数据字典
Private m_ActiveQueue As Object        ' taskId → True（活跃任务）
Private m_RoundRobinCursor As Long
' 配置
Private m_MaxTasksPerTick As Long
Private m_HotReloadEnabled As Boolean
' 配置常量
Private Const DEFAULT_MAX_TASKS_PER_TICK As Long = 1
Private Const DEFAULT_HOT_RELOAD_ENABLED As Boolean = True

' ====初始化与清理====
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
' 绑定工作簿
Public Sub BindWorkbook(wb As Workbook)
    Set m_Workbook = wb
    m_WorkbookKey = CoreRegistry.GetWorkbookKey(wb)
    
    ' 设置路径
    ' 假设系统路径为："C:\lua\main.lua"
    m_SystemMainPath = "C:\Lua\LuaForExcel\main.lua"
    m_FunctionsPath = wb.Path & "\functions.lua"
    
    ' 初始化修改时间
    m_LastMainModified = #1/1/1900#
    m_LastFunctionsModified = #1/1/1900#
    
    ' 首次加载
    InitialLoadLuaFiles
End Sub

' ====对Scheduler的接口====
Public Function HasRunnable() As Boolean
    HasRunnable = (m_ActiveQueue.Count > 0)
End Function

Public Sub Tick()
    On Error GoTo ErrorHandler

    If m_ActiveQueue Is Nothing Or m_ActiveQueue.Count = 0 Then Exit Sub

    CheckHotReload

    ' 快照活跃任务
    Dim taskIds() As Variant
    ReDim taskIds(0 To m_ActiveQueue.Count - 1)

    Dim idx As Long
    Dim taskId As Variant
    idx = 0
    For Each taskId In m_ActiveQueue.Keys
        taskIds(idx) = taskId
        idx = idx + 1
    Next

    Dim total As Long
    total = UBound(taskIds) + 1
    If total = 0 Then Exit Sub

    ' Round-Robin 调度
    Dim executed As Long
    executed = 0

    Dim cur As Long
    cur = m_RoundRobinCursor Mod total

    Dim toRemove As Object
    Set toRemove = CreateObject("System.Collections.ArrayList")

    Do While executed < m_MaxTasksPerTick And executed < total
        taskId = taskIds(cur)

        If m_TaskTable.Exists(CStr(taskId)) Then
            ResumeTask CStr(taskId)
            executed = executed + 1

            Dim task As Object
            Set task = m_TaskTable(taskId)

            If task("status") = "done" Or task("status") = "error" Then
                toRemove.Add taskId
            End If
        Else
            toRemove.Add taskId
        End If

        cur = (cur + 1) Mod total
    Loop

    ' ===== 修复：在删除前更新游标 =====
    m_RoundRobinCursor = cur
    
    ' 如果游标超出范围，重置为 0
    If m_RoundRobinCursor >= total Then
        m_RoundRobinCursor = 0
    End If

    ' 清理非活跃任务
    Dim i As Long
    For i = 0 To toRemove.Count - 1
        If m_ActiveQueue.Exists(toRemove(i)) Then
            m_ActiveQueue.Remove toRemove(i)
        End If
    Next i

    Exit Sub

ErrorHandler:
    Debug.Print "[WorkbookRuntime.Tick] Error: " & Err.Description
End Sub

' ====公共方法（供 UDF/菜单调用）====
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
    
    ' 注册到全局路由表（包含 cellAddr 索引）
    CoreRegistry.RegisterTask taskId, Me, cellAddr
    
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
    
    ' 获取 cellAddr
    Dim task As Object
    Set task = m_TaskTable(taskId)
    Dim cellAddr As String
    cellAddr = task("cellAddr")
    
    ' 从活跃队列移除
    If m_ActiveQueue.Exists(taskId) Then
        m_ActiveQueue.Remove taskId
    End If
    
    ' 从任务表移除
    m_TaskTable.Remove taskId
    
    ' 从全局路由表注销（包含 cellAddr 索引）
    CoreRegistry.UnregisterTask taskId, cellAddr
End Sub
' 释放运行时资源
Public Sub Dispose()
    On Error Resume Next
    
    If m_LuaState <> 0 Then
        lua_close m_LuaState
        m_LuaState = 0
    End If
    
    m_TaskTable.RemoveAll
    m_ActiveQueue.RemoveAll
    
    Set m_TaskTable = Nothing
    Set m_ActiveQueue = Nothing
End Sub

' ====LuaEval LuaGet 调用接口====
' 执行 Lua 表达式
Public Function EvalExpression(expression As String) As Variant
    On Error GoTo ErrorHandler
    
    ' 检查热重载
    CheckHotReload
    
    ' 构造代码
    Dim fullCode As String
    fullCode = "return " & expression
    
    ' 加载代码
    Dim result As Long
    result = luaL_loadstring(m_LuaState, fullCode)
    If result <> 0 Then
        EvalExpression = "#LUA_SYNTAX_ERROR: " & GetStringFromState(m_LuaState, -1)
        lua_settop m_LuaState, 0
        Exit Function
    End If
    
    ' 执行代码
    result = lua_pcallk(m_LuaState, 0, 1, 0, 0, 0)
    If result <> 0 Then
        EvalExpression = "#LUA_RUNTIME_ERROR: " & GetStringFromState(m_LuaState, -1)
        lua_settop m_LuaState, 0
        Exit Function
    End If
    
    ' 获取结果
    EvalExpression = GetValue(m_LuaState, -1)
    lua_settop m_LuaState, 0
    Exit Function

ErrorHandler:
    EvalExpression = "#VBA_ERROR: " & Err.Description
    On Error Resume Next
    If m_LuaState <> 0 Then lua_settop m_LuaState, 0
End Function
' 调用 functions.lua 中的函数
Public Function CallFunction(funcName As String, args As Variant) As Variant
    On Error GoTo ErrorHandler
    
    ' 检查热重载
    CheckHotReload
    
    ' 获取函数
    lua_getglobal m_LuaState, funcName
    If lua_type(m_LuaState, -1) <> LUA_TFUNCTION Then
        lua_settop m_LuaState, 0
        CallFunction = "ERROR: 函数 '" & funcName & "' 不存在"
        Exit Function
    End If
    
    ' 推入参数
    Dim i As Long, argCount As Long
    argCount = 0
    For i = LBound(args) To UBound(args)
        PushValue m_LuaState, args(i)
        argCount = argCount + 1
    Next

    ' 调用函数
    Dim result As Long
    result = lua_pcallk(m_LuaState, argCount, -1, 0, 0, 0)
    If result <> 0 Then
        CallFunction = "#LUA_ERROR: " & GetStringFromState(m_LuaState, -1)
        lua_settop m_LuaState, 0
        Exit Function
    End If
    
    ' 处理返回值
    Dim nResults As Long
    nResults = lua_gettop(m_LuaState)
    
    If nResults = 0 Then
        CallFunction = Empty
    ElseIf nResults = 1 Then
        CallFunction = GetValue(m_LuaState, -1)
    Else
        ' 多个返回值 - 返回为水平数组
        Dim results() As Variant
        ReDim results(1 To 1, 1 To nResults)
        For i = 1 To nResults
            results(1, i) = GetValue(m_LuaState, i)
        Next i
        CallFunction = results
    End If
    
    lua_settop m_LuaState, 0
    Exit Function

ErrorHandler:
    CallFunction = "#VBA_ERROR: " & Err.Description
    On Error Resume Next
    If m_LuaState <> 0 Then lua_settop m_LuaState, 0
End Function

' ====Lua热更新（模块级）====
' 初始加载（工作簿打开时）
Private Sub InitialLoadLuaFiles()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 加载系统 main.lua（必需）
    If fso.FileExists(m_SystemMainPath) Then
        If Not LoadLuaFile(m_SystemMainPath, "main") Then
            MsgBox "警告：系统 main.lua 加载失败" & vbCrLf & _
                   "路径：" & m_SystemMainPath, _
                   vbExclamation, "初始化警告"
        Else
            m_LastMainModified = FileDateTime(m_SystemMainPath)
        End If
    Else
        MsgBox "错误：找不到系统 main.lua" & vbCrLf & _
               "路径：" & m_SystemMainPath, _
               vbCritical, "初始化失败"
    End If
    
    ' 加载工作簿 functions.lua（可选）
    If fso.FileExists(m_FunctionsPath) Then
        If LoadLuaFile(m_FunctionsPath, "functions") Then
            m_LastFunctionsModified = FileDateTime(m_FunctionsPath)
        End If
    End If
End Sub
' 热更新检查（在 Tick 中调用）
Private Sub CheckHotReload()
    If Not m_HotReloadEnabled Then Exit Sub
    
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 检查 main.lua
    If fso.FileExists(m_SystemMainPath) Then
        Dim currentMainModified As Date
        currentMainModified = FileDateTime(m_SystemMainPath)
        
        If currentMainModified <> m_LastMainModified Then
            If ReloadLuaFile(m_SystemMainPath, "main") Then
                m_LastMainModified = currentMainModified
                Debug.Print "[HotReload] main.lua 已更新"
            End If
        End If
    End If
    
    ' 检查 functions.lua
    If fso.FileExists(m_FunctionsPath) Then
        Dim currentFunctionsModified As Date
        currentFunctionsModified = FileDateTime(m_FunctionsPath)
        
        If currentFunctionsModified <> m_LastFunctionsModified Then
            If ReloadLuaFile(m_FunctionsPath, "functions") Then
                m_LastFunctionsModified = currentFunctionsModified
                Debug.Print "[HotReload] functions.lua 已更新"
            End If
        End If
    End If
End Sub
' 手动重载接口
' 重载 functions.lua
Public Function ReloadFunctions() As Boolean
    ReloadFunctions = ReloadLuaFile(m_FunctionsPath, "functions")
    If ReloadFunctions Then
        m_LastFunctionsModified = FileDateTime(m_FunctionsPath)
    End If
End Function
' 重载 main.lua（管理员操作）
Public Function ReloadMainLua() As Boolean
    ReloadMainLua = ReloadLuaFile(m_SystemMainPath, "main")
    If ReloadMainLua Then
        m_LastMainModified = FileDateTime(m_SystemMainPath)
    End If
End Function
' 加载 Lua 文件（无验证，用于初始加载）
Private Function LoadLuaFile(filePath As String, moduleName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim result As Long
    result = luaL_loadfilex(m_LuaState, filePath, 0)
    If result = 0 Then result = lua_pcallk(m_LuaState, 0, 0, 0, 0, 0)
    
    If result <> 0 Then
        Dim errMsg As String
        errMsg = GetStringFromState(m_LuaState, -1)
        lua_settop m_LuaState, 0
        
        Debug.Print "[LoadError] " & moduleName & ": " & errMsg
        LoadLuaFile = False
        Exit Function
    End If
    
    LoadLuaFile = True
    Exit Function
    
ErrorHandler:
    Debug.Print "[LoadError] VBA Error: " & Err.Description
    LoadLuaFile = False
End Function
' 重载 Lua 文件（带验证）
Private Function ReloadLuaFile(filePath As String, moduleName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' ===== 第一步：在临时状态中验证语法 =====
    Dim tempL As LongPtr
    tempL = luaL_newstate()
    If tempL = 0 Then
        ReloadLuaFile = False
        Exit Function
    End If
    
    luaL_openlibs tempL
    
    ' 验证加载
    Dim result As Long
    result = luaL_loadfilex(tempL, filePath, 0)
    If result = 0 Then result = lua_pcallk(tempL, 0, 0, 0, 0, 0)
    
    If result <> 0 Then
        Dim errMsg As String
        errMsg = GetStringFromState(tempL, -1)
        lua_close tempL
        
        MsgBox moduleName & ".lua 存在语法错误：" & vbCrLf & vbCrLf & _
               errMsg & vbCrLf & vbCrLf & _
               "旧版本将继续使用", _
               vbCritical, "热更新失败"
        
        ReloadLuaFile = False
        Exit Function
    End If
    
    lua_close tempL
    
    ' ===== 第二步：清除旧模块缓存 =====
    ' package.loaded[moduleName] = nil
    lua_getglobal m_LuaState, "package"
    lua_getfield m_LuaState, -1, "loaded"
    lua_pushnil m_LuaState
    lua_setfield m_LuaState, -2, moduleName
    lua_settop m_LuaState, 0
    
    ' ===== 第三步：在主状态中重新加载 =====
    result = luaL_loadfilex(m_LuaState, filePath, 0)
    If result = 0 Then result = lua_pcallk(m_LuaState, 0, 0, 0, 0, 0)
    
    If result <> 0 Then
        errMsg = GetStringFromState(m_LuaState, -1)
        lua_settop m_LuaState, 0
        
        MsgBox "主状态加载 " & moduleName & ".lua 失败：" & vbCrLf & vbCrLf & _
               errMsg, _
               vbCritical, "严重错误"
        
        ReloadLuaFile = False
        Exit Function
    End If
    
    ReloadLuaFile = True
    Exit Function
    
ErrorHandler:
    If tempL <> 0 Then lua_close tempL
    MsgBox "重载过程发生 VBA 错误: " & Err.Description, vbCritical, "严重错误"
    ReloadLuaFile = False
End Function

' ====协程调度（私有）====
Private Sub ResumeTask(taskId As String)
    On Error GoTo ErrorHandler

    Dim task As Object
    Set task = m_TaskTable(taskId)

    If task("status") <> "yielded" Then Exit Sub

    Dim coThread As LongPtr
    coThread = task("coThread")

    lua_settop coThread, 0

    Dim nargs As Long
    nargs = 0

    ' 1. 优先使用 resumeSpec
    If task.Exists("resumeSpec") Then
        If IsArray(task("resumeSpec")) Then
            Dim i As Long
            For i = LBound(task("resumeSpec")) To UBound(task("resumeSpec"))
                Dim spec As Variant
                spec = task("resumeSpec")(i)

                If IsMissing(spec) Or IsEmpty(spec) Then GoTo NextSpec

                ' Range 对象
                If TypeName(spec) = "Range" Then
                    PushValue coThread, spec.Value
                    nargs = nargs + 1

                ' 字符串地址
                ElseIf VarType(spec) = vbString Then
                    On Error Resume Next
                    Dim rngVal As Variant
                    rngVal = m_Workbook.Application.Range(spec).Value
                    If Err.Number = 0 Then
                        PushValue coThread, rngVal
                        nargs = nargs + 1
                    End If
                    On Error GoTo ErrorHandler

                ' 其他值
                Else
                    PushValue coThread, spec
                    nargs = nargs + 1
                End If
NextSpec:
            Next
        End If
    End If

    ' 2. 修复：正确获取任务单元格的值
    If nargs = 0 Then
        Dim cellAddr As String
        cellAddr = task("cellAddr")
        
        On Error Resume Next
        Dim cellVal As Variant
        cellVal = m_Workbook.Application.Range(cellAddr).Value
        
        If Err.Number = 0 Then
            If Not IsEmpty(cellVal) And Not IsError(cellVal) Then
                PushValue coThread, cellVal
                nargs = 1
            End If
        End If
        On Error GoTo ErrorHandler
    End If

    ' 3. Resume 协程
    Dim nres As LongPtr
    Dim result As Long

    result = lua_resume(coThread, m_LuaState, nargs, VarPtr(nres))

    HandleResumeResult taskId, result, CLng(nres)
    Exit Sub

ErrorHandler:
    task("status") = "error"
    task("error") = "Resume 错误: " & Err.Description
End Sub
' 处理 yield/return 结果
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

' ==== Lua 栈操作工具（私有） ====
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
    On Error GoTo ErrorHandler

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
    Exit Function
ErrorHandler:
    GetStringFromState = ""
    If Not stream Is Nothing Then
        On Error Resume Next
        stream.Close
    End If
End Function

Private Function TableToVariant(ByVal L As LongPtr, ByVal idx As Long) As Variant
    On Error GoTo ErrorHandler
    
    If idx < 0 Then idx = lua_gettop(L) + idx + 1
    
    Dim length As LongPtr
    length = lua_rawlen(L, idx)
    
    ' length = 0 → 可能是字典
    If length = 0 Then
        lua_pushnil L
        If lua_next(L, idx) <> 0 Then
            lua_settop L, lua_gettop(L) - 2
            TableToVariant = TableToDictArray(L, idx)
        Else
            TableToVariant = Empty
        End If
        Exit Function
    End If
    
    ' 判断是否纯数组
    Dim isPureArray As Boolean
    isPureArray = True
    
    lua_pushnil L
    Do While lua_next(L, idx) <> 0
        If lua_type(L, -2) <> LUA_TNUMBER Then
            isPureArray = False
            lua_settop L, lua_gettop(L) - 2
            Exit Do
        End If
        lua_settop L, -2
    Loop
    
    If Not isPureArray Then
        TableToVariant = TableToDictArray(L, idx)
        Exit Function
    End If
    
    ' 纯数组：判断是否二维
    lua_rawgeti L, idx, 1
    Dim is2D As Boolean
    is2D = (lua_type(L, -1) = LUA_TTABLE)
    lua_settop L, -2
    
    Dim i As Long, j As Long
    
    If is2D Then
        lua_rawgeti L, idx, 1
        Dim cols As LongPtr
        cols = lua_rawlen(L, -1)
        lua_settop L, -2
        
        Dim arr2D() As Variant
        ReDim arr2D(1 To CLng(length), 1 To CLng(cols))
        
        For i = 1 To CLng(length)
            lua_rawgeti L, idx, i
            For j = 1 To CLng(cols)
                lua_rawgeti L, -1, j
                arr2D(i, j) = GetValue(L, -1)
                lua_settop L, -2
            Next
            lua_settop L, -2
        Next
        
        TableToVariant = arr2D
    Else
        Dim arr1D() As Variant
        ReDim arr1D(1 To 1, 1 To CLng(length))
        
        For i = 1 To CLng(length)
            lua_rawgeti L, idx, i
            arr1D(1, i) = GetValue(L, -1)
            lua_settop L, -2
        Next
        
        TableToVariant = arr1D
    End If
    
    Exit Function

ErrorHandler:
    TableToVariant = "#TABLE_ERROR: " & Err.Description
End Function

Private Function TableToDictArray(ByVal L As LongPtr, ByVal idx As Long) As Variant
    On Error GoTo ErrorHandler
    
    If idx < 0 Then idx = lua_gettop(L) + idx + 1
    
    Dim count As Long
    count = 0
    
    lua_pushnil L
    Do While lua_next(L, idx) <> 0
        count = count + 1
        lua_settop L, -2
    Loop
    
    If count = 0 Then
        TableToDictArray = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To count, 1 To 2)
    
    Dim i As Long
    i = 1
    
    lua_pushnil L
    Do While lua_next(L, idx) <> 0
        Select Case lua_type(L, -2)
            Case LUA_TSTRING
                result(i, 1) = GetStringFromState(L, -2)
            Case LUA_TNUMBER
                result(i, 1) = lua_tonumberx(L, -2, 0)
            Case LUA_TBOOLEAN
                result(i, 1) = (lua_toboolean(L, -2) <> 0)
            Case Else
                result(i, 1) = "#KEY_TYPE_" & lua_type(L, -2)
        End Select
        
        result(i, 2) = GetValue(L, -1)
        i = i + 1
        lua_settop L, -2
    Loop
    
    TableToDictArray = result
    Exit Function

ErrorHandler:
    TableToDictArray = "#DICT_ERROR: " & Err.Description
End Function

' ====给LuaMenu的接口====
' 设置热重载开关
Public Sub SetHotReloadEnabled(enabled As Boolean)
    m_HotReloadEnabled = enabled
End Sub
' 获取热重载状态
Public Function GetHotReloadEnabled() As Boolean
    GetHotReloadEnabled = m_HotReloadEnabled
End Function
' 清理已完成/错误的任务
Public Function CleanupFinishedTasks() As Long
    Dim count As Long
    count = 0
    
    Dim toRemove As Object
    Set toRemove = CreateObject("Scripting.Dictionary")
    
    Dim taskId As Variant
    Dim task As Object
    
    For Each taskId In m_TaskTable.Keys
        Set task = m_TaskTable(taskId)
        If task("status") = "done" Or task("status") = "error" Then
            toRemove(taskId) = True
            count = count + 1
        End If
    Next
    
    For Each taskId In toRemove.Keys
        TerminateTask CStr(taskId)
    Next
    
    CleanupFinishedTasks = count
End Function
' 清除所有任务
Public Sub ClearAllTasks()
    On Error GoTo ErrorHandler

    If m_TaskTable Is Nothing Or m_TaskTable.Count = 0 Then Exit Sub

    ' 先快照所有 taskId，避免遍历中修改 Dictionary
    Dim taskIds() As Variant
    taskIds = m_TaskTable.Keys

    Dim i As Long
    For i = LBound(taskIds) To UBound(taskIds)
        TerminateTask CStr(taskIds(i))
    Next i

    ' 兜底清理（理论上应已清空）
    If Not m_ActiveQueue Is Nothing Then
        m_ActiveQueue.RemoveAll
    End If

    ' 状态复位
    m_RoundRobinCursor = 0

    Exit Sub

ErrorHandler:
    Debug.Print "[ClearAllTasks] Error: " & Err.Description
End Sub

' 获取任务统计信息
Public Function GetTaskStats() As String
    Dim total As Long, defined As Long, yielded As Long
    Dim done As Long, errored As Long, paused As Long
    
    total = m_TaskTable.Count
    
    Dim taskId As Variant
    Dim task As Object
    
    For Each taskId In m_TaskTable.Keys
        Set task = m_TaskTable(taskId)
        Select Case task("status")
            Case "defined": defined = defined + 1
            Case "yielded": yielded = yielded + 1
            Case "done": done = done + 1
            Case "error": errored = errored + 1
            Case "paused": paused = paused + 1
        End Select
    Next
    
    GetTaskStats = "任务统计" & vbCrLf & _
                   "总数: " & total & vbCrLf & _
                   "已定义: " & defined & vbCrLf & _
                   "运行中: " & yielded & vbCrLf & _
                   "已暂停: " & paused & vbCrLf & _
                   "已完成: " & done & vbCrLf & _
                   "错误: " & errored
End Function
' 获取文件状态信息（用于调试）
Public Function GetFileStatus() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim msg As String
    msg = "Lua 文件状态" & vbCrLf & vbCrLf
    
    ' main.lua 状态
    msg = msg & "系统 main.lua:" & vbCrLf
    msg = msg & "  路径: " & m_SystemMainPath & vbCrLf
    If fso.FileExists(m_SystemMainPath) Then
        msg = msg & "  状态: 存在" & vbCrLf
        msg = msg & "  最后加载: " & m_LastMainModified & vbCrLf
        msg = msg & "  当前修改: " & FileDateTime(m_SystemMainPath) & vbCrLf
    Else
        msg = msg & "  状态: 不存在" & vbCrLf
    End If
    
    msg = msg & vbCrLf
    
    ' functions.lua 状态
    msg = msg & "工作簿 functions.lua:" & vbCrLf
    msg = msg & "  路径: " & m_FunctionsPath & vbCrLf
    If fso.FileExists(m_FunctionsPath) Then
        msg = msg & "  状态: 存在" & vbCrLf
        msg = msg & "  最后加载: " & m_LastFunctionsModified & vbCrLf
        msg = msg & "  当前修改: " & FileDateTime(m_FunctionsPath) & vbCrLf
    Else
        msg = msg & "  状态: 不存在" & vbCrLf
    End If
    
    msg = msg & vbCrLf
    msg = msg & "热更新: " & IIf(m_HotReloadEnabled, "已启用", "已禁用")
    
    GetFileStatus = msg
End Function

Public Function GetAllTasksInfo() As String
    Dim msg As String
    Dim taskId As Variant
    Dim task As Object
    
    Dim total As Long
    Dim defined As Long, yielded As Long
    Dim done As Long, errored As Long, paused As Long
    
    total = m_TaskTable.Count
    
    msg = "========================================" & vbCrLf
    msg = msg & "  Lua 协程任务列表" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf
    
    msg = msg & "任务总数: " & total & vbCrLf & vbCrLf
    
    ' ===== 统计状态 =====
    For Each taskId In m_TaskTable.Keys
        Set task = m_TaskTable(taskId)
        Select Case task("status")
            Case "defined": defined = defined + 1
            Case "yielded": yielded = yielded + 1
            Case "done": done = done + 1
            Case "error": errored = errored + 1
            Case "paused": paused = paused + 1
        End Select
    Next
    
    msg = msg & "状态统计:" & vbCrLf
    msg = msg & "  defined : " & defined & vbCrLf
    msg = msg & "  yielded : " & yielded & vbCrLf
    msg = msg & "  paused  : " & paused & vbCrLf
    msg = msg & "  done    : " & done & vbCrLf
    msg = msg & "  error   : " & errored & vbCrLf
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf & vbCrLf
    
    ' ===== 逐任务详情 =====
    Dim index As Long
    index = 1
    
    For Each taskId In m_TaskTable.Keys
        Set task = m_TaskTable(taskId)
        
        msg = msg & "【任务 #" & index & "】" & vbCrLf
        msg = msg & "ID      : " & taskId & vbCrLf
        msg = msg & "函数    : " & task("funcName") & vbCrLf
        msg = msg & "单元格  : " & task("cellAddr") & vbCrLf
        msg = msg & "状态    : " & task("status") & vbCrLf
        msg = msg & "进度    : " & Format(task("progress"), "0.0") & "%" & vbCrLf
        
        If Not IsEmpty(task("message")) Then
            msg = msg & "消息    : " & CStr(task("message")) & vbCrLf
        End If
        
        If task("status") = "error" And task("error") <> "" Then
            msg = msg & "错误    : " & task("error") & vbCrLf
        End If
        
        msg = msg & "活跃队列: " & IIf(m_ActiveQueue.Exists(taskId), "是", "否") & vbCrLf
        
        msg = msg & "----------------------------------------" & vbCrLf
        index = index + 1
    Next
    
    GetAllTasksInfo = msg
End Function

Public Function GetTaskDetail(taskId As String) As Object
    If Not m_TaskTable.Exists(taskId) Then
        Set GetTaskDetail = Nothing
        Exit Function
    End If

    Dim src As Object
    Set src = m_TaskTable(taskId)

    ' 返回一个只读副本（防止 LuaMenu 篡改内部状态）
    Dim info As Object
    Set info = CreateObject("Scripting.Dictionary")

    info("taskId") = taskId
    info("funcName") = src("funcName")
    info("cellAddr") = src("cellAddr")
    info("status") = src("status")
    info("progress") = src("progress")
    info("message") = src("message")
    info("value") = src("value")
    info("error") = src("error")
    info("coThread") = src("coThread")
    info("startArgs") = src("startArgs")
    info("resumeSpec") = src("resumeSpec")

    Set GetTaskDetail = info
End Function
    
Public Property Get GetWorkbook() As Workbook
    Set GetWorkbook = m_Workbook
End Property