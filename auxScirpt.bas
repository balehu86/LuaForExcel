' ===== 日志工具 =====
Public Sub LogInfo(msg As String)
    Debug.Print "[INFO] " & msg
End Sub
Public Sub LogError(msg As String)
    Debug.Print "[ERROR] " & msg
    MsgBox msg, vbCritical, "错误"
End Sub
' 工作簿关闭时自动清理
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    CleanupLua
    DisableLuaTaskMenu
End Sub

' 工作簿打开时自动运行
Private Sub Workbook_Open()
    EnableLuaTaskMenu
End Sub

' 辅助函数：获取数组维度
Private Function ArrayDimensions(arr As Variant) As String
    On Error Resume Next
    Dim dims As String
    Dim i As Long
    
    For i = 1 To 10  ' 最多检查10维
        Dim ub As Long
        ub = UBound(arr, i)
        If Err.Number <> 0 Then
            Exit For
        End If
        
        Dim lb As Long
        lb = LBound(arr, i)
        
        If dims <> "" Then dims = dims & " x "
        dims = dims & (ub - lb + 1)
    Next i
    
    If dims = "" Then dims = "未知"
    ArrayDimensions = dims
End Function

' 辅助：增加子菜单项
Private Sub AddLuaMenuItem(parent As CommandBarControl, caption As String, onAction As String)
    Dim ctrl As CommandBarControl
    Set ctrl = parent.Controls.Add(Type:=msoControlButton, Temporary:=True)
    ctrl.Caption = caption
    ctrl.OnAction = onAction
End Sub

' ============================================
' 第七部分：可视化操作函数
' ============================================

' 根据单元格地址获取任务ID
Private Function GetTaskIdFromSelection() As String
    Dim cellAddr As String
    cellAddr = Selection.Address(External:=True)
    GetTaskIdFromSelection = FindTaskByCell(cellAddr)
End Function

' 打开并创建右键菜单功能
Public Sub EnableLuaTaskMenu()
    On Error Resume Next

    ' 删除已有菜单，避免重复
    Call DisableLuaTaskMenu

    ' 获取右键菜单（Cell）
    Dim cMenu As CommandBar
    Set cMenu = Application.CommandBars("Cell")

    ' 添加单个任务的主菜单
    Dim luaTaskMenu As CommandBarControl
    Set luaTaskMenu = cMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaTaskMenu.Caption = "Lua 任务管理"
    luaTaskMenu.Tag = "LuaTaskMenu"
    ' 添加单个任务的子菜单
    AddLuaMenuItem luaTaskMenu, "启动任务", "LuaTaskMenu_StartTask"
    AddLuaMenuItem luaTaskMenu, "暂停任务", "LuaTaskMenu_PauseTask"
    AddLuaMenuItem luaTaskMenu, "恢复任务", "LuaTaskMenu_ResumeTask"
    AddLuaMenuItem luaTaskMenu, "终止任务", "LuaTaskMenu_TerminateTask"
    AddLuaMenuItem luaTaskMenu, "查看任务详情", "LuaTaskMenu_ShowDetail"

    ' 添加调度的主菜单
    Dim luaSchedulerMenu As CommandBarControl
    Set luaSchedulerMenu = cMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaSchedulerMenu.Caption = "Lua 调度管理"
    luaSchedulerMenu.Tag = "LuaSchedulerMenu"
    ' 添加调度的子菜单
    AddLuaMenuItem luaSchedulerMenu, "启动所有 defined 任务", "LuaSchedulerMenu_StartAllDefinedTasks"
    AddLuaMenuItem luaSchedulerMenu, "清理所有完成、错误任务", "LuaSchedulerMenu_CleanupFinishedTasks"
    AddLuaMenuItem luaSchedulerMenu, "删除所有任务", "LuaSchedulerMenu_ClearAllTasks"
    AddLuaMenuItem luaSchedulerMenu, "显示所有任务信息", "LuaSchedulerMenu_ShowAllTasks"

    ' 添加管理的主菜单
    Dim luaConfigMenu As CommandBarControl
    Set luaConfigMenu = cMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaConfigMenu.Caption = "Lua 设置管理"
    luaConfigMenu.Tag = "luaConfigMenu"
    ' 添加管理的子菜单
    AddLuaMenuItem luaConfigMenu, "启用热重载", "LuaConfigMenu_EnableHotReload"
    AddLuaMenuItem luaConfigMenu, "禁用热重载", "LuaConfigMenu_DisableHotReload"
    AddLuaMenuItem luaConfigMenu, "手动重载 functions.lua", "LuaConfigMenu_ReloadFunctions"
    AddLuaMenuItem luaConfigMenu, "设置调度间隔（秒）", "LuaConfigMenu_SetSchedulerInterval"
    AddLuaMenuItem luaConfigMenu, "设置调度步数", "LuaConfigMenu_SetSchedulerBatchSize"

    MsgBox "Lua 任务右键菜单已启用。", vbInformation
End Sub

' 关闭右键菜单
Public Sub DisableLuaTaskMenu()
    On Error Resume Next
    Dim cMenu As CommandBar
    Set cMenu = Application.CommandBars("Cell")

    Dim ctrl As CommandBarControl
    For Each ctrl In cMenu.Controls
        If ctrl.Tag = "LuaTaskMenu" Then ctrl.Delete
        If ctrl.Tag = "LuaSchedulerMenu" Then ctrl.Delete
        If ctrl.Tag = "LuaConfigMenu" Then ctrl.Delete
    Next
End Sub

' ===== 启动任务 =====
Private Sub LuaTaskMenu_StartTask()
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If
    If g_TaskStatus(taskId) = "defined" Then
        StartLuaCoroutine taskId
        MsgBox "任务已启动: " & taskId, vbInformation
    Else
        MsgBox "任务状态为 " & g_TaskStatus(taskId) & "，无法启动。", vbExclamation
    End If
End Sub

' ===== 暂停任务 =====
Private Sub LuaTaskMenu_PauseTask()
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If
    If Not g_TaskFunc.Exists(taskId) Then
        MsgBox "任务已不存在。", vbExclamation
        Exit Sub
    End If

    If g_TaskQueue.Exists(taskId) Then
        g_TaskQueue.Remove taskId
        g_TaskStatus(taskId) = "paused"
        MsgBox "任务 " & taskId & " 已暂停。" & vbCrLf & _
               "使用 ResumeTask 恢复。", vbInformation, "任务已暂停"
    Else
        MsgBox "任务 " & taskId & " 不在活跃队列中。", vbExclamation, "提示"
    End If
End Sub

' ===== 恢复任务 =====
Private Sub LuaTaskMenu_ResumeTask()
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If
    If Not g_TaskFunc.Exists(taskId) Then
        MsgBox "任务 " & taskId & " 不存在。", vbCritical, "错误"
        Exit Sub
    End If

    Dim status As String
    status = g_TaskStatus(taskId)
    If status <> "yielded" And status <> "paused" Then
        MsgBox "任务 " & taskId & " 状态为 " & status & "，无法恢复。", vbExclamation, "无法恢复"
        Exit Sub
    End If

    If Not g_TaskQueue.Exists(taskId) Then
        g_TaskQueue(taskId) = True
        StartSchedulerIfNeeded
        MsgBox "任务 " & taskId & " 已恢复。", vbInformation, "任务已恢复"
    Else
        MsgBox "任务 " & taskId & " 已在活跃队列中。", vbInformation, "提示"
    End If
End Sub

' ===== 终止任务 =====
Private Sub LuaTaskMenu_terminateTask()
    On Error Resume Next
    If g_TaskFunc Is Nothing Then Exit Sub

    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Or Not g_TaskFunc.Exists(taskId) Then
        MsgBox "任务不存在或已删除", vbExclamation
        Exit Sub
    End If

    If g_TaskQueue Is Nothing Then
        MsgBox "任务队列未创建", vbExclamation
        If g_TaskQueue.Exists(taskId) Then g_TaskQueue.Remove taskId
    Else
        If g_TaskQueue.Exists(taskId) Then g_TaskQueue.Remove taskId
    End If

    g_TaskStatus(taskId) = "terminated"

    If g_TaskFunc.Exists(taskId) Then g_TaskFunc.Remove taskId
    If g_TaskStartArgs.Exists(taskId) Then g_TaskStartArgs.Remove taskId
    If g_TaskResumeSpec.Exists(taskId) Then g_TaskResumeSpec.Remove taskId
    ' 删除 dynamicTargets 和 writeTargets
    If g_TaskCell.Exists(taskId) Then g_TaskCell.Remove taskId
    If g_TaskStatus.Exists(taskId) Then g_TaskStatus.Remove taskId
    If g_TaskProgress.Exists(taskId) Then g_TaskProgress.Remove taskId
    If g_TaskMessage.Exists(taskId) Then g_TaskMessage.Remove taskId
    If g_TaskValue.Exists(taskId) Then g_TaskValue.Remove taskId
    If g_TaskError.Exists(taskId) Then g_TaskError.Remove taskId
    If g_TaskCoThread.Exists(taskId) Then g_TaskCoThread.Remove taskId

    MsgBox "任务已终止并删除: " & taskId, vbInformation
End Sub

' ===== 查看任务详情 =====
Private Sub LuaTaskMenu_ShowDetail()
    On Error GoTo ErrorHandler

    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If

    If g_TaskFunc Is Nothing Then
        InitCoroutineSystem
    End If
    
    If Not g_TaskFunc.Exists(taskId) Then
        MsgBox "任务 " & taskId & " 不存在！", vbCritical, "错误"
        Exit Sub
    End If
    
    Dim msg As String
    msg = "========================================" & vbCrLf
    msg = msg & "  任务详细信息" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf
    
    msg = msg & "任务ID: " & taskId & vbCrLf
    msg = msg & "函数名: " & g_TaskFunc(taskId) & vbCrLf
    msg = msg & "单元格: " & g_TaskCell(taskId) & vbCrLf
    msg = msg & "状态: " & g_TaskStatus(taskId) & vbCrLf
    msg = msg & "进度: " & Format(g_TaskProgress(taskId), "0.00") & "%" & vbCrLf
    msg = msg & "消息: " & CStr(g_TaskMessage(taskId)) & vbCrLf
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf & vbCrLf
    
    ' 启动参数
    msg = msg & "启动参数:" & vbCrLf
    Dim startArgs As Variant
    startArgs = g_TaskStartArgs(taskId)
    If IsArray(startArgs) Then
        Dim i As Long
        For i = LBound(startArgs) To UBound(startArgs)
            msg = msg & "   [" & i & "] " & CStr(startArgs(i)) & vbCrLf
        Next i
    Else
        msg = msg & "   (无)" & vbCrLf
    End If

    ' Resume 参数
    msg = msg & vbCrLf & "Resume 参数:" & vbCrLf
    Dim resumeSpec As Variant
    resumeSpec = g_TaskResumeSpec(taskId)
    If IsArray(resumeSpec) Then
        For i = LBound(resumeSpec) To UBound(resumeSpec)
            msg = msg & "   [" & i & "] " & CStr(resumeSpec(i)) & vbCrLf
        Next i
    Else
        msg = msg & "   (无)" & vbCrLf
    End If

    ' 当前值
    msg = msg & vbCrLf & "当前值:" & vbCrLf
    Dim value As Variant
    value = g_TaskValue(taskId)
    If IsArray(value) Then
        msg = msg & "   (数组，维度: " & ArrayDimensions(value) & ")" & vbCrLf
    ElseIf IsEmpty(value) Then
        msg = msg & "   (空)" & vbCrLf
    Else
        Dim valueStr As String
        valueStr = CStr(value)
        If Len(valueStr) > 100 Then valueStr = Left(valueStr, 97) & "..."
        msg = msg & "   " & valueStr & vbCrLf
    End If
    
    ' 错误信息
    If g_TaskStatus(taskId) = "error" Then
        msg = msg & vbCrLf & " 错误信息:" & vbCrLf
        msg = msg & "   " & g_TaskError(taskId) & vbCrLf
    End If
    
    ' 调度信息
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf
    msg = msg & "在活跃队列中: " & IIf(g_TaskQueue.Exists(taskId), "是", "否") & vbCrLf
    msg = msg & "协程线程: " & IIf(g_TaskCoThread(taskId) = 0, "未创建", "0x" & Hex(g_TaskCoThread(taskId))) & vbCrLf
    
    MsgBox msg, vbInformation, "任务详情 - " & taskId
    
    Exit Sub

ErrorHandler:
    MsgBox "显示任务详情时出错: " & Err.Description, vbCritical, "错误"
End Sub

' ===== 批量启动所有 defined 状态的任务 =====
Private Sub LuaSchedulerMenu_StartAllDefinedTasks()
    Dim taskId As Variant
    Dim count As Long
    
    For Each taskId In g_TaskFunc.Keys
        If g_TaskStatus(CStr(taskId)) = "defined" Then
            StartLuaCoroutine CStr(taskId)
            count = count + 1
        End If
    Next
    
    MsgBox "已启动 " & count & " 个任务", vbInformation
End Sub

' ===== 清理所有已完成或错误的任务 =====
Private Sub LuaSchedulerMenu_CleanupFinishedTasks()
    On Error Resume Next
    
    If g_TaskFunc Is Nothing Then
        MsgBox "没有任务需要清理。", vbInformation, "清理任务"
        Exit Sub
    End If
    
    Dim taskId As Variant
    Dim tasksToRemove As Object
    Set tasksToRemove = CreateObject("System.Collections.ArrayList")
    
    ' 收集需要清理的任务
    For Each taskId In g_TaskFunc.Keys
        Dim status As String
        status = g_TaskStatus(CStr(taskId))
        If status = "done" Or status = "error" Then
            tasksToRemove.Add taskId
        End If
    Next taskId
    
    ' 清理任务
    Dim i As Long
    For i = 0 To tasksToRemove.Count - 1
        Dim tid As String
        tid = CStr(tasksToRemove(i))
        
        g_TaskFunc.Remove tid
        g_TaskStartArgs.Remove tid
        g_TaskResumeSpec.Remove tid
        g_TaskCell.Remove tid
        g_TaskStatus.Remove tid
        g_TaskProgress.Remove tid
        g_TaskMessage.Remove tid
        g_TaskValue.Remove tid
        g_TaskError.Remove tid
        g_TaskCoThread.Remove tid
        
        If g_TaskQueue.Exists(tid) Then
            g_TaskQueue.Remove tid
        End If
    Next i
    
    MsgBox "已清理 " & tasksToRemove.Count & " 个已完成或错误的任务。" & vbCrLf & _
           "剩余任务: " & g_TaskFunc.Count, vbInformation, "清理完成"
End Sub

' ===== 清空所有任务和队列 =====
Private Sub LuaSchedulerMenu_ClearAllTasks()
    Dim result As VbMsgBoxResult
    result = MsgBox("确定要清空所有任务吗？" & vbCrLf & vbCrLf & _
                    "这将删除所有任务数据，无法恢复！", _
                    vbExclamation + vbYesNo, "确认清空")
    
    If result = vbNo Then Exit Sub
    
    ' 停止调度器
    g_SchedulerRunning = False
    
    ' 清空所有 Dictionary
    If Not g_TaskFunc Is Nothing Then
        g_TaskFunc.RemoveAll
        g_TaskStartArgs.RemoveAll
        g_TaskResumeSpec.RemoveAll
        g_TaskCell.RemoveAll
        g_TaskStatus.RemoveAll
        g_TaskProgress.RemoveAll
        g_TaskMessage.RemoveAll
        g_TaskValue.RemoveAll
        g_TaskError.RemoveAll
        g_TaskCoThread.RemoveAll
        g_TaskQueue.RemoveAll
    End If
    
    MsgBox "所有任务已清空。", vbInformation, "清空完成"
End Sub

' ===== 显示所有任务信息 =====
Private Sub LuaSchedulerMenu_ShowAllTasks()
    On Error GoTo ErrorHandler
    
    If g_TaskFunc Is Nothing Then
        InitCoroutineSystem
    End If
    
    If g_TaskFunc.Count = 0 Then
        MsgBox "当前没有任何任务。", vbInformation, "任务列表"
        Exit Sub
    End If
    
    ' 构建任务信息字符串
    Dim msg As String
    Dim taskId As Variant
    Dim taskCount As Long
    Dim runningCount As Long, yieldedCount As Long, doneCount As Long, errorCount As Long
    
    msg = "========================================" & vbCrLf
    msg = msg & "  Lua 协程任务管理器" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf
    
    msg = msg & "任务总数: " & g_TaskFunc.Count & vbCrLf
    msg = msg & "活跃队列: " & g_TaskQueue.Count & vbCrLf
    msg = msg & "调度器: " & IIf(g_SchedulerRunning, "运行中", "已停止") & vbCrLf
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf & vbCrLf
    
    ' 统计各状态任务数
    For Each taskId In g_TaskFunc.Keys
        Select Case g_TaskStatus(CStr(taskId))
            Case "running": runningCount = runningCount + 1
            Case "yielded": yieldedCount = yieldedCount + 1
            Case "done": doneCount = doneCount + 1
            Case "error": errorCount = errorCount + 1
        End Select
    Next taskId
    
    msg = msg & "状态统计:" & vbCrLf
    msg = msg & "   运行中: " & runningCount & vbCrLf
    msg = msg & "   暂停中: " & yieldedCount & vbCrLf
    msg = msg & "   已完成: " & doneCount & vbCrLf
    msg = msg & "   错误: " & errorCount & vbCrLf
    msg = msg & vbCrLf & "========================================" & vbCrLf & vbCrLf
    
    ' 详细列出每个任务
    taskCount = 0
    For Each taskId In g_TaskFunc.Keys
        taskCount = taskCount + 1
        msg = msg & "【任务 #" & taskCount & "】" & vbCrLf
        msg = msg & "  ID: " & CStr(taskId) & vbCrLf
        msg = msg & "  函数: " & g_TaskFunc(CStr(taskId)) & vbCrLf
        msg = msg & "  单元格: " & g_TaskCell(CStr(taskId)) & vbCrLf
        msg = msg & "  状态: " & g_TaskStatus(CStr(taskId)) & vbCrLf
        msg = msg & "  进度: " & Format(g_TaskProgress(CStr(taskId)), "0.0") & "%" & vbCrLf
        
        ' 显示消息
        Dim msgText As String
        msgText = CStr(g_TaskMessage(CStr(taskId)))
        If Len(msgText) > 50 Then msgText = Left(msgText, 47) & "..."
        msg = msg & "  消息: " & msgText & vbCrLf
        
        ' 如果有错误，显示错误信息
        If g_TaskStatus(CStr(taskId)) = "error" Then
            Dim errText As String
            errText = CStr(g_TaskError(CStr(taskId)))
            If Len(errText) > 60 Then errText = Left(errText, 57) & "..."
            msg = msg & "   错误: " & errText & vbCrLf
        End If
        
        ' 显示是否在活跃队列中
        If g_TaskQueue.Exists(CStr(taskId)) Then
            msg = msg & "  队列: 是" & vbCrLf
        End If
        
        msg = msg & "----------------------------------------" & vbCrLf
    Next taskId
    
    ' 显示消息框
    MsgBox msg, vbInformation, "Lua 协程任务列表 (" & g_TaskFunc.Count & " 个任务)"
    
    Exit Sub
ErrorHandler:
    MsgBox "显示任务信息时出错: " & Err.Description, vbCritical, "错误"
End Sub

' ===== 启用热重载 =====
Private Sub LuaConfigMenu_EnableHotReload()
    g_HotReloadEnabled = True
    MsgBox "Lua 自动热重载已启用。" & vbCrLf & _
           "当 functions.lua 修改后，系统将自动重新加载。", _
           vbInformation, "热重载已启用"
End Sub

' ===== 禁用热重载 =====
Private Sub LuaConfigMenu_DisableHotReload()
    g_HotReloadEnabled = False
    MsgBox "Lua 自动热重载已禁用。" & vbCrLf & _
           "如需更新 functions.lua，请手动运行 ""ReloadFunctions""。", _
           vbExclamation, "热重载已禁用"
End Sub

' ===== 手动重载 functions.lua =====
Private Sub LuaConfigMenu_ReloadFunctions()
    If Not g_Initialized Then
        If Not InitLuaState() Then
            MsgBox "无法初始化 Lua 引擎。", vbCritical, "重载失败"
            Exit Sub
        End If
    End If
    
    If TryLoadFunctionsFile() Then
        MsgBox "functions.lua 已成功重载！", vbInformation, "重载成功"
    Else
        MsgBox "functions.lua 重载失败。" & vbCrLf & _
               "请检查文件语法。", vbCritical, "重载失败"
    End If
End Sub

' ===== 设置调度间隔（毫秒） =====
Private Sub LuaConfigMenu_SetSchedulerInterval()
    If g_SchedulerRunning Then
        StopScheduler
    End If

    Dim seconds As Long
    seconds = Application.InputBox( _
            "请输入调度的间隔时间（>=10ms且<=60000ms）", _
            "调度参数", _
            g_SchedulerIntervalMilliSec, _
            Type:=1 _
        )

    If seconds = False Then Exit Sub
    If seconds < 10 Or seconds > 60 Then
        MsgBox "间隔不能小于 10 ms。且不能大于 60 秒。", vbExclamation, "无效值"
        Exit Sub
    End If

    g_SchedulerIntervalMilliSec = seconds
    ResumeScheduler
End Sub

' ===== 设置调度步数 ===== 
Private Sub LuaConfigMenu_SetSchedulerBatchSize()
    If g_SchedulerRunning Then
        StopScheduler
    End If

    Dim v As Variant
    v = Application.InputBox( _
            "请输入每次调度的最大任务数（>=1）", _
            "调度参数", _
            g_MaxIterationsPerTick, _
            Type:=1 _
        )

    If v = False Then Exit Sub
    If v < 1 Or v <> CLng(v) Then
        MsgBox "请输入 >=1 的整数", vbExclamation
        Exit Sub
    End If

    g_MaxIterationsPerTick = CLng(v)
    ResumeScheduler
End Sub

' 手动设置 functions.lua 路径（供高级用户使用）
Private Sub LuaConfigMenu_SetFunctionsPath(filePath As String)
    On Error GoTo ErrorHandler
    
    If Not g_Initialized Then
        If Not InitLuaState() Then
            MsgBox "Lua 初始化失败", vbCritical
            Exit Sub
        End If
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(filePath) Then
        MsgBox "文件不存在: " & filePath, vbExclamation
        Exit Sub
    End If
    
    g_LuaState.functionsPath = filePath
    g_LuaState.lastModified = #1/1/1900#
    
    If TryLoadFunctionsFile() Then
        MsgBox "functions.lua 已加载: " & filePath, vbInformation
    Else
        MsgBox "加载失败", vbCritical
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "设置路径失败: " & Err.Description, vbCritical
End Sub