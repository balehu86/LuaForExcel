' ============================================
' LuaMenu.bas - 右键菜单和用户宏
' ============================================
Option Explicit

' 辅助：增加子菜单项
Private Sub AddLuaMenuItem(parent As CommandBarControl, caption As String, onAction As String)
    Dim ctrl As CommandBarControl
    Set ctrl = parent.Controls.Add(Type:=msoControlButton, Temporary:=True)
    ctrl.Caption = caption
    ctrl.OnAction = onAction
End Sub

' 创建右键菜单
Public Sub CreateContextMenu()
    On Error Resume Next
    ' 删除已有菜单，避免重复
    RemoveContextMenu

    ' 获取右键菜单（Cell）
    Dim cellMenu As CommandBar
    Set cellMenu = Application.CommandBars("Cell")
    
    ' 添加单个任务的主菜单
    Dim luaTaskMenu As CommandBarControl
    Set luaTaskMenu = cellMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
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
    Set luaSchedulerMenu = cellMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaSchedulerMenu.Caption = "Lua 调度管理"
    luaSchedulerMenu.Tag = "LuaSchedulerMenu"
    ' 添加调度的子菜单
    AddLuaMenuItem luaSchedulerMenu, "启动所有 defined 任务", "LuaSchedulerMenu_StartAllDefinedTasks"
    AddLuaMenuItem luaSchedulerMenu, "清理所有完成、错误任务", "LuaSchedulerMenu_CleanupFinishedTasks"
    AddLuaMenuItem luaSchedulerMenu, "删除所有任务", "LuaSchedulerMenu_ClearAllTasks"
    AddLuaMenuItem luaSchedulerMenu, "显示所有任务信息", "LuaSchedulerMenu_ShowAllTasks"

    ' 添加管理的主菜单
    Dim luaConfigMenu As CommandBarControl
    Set luaConfigMenu = cellMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
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

' 移除右键菜单
Public Sub RemoveContextMenu()
    On Error Resume Next
    
    Dim cellMenu As CommandBar
    Set cellMenu = Application.CommandBars("Cell")
    
    Dim ctrl As CommandBarControl
    For Each ctrl In cellMenu.Controls
        If ctrl.Tag = "LuaTaskMenu" Then ctrl.Delete
        If ctrl.Tag = "LuaSchedulerMenu" Then ctrl.Delete
        If ctrl.Tag = "LuaConfigMenu" Then ctrl.Delete
    Next
End Sub

' 任务菜单回调：启动任务
Public Sub LuaTaskMenu_StartTask()
    On Error GoTo ErrorHandler
    
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    
    If taskId = "" Then
        MsgBox "所选单元格不是有效的任务。" & vbCrLf & _
               "请选择包含 =LuaTask() 的单元格。", vbExclamation, "无效选择"
        Exit Sub
    End If
    
    If LuaTasks.StartTask(taskId) Then
        MsgBox "任务已启动: " & taskId, vbInformation, "成功"
    Else
        MsgBox "任务启动失败。", vbExclamation, "失败"
    End If
    Exit Sub
ErrorHandler:
    MsgBox "启动任务失败: " & Err.Description, vbCritical, "错误"
End Sub

' 任务菜单回调：终止任务
Public Sub LuaTaskMenu_TerminateTask()
    On Error GoTo ErrorHandler
    
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    
    If taskId = "" Then
        MsgBox "所选单元格不是有效的任务。", vbExclamation, "无效选择"
        Exit Sub
    End If
    
    If Not LuaCore.g_TaskIndex.Exists(taskId) Then
        MsgBox "任务不存在。", vbExclamation, "错误"
        Exit Sub
    End If
    
    Dim wbKey As String
    wbKey = LuaCore.g_TaskIndex(taskId)
    
    If Not LuaCore.g_Runtimes.Exists(wbKey) Then
        MsgBox "Runtime不存在。", vbCritical, "错误"
        Exit Sub
    End If
    
    Dim rt As LuaCore.WorkbookRuntime
    rt = LuaCore.g_Runtimes(wbKey)
    
    rt.TaskStatus(taskId) = "terminated"
    If rt.ActiveTasks.Exists(taskId) Then
        rt.ActiveTasks.Remove taskId
    End If
    rt.StateDirty = True
    
    LuaCore.g_Runtimes(wbKey) = rt
    
    MsgBox "任务已终止: " & taskId, vbInformation, "成功"
    Exit Sub

ErrorHandler:
    MsgBox "终止任务失败: " & Err.Description, vbCritical, "错误"
End Sub

' 任务菜单回调：暂停任务
Public Sub LuaTaskMenu_PauseTask()
    On Error GoTo ErrorHandler
    
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    
    If taskId = "" Then
        MsgBox "所选单元格不是有效的任务。", vbExclamation, "无效选择"
        Exit Sub
    End If
    
    If Not LuaCore.g_TaskIndex.Exists(taskId) Then
        MsgBox "任务不存在。", vbExclamation, "错误"
        Exit Sub
    End If
    
    Dim wbKey As String
    wbKey = LuaCore.g_TaskIndex(taskId)
    
    Dim rt As LuaCore.WorkbookRuntime
    rt = LuaCore.g_Runtimes(wbKey)
    
    If rt.TaskStatus(taskId) <> "yielded" Then
        MsgBox "只能暂停 yielded 状态的任务。", vbExclamation, "无效操作"
        Exit Sub
    End If
    
    rt.TaskStatus(taskId) = "paused"
    If rt.ActiveTasks.Exists(taskId) Then
        rt.ActiveTasks.Remove taskId
    End If
    rt.StateDirty = True
    
    LuaCore.g_Runtimes(wbKey) = rt
    
    MsgBox "任务已暂停: " & taskId, vbInformation, "成功"
    Exit Sub

ErrorHandler:
    MsgBox "暂停任务失败: " & Err.Description, vbCritical, "错误"
End Sub

' 任务菜单回调：恢复任务
Public Sub LuaTaskMenu_ResumeTask()
    On Error GoTo ErrorHandler
    
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    
    If taskId = "" Then
        MsgBox "所选单元格不是有效的任务。", vbExclamation, "无效选择"
        Exit Sub
    End If
    
    If Not LuaCore.g_TaskIndex.Exists(taskId) Then
        MsgBox "任务不存在。", vbExclamation, "错误"
        Exit Sub
    End If
    
    Dim wbKey As String
    wbKey = LuaCore.g_TaskIndex(taskId)
    
    Dim rt As LuaCore.WorkbookRuntime
    rt = LuaCore.g_Runtimes(wbKey)
    
    If rt.TaskStatus(taskId) <> "paused" Then
        MsgBox "只能恢复 paused 状态的任务。", vbExclamation, "无效操作"
        Exit Sub
    End If
    
    rt.TaskStatus(taskId) = "yielded"
    rt.ActiveTasks(taskId) = True
    rt.StateDirty = True
    
    LuaCore.g_Runtimes(wbKey) = rt
    
    MsgBox "任务已恢复: " & taskId, vbInformation, "成功"
    Exit Sub

ErrorHandler:
    MsgBox "恢复任务失败: " & Err.Description, vbCritical, "错误"
End Sub

' 设置菜单回调：热更新
Public Sub LuaConfigMenu_ReloadFunctions()
    On Error GoTo ErrorHandler
    
    If ActiveWorkbook Is Nothing Then
        MsgBox "没有活动工作簿。", vbExclamation, "错误"
        Exit Sub
    End If
    
    Dim wbKey As String
    wbKey = ActiveWorkbook.FullName
    
    If Not LuaCore.g_Runtimes.Exists(wbKey) Then
        MsgBox "当前工作簿未注册。", vbExclamation, "错误"
        Exit Sub
    End If
    
    Dim rt As LuaCore.WorkbookRuntime
    rt = LuaCore.g_Runtimes(wbKey)
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(rt.FunctionsPath) Then
        MsgBox "找不到 functions.lua 文件: " & vbCrLf & rt.FunctionsPath, vbExclamation, "文件不存在"
        Exit Sub
    End If
    
    Dim result As Long
    result = luaL_loadfilex(rt.LuaState, rt.FunctionsPath, 0)
    If result = 0 Then result = lua_pcallk(rt.LuaState, 0, 0, 0, 0, 0)
    
    lua_settop rt.LuaState, 0
    
    If result = 0 Then
        rt.LastModified = FileDateTime(rt.FunctionsPath)
        LuaCore.g_Runtimes(wbKey) = rt
        MsgBox "functions.lua 已重新加载。", vbInformation, "成功"
    Else
        MsgBox "热更新失败，请检查 Lua 语法。", vbExclamation, "失败"
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "热更新失败: " & Err.Description, vbCritical, "错误"
End Sub

' 设置菜单回调：显示状态
Public Sub LuaConfigMenu_ShowStatus()
    Dim msg As String
    msg = "Lua 调度器状态" & vbCrLf & vbCrLf
    msg = msg & "运行中: " & IIf(LuaCore.g_SchedulerRunning, "是", "否") & vbCrLf
    msg = msg & "已注册工作簿: " & LuaCore.g_Runtimes.Count & vbCrLf
    msg = msg & "任务总数: " & LuaCore.g_TaskIndex.Count & vbCrLf
    msg = msg & "调度间隔: " & LuaCore.g_IntervalMs & " 毫秒"
    
    MsgBox msg, vbInformation, "调度器状态"
End Sub

' 用户宏：启动调度器
Public Sub StartScheduler()
    LuaCore.StartScheduler
    MsgBox "调度器已启动。", vbInformation, "调度器"
End Sub

' 用户宏：停止调度器
Public Sub StopScheduler()
    LuaCore.StopScheduler
    MsgBox "调度器已停止。", vbInformation, "调度器"
End Sub

' 辅助函数：从选中单元格获取 taskId
Private Function GetTaskIdFromSelection() As String
    On Error Resume Next
    
    If Selection Is Nothing Then
        GetTaskIdFromSelection = ""
        Exit Function
    End If
    
    If TypeName(Selection) <> "Range" Then
        GetTaskIdFromSelection = ""
        Exit Function
    End If
    
    Dim cell As Range
    Set cell = Selection.Cells(1, 1)
    
    If Not cell.HasFormula Then
        GetTaskIdFromSelection = ""
        Exit Function
    End If
    
    Dim formula As String
    formula = cell.formula
    
    If Not (formula Like "=LuaTask(*") Then
        GetTaskIdFromSelection = ""
        Exit Function
    End If
    
    Dim taskId As Variant
    taskId = cell.value
    
    If VarType(taskId) = vbString Then
        GetTaskIdFromSelection = CStr(taskId)
    Else
        GetTaskIdFromSelection = ""
    End If
End Function