' ============================================
' XLAM加载宏 - ThisWorkbook模块
' 负责加载宏的初始化、卸载和全局事件处理
' ============================================

Option Explicit

' ============================================
' 加载宏生命周期事件
' ============================================

' 加载宏打开时自动运行
Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    
    ' 初始化Lua引擎
    'If Not InitLuaState() Then
    '    MsgBox "ThisWorkbook.Workbook_Open: Lua引擎初始化失败！" & vbCrLf & _
    '           "部分功能可能不可用。", vbExclamation, "Workbook_Open加载宏启动警告"
    'End If
    
    ' 启用右键菜单
    DisableLuaTaskMenu
    EnableLuaTaskMenu
    
    ' 注册应用程序级事件
    ' RegisterApplicationEvents
    
    ' 显示欢迎信息（可选）
    ' MsgBox "Excel-Lua 5.4 加载宏已加载！", vbInformation, "欢迎"
    
    Exit Sub

ErrorHandler:
    MsgBox "ThisWorkbook.Workbook_Open: 加载宏启动失败: " & Err.Description, vbCritical, "严重错误"
End Sub

' 加载宏关闭前清理
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    ' 停止调度器
    StopScheduler
    
    ' 禁用右键菜单
    DisableLuaTaskMenu
    
    ' 注销应用程序级事件
    ' UnregisterApplicationEvents
    
    ' 清理Lua引擎（但保留其他工作簿的任务）
    ' 注意：这里不调用 CleanupLua，因为其他工作簿可能还在使用
    
    ' 显示退出信息（可选）
    MsgBox "Excel-Lua 5.4 加载宏已卸载。", vbInformation, "再见"
End Sub

' ============================================
' 应用程序级事件管理
' ============================================

' 工作簿关闭前事件 - 清理该工作簿的任务
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next

    ' 跳过加载宏自身
    If Wb.Name = ThisWorkbook.Name Then Exit Sub

    ' 清理该工作簿的所有任务
    CleanupWorkbookTasks Wb.Name
    ' 显示退出信息（可选）
    MsgBox "已清理该工作簿的所有任务", vbInformation, "再见"
End Sub

' 新建工作簿时的处理（可选）
Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    On Error Resume Next
    ' 可以在这里为新打开的工作簿做一些初始化
    ' 例如：自动检查是否需要functions.lua
End Sub
