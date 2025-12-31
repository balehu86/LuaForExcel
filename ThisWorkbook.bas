' ============================================
' XLAM加载宏 - ThisWorkbook模块
' 负责加载宏的初始化、卸载和全局事件处理
' ============================================
Option Explicit
' Application 级事件对象
Private WithEvents App As Application
' ============================================
' 加载宏生命周期事件
' ============================================
' 加载宏打开时自动运行
Private Sub Workbook_Open()
    On Error GoTo ErrorHandler

    ' 绑定 Application 事件
    Set App = Application

    ' 初始化Lua引擎
    'If Not InitLuaState() Then
    '    MsgBox "ThisWorkbook.Workbook_Open: Lua引擎初始化失败！" & vbCrLf & _
    '           "部分功能可能不可用。", vbExclamation, "Workbook_Open加载宏启动警告"
    'End If

    ' 启用右键菜单
    DisableLuaTaskMenu
    EnableLuaTaskMenu
    ' 初始化全局工作簿字典
    If g_Workbooks Is Nothing Then Set g_Workbooks = CreateObject("Scripting.Dictionary")

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
    Set App = Nothing
    If Not g_Workbooks Is Nothing Then
        g_Workbooks.RemoveAll
    End If

    ' 清理Lua引擎（但保留其他工作簿的任务）
    ' 注意：这里不调用 CleanupLua，因为其他工作簿可能还在使用

    ' 显示退出信息（可选）
    ' MsgBox "Excel-Lua 5.4 加载宏已卸载。", vbInformation, "再见"
End Sub

' ============================================
' 应用程序级事件管理
' ============================================

' 打开普通工作簿
Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    On Error GoTo SafeExit
    ' 跳过加载宏自身
    If Wb Is ThisWorkbook Then Exit Sub

    Dim wbInfo As New WorkbookInfo
    wbInfo.Name = Wb.Name

    g_Workbooks.Add Wb.Name, wbInfo
    ' 可以在这里为新打开的工作簿做一些初始化
    ' 例如：自动检查是否需要functions.lua
SafeExit:
    MsgBox "ThisWorkbook.App_WorkbookOpen: 打开工作簿出错: " & Err.Description, vbCritical, "错误"
End Sub
' 工作簿关闭前事件 - 清理该工作簿的任务
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next

    ' 跳过加载宏自身
    If Wb Is ThisWorkbook Then Exit Sub

    ' 清理该工作簿的所有任务
    CleanupWorkbookTasks Wb.Name
    ' 从字典中移除
    If Not g_Workbooks Is Nothing Then
        If g_Workbooks.Exists(Wb.Name) Then
            g_Workbooks.Remove Wb.Name
        End If
    End If
    ' 显示退出信息（可选）
    ' MsgBox "已清理该工作簿的所有任务", vbInformation, "再见"
End Sub