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
    ' 初始化必要的全局变量
    g_NextTaskId = 1
    If g_Tasks Is Nothing Then Set g_Tasks = CreateObject("Scripting.Dictionary")
    If g_Workbooks Is Nothing Then Set g_Workbooks = CreateObject("Scripting.Dictionary")
    If g_TaskQueue Is Nothing Then Set g_TaskQueue = New Collection
    If g_Watches Is Nothing Then Set g_Watches = CreateObject("Scripting.Dictionary")
    ' 删除 g_DirtyWatches 初始化
    DisableLuaTaskMenu
    EnableLuaTaskMenu
    Exit Sub
ErrorHandler:
    MsgBox "ThisWorkbook.Workbook_Open: 加载宏启动失败: " & Err.Description, vbCritical, "严重错误"
End Sub

' 加载宏关闭前清理
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    StopScheduler
    DisableLuaTaskMenu
    ' 注销应用程序级事件
    Set App = Nothing
    If Not g_Workbooks Is Nothing Then g_Workbooks.RemoveAll
    Set g_Workbooks = Nothing
    If Not g_Tasks Is Nothing Then g_Tasks.RemoveAll
    Set g_Tasks = Nothing
    Set g_TaskQueue = Nothing
    If Not g_Watches Is Nothing Then g_Watches.RemoveAll
    Set g_Watches = Nothing
End Sub
' ============================================
' 应用程序级事件管理
' ============================================
' 打开普通工作簿
Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    On Error GoTo SafeExit
    If Wb Is ThisWorkbook Then Exit Sub

    ' 自动注册工作簿
    If Not g_Workbooks.Exists(Wb.Name) Then
        Dim wbInfo As New WorkbookInfo
        wbInfo.wbName = Wb.Name
        g_Workbooks.Add Wb.Name, wbInfo
        Debug.Print "App自动注册工作簿: " & Wb.Name
    End If
    Exit Sub
SafeExit:
    MsgBox "ThisWorkbook.App_WorkbookOpen: 打开工作簿出错: " & Err.Description, vbCritical, "错误"
End Sub

' 工作簿关闭前事件 - 清理该工作簿的任务
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    ' 跳过加载宏自身
    If Wb Is ThisWorkbook Then Exit Sub

    ' CleanupWorkbookTasks 内部已调用 CleanupWorkbookWatches
    CleanupWorkbookTasks Wb.Name

    If Not g_Workbooks Is Nothing Then
        If g_Workbooks.Exists(Wb.Name) Then
            g_Workbooks.Remove Wb.Name
        End If
    End If
End Sub