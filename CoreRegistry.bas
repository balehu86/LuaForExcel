' ============================================
' CoreRegistry.bas - 运行时注册表
' ============================================
' 设计原因：
' 1. 封装所有 Dictionary 访问
' 2. 提供 taskId → IRuntime 路由
' 3. 管理 Workbook → Runtime 映射
' 4. 提供 WorkbookKey 工具
' ============================================

Option Explicit

' ===== 全局注册表（仅本模块访问）=====
Private g_RuntimeByWorkbook As Object  ' wbKey → WorkbookRuntime
Private g_TaskIndex As Object          ' taskId → IRuntime

' ============================================
' 初始化
' ============================================

Private Sub InitRegistry()
    If g_RuntimeByWorkbook Is Nothing Then
        Set g_RuntimeByWorkbook = CreateObject("Scripting.Dictionary")
        Set g_TaskIndex = CreateObject("Scripting.Dictionary")
    End If
End Sub

' ============================================
' 运行时管理
' ============================================

' 注册工作簿运行时
Public Sub RegisterWorkbookRuntime(wb As Workbook, rt As IRuntime)
    InitRegistry
    Dim wbKey As String
    wbKey = GetWorkbookKey(wb)
    g_RuntimeByWorkbook(wbKey) = rt
    
    ' 注册到调度器
    Scheduler.RegisterRunnable rt
End Sub

' 注销工作簿运行时
Public Sub UnregisterWorkbookRuntime(wb As Workbook)
    InitRegistry
    Dim wbKey As String
    wbKey = GetWorkbookKey(wb)
    
    If g_RuntimeByWorkbook.Exists(wbKey) Then
        Dim rt As IRuntime
        Set rt = g_RuntimeByWorkbook(wbKey)
        
        ' 从调度器注销
        Scheduler.UnregisterRunnable rt
        
        ' 清理资源
        rt.Dispose
        
        ' 从注册表移除
        g_RuntimeByWorkbook.Remove wbKey
    End If
End Sub

' 根据 Workbook 获取运行时
Public Function GetRuntimeByWorkbook(wb As Workbook) As IRuntime
    InitRegistry
    Dim wbKey As String
    wbKey = GetWorkbookKey(wb)
    
    If g_RuntimeByWorkbook.Exists(wbKey) Then
        Set GetRuntimeByWorkbook = g_RuntimeByWorkbook(wbKey)
    Else
        Set GetRuntimeByWorkbook = Nothing
    End If
End Function

' ============================================
' Task 路由表（GlobalTaskIndex）
' ============================================

' 注册任务到运行时的映射
Public Sub RegisterTask(taskId As String, rt As IRuntime)
    InitRegistry
    g_TaskIndex(taskId) = rt
End Sub

' 根据 taskId 解析运行时
Public Function ResolveRuntime(taskId As String) As IRuntime
    InitRegistry
    If g_TaskIndex.Exists(taskId) Then
        Set ResolveRuntime = g_TaskIndex(taskId)
    Else
        Set ResolveRuntime = Nothing
    End If
End Function

' 注销任务
Public Sub UnregisterTask(taskId As String)
    InitRegistry
    If g_TaskIndex.Exists(taskId) Then
        g_TaskIndex.Remove taskId
    End If
End Sub

' ============================================
' 工具函数（WorkbookKey 生成）
' ============================================

' 生成唯一的 Workbook 标识
' 设计原因：FullName 可能重复（不同路径同名文件）
' 使用 Name + CreationDate 组合
Public Function GetWorkbookKey(wb As Workbook) As String
    On Error Resume Next
    GetWorkbookKey = wb.Name & "_" & Format(wb.BuiltinDocumentProperties("Creation Date"), "yyyymmddhhnnss")
    If Err.Number <> 0 Then
        ' 回退方案：使用 Name + 当前时间
        GetWorkbookKey = wb.Name & "_" & Format(Now, "yyyymmddhhnnss")
    End If
End Function

' 从外部地址提取 WorkbookKey
' 例如：'[Book1.xlsx]Sheet1'!$A$1 → Book1.xlsx
Public Function ExtractWorkbookFromAddress(addr As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(addr, "[")
    endPos = InStr(addr, "]")
    
    If startPos > 0 And endPos > startPos Then
        ExtractWorkbookFromAddress = Mid(addr, startPos + 1, endPos - startPos - 1)
    Else
        ExtractWorkbookFromAddress = ""
    End If
End Function