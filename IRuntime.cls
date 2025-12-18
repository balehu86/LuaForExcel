' ============================================
' IRuntime.cls - 运行时接口定义
' ============================================
' 设计原因：
' 1. 调度器只依赖此接口，不知道实现细节
' 2. 允许未来扩展其他类型的运行时（如 Python）
' 3. 强制封装，防止调度器越界访问
' ============================================

Option Explicit

' 执行一次调度（处理一批任务）
Public Sub Tick()
End Sub

' 判断是否有可运行的任务
' 返回 False 时调度器会跳过此运行时
Public Function HasRunnable() As Boolean
End Function

' 清理资源（Workbook 关闭时调用）
Public Sub Dispose()
End Sub