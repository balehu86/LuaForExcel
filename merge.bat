@echo off
copy /b "WorkbookRuntime.cls"+"Scheduler.bas"+"CoreRegistry.bas"+"CoreUI.bas"+"UDF.bas"+"ThisWorkbook.cls" combined.txt
echo 合并完成！
pause