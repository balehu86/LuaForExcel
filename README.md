# 在excel里调用lua函数和库
下载`luaForExcel.xlam`到任意地方，在同一目录下，创建`functions.lua`文件，无须`require`额外包，就可编写你需要调用的`lua`函数

## 特性
支持普通函数、协程函数。
* 普通函数通过`=LuaCall("your_lua_function",args...)`调用。
* 协程函数通过`=LuaTask("your_lua_coroutine",start_args...[,"|",resume_args...])`调用。
* 协程结果通过`=LuaGet(taskId,"value")`等标记获取任务参数、结果、调试信息等。

支持`functions.lua`文件自动热重载。  
支持