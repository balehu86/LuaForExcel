-- ============================================
-- Lua 协程函数编写模板
-- ============================================

-- 模板 1: 基础协程函数（带进度报告）
-- 参数：
--   taskCell: 任务单元格地址（自动传入）
--   ...：启动参数（在 LuaTask 中定义）
function my_coroutine_task(taskCell, arg1, arg2, ...)
    -- 初始化
    local progress = 0
    local total_steps = 10  -- 总步骤数
    
    -- 执行步骤并报告进度
    for i = 1, total_steps do
        -- 执行实际工作
        local result = do_some_work(i, arg1, arg2)
        
        -- 计算进度
        progress = (i / total_steps) * 100
        
        -- yield 暂停并返回状态
        -- 返回格式必须是字典 {}
        coroutine.yield({
            status = "yielded",      -- 必须是 "yielded"
            progress = progress,     -- 进度百分比 (0-100)
            message = "处理步骤 " .. i .. "/" .. total_steps,
            value = result           -- 当前步骤的结果（可选）
        })
        
        -- 在下一次 resume 时，可以接收参数
        -- 例如：local resume_arg1, resume_arg2 = coroutine.yield(...)
    end
    
    -- 最终返回
    -- 返回格式必须是字典 {}
    return {
        status = "done",             -- 完成时不需要此字段（会自动设置）
        progress = 100,              -- 最终进度
        message = "任务完成",
        value = final_result         -- 最终结果
    }
end


-- 模板 2: 带错误处理的协程函数
function robust_coroutine_task(taskCell, input_data)
    -- 使用 pcall 保护执行
    local success, result = pcall(function()
        local progress = 0
        
        -- 数据验证
        if not input_data or input_data == "" then
            error("输入数据无效")
        end
        
        -- 分步处理
        for step = 1, 5 do
            -- 模拟耗时操作
            local step_result = process_step(step, input_data)
            
            progress = step * 20
            
            -- 报告进度
            coroutine.yield({
                status = "yielded",
                progress = progress,
                message = "步骤 " .. step .. " 完成",
                value = step_result
            })
        end
        
        return {
            progress = 100,
            message = "全部完成",
            value = "Success"
        }
    end)
    
    -- 错误处理
    if not success then
        return {
            progress = 0,
            message = "执行失败",
            value = nil,
            error = tostring(result)  -- 错误信息
        }
    end
    
    return result
end


-- 模板 3: 接收 resume 参数的协程函数
-- 在 LuaTask 中定义 resume 参数：
-- =LuaTask("my_task", start_arg, "|", "A1", "B1")
-- 其中 "|" 后的参数是 resume 时从单元格读取的值
function interactive_coroutine_task(taskCell, initial_value)
    local current_value = initial_value
    local step = 0
    
    while step < 5 do
        step = step + 1
        
        -- 执行操作
        current_value = current_value * 2
        
        -- yield 并接收下一次 resume 的参数
        local user_input1, user_input2 = coroutine.yield({
            status = "yielded",
            progress = step * 20,
            message = "等待输入，当前值: " .. current_value,
            value = current_value
        })
        
        -- 使用 resume 传入的参数
        if user_input1 then
            current_value = current_value + user_input1
        end
        if user_input2 then
            current_value = current_value + user_input2
        end
    end
    
    return {
        progress = 100,
        message = "计算完成",
        value = current_value
    }
end


-- 模板 4: 返回复杂数据的协程函数
function data_processing_task(taskCell, data_range)
    -- 处理数组数据
    local results = {}
    local total = #data_range
    
    for i, item in ipairs(data_range) do
        -- 处理单个项目
        local processed = process_item(item)
        table.insert(results, processed)
        
        -- 报告进度
        coroutine.yield({
            status = "yielded",
            progress = (i / total) * 100,
            message = "已处理 " .. i .. "/" .. total,
            value = {
                current_item = item,
                processed_count = i,
                latest_result = processed
            }
        })
    end
    
    return {
        progress = 100,
        message = "数据处理完成",
        value = {
            total_processed = #results,
            results = results,
            summary = "所有项目已处理"
        }
    }
end


-- ============================================
-- 重要说明
-- ============================================

--[[ 
1. 函数签名规则：
   - 第一个参数必须是 taskCell（任务单元格地址）
   - 后续参数对应 LuaTask 的启动参数（"|" 之前）
   - resume 参数通过 coroutine.yield() 的返回值接收

2. yield 返回格式（为一节或二阶列表、字典。列表默认作为为value，字典按如下规则）：
   {
       status = "yielded",      -- 可选，应为yielded、done、error，指挥VBA调度器接下来怎么处理此协程，yielded：等待下一次resume；done：提前结束，被清理出协程队列；error：手动触发VBA调度错误，被清理出队列。如果省略此字段则默认视为yielded
       progress = 50,           -- 可选，进度百分比
       message = "消息",        -- 可选，状态消息
       value = result_data      -- 可选，当前结果，单值或列表
   }

3. return 返回格式（为一阶或二阶列表、字典。列表默认作为value）：
   {
       status = "done",         -- 可选，此字段一般省略，字段会被自动设置为 "done"
       progress = 100,          -- 可选，最终进度
       message = "完成",        -- 可选，完成消息
       value = final_result     -- 可选，最终结果，单值或列表
   }

4. Excel 中读取结果：
   - =LuaGet(taskId, "status")   -> 获取状态
   - =LuaGet(taskId, "progress") -> 获取进度
   - =LuaGet(taskId, "message")  -> 获取消息
   - =LuaGet(taskId, "value")    -> 获取结果值
   - =LuaGet(taskId, "error")    -> 获取错误信息

5. 启动协程：
   在 VBA 中调用：StartLuaCoroutine(taskId)
   或使用宏按钮绑定

6. 调度器配置：
   - g_MaxIterationsPerTick: 每次调度执行的任务数
   - g_SchedulerIntervalSec: 调度间隔（秒）
]]