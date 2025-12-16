-- ============================================
-- functions.lua - Excel 自定义 Lua 函数库
-- ============================================

-- 求和函数（支持多个参数，包括表）
function sum(...)
    local total = 0
    local args = {...}
    
    for _, v in ipairs(args) do
        if type(v) == "number" then
            total = total + v
        elseif type(v) == "table" then
            -- 递归处理二维表
            for _, row in ipairs(v) do
                if type(row) == "table" then
                    for _, cell in ipairs(row) do
                        if type(cell) == "number" then
                            total = total + cell
                        end
                    end
                elseif type(row) == "number" then
                    total = total + row
                end
            end
        end
    end
    
    return total
end

-- 矩阵转置
function transpose(matrix)
    if type(matrix) ~= "table" or #matrix == 0 then
        return {}
    end
    
    local result = {}
    local rows = #matrix
    local cols = type(matrix[1]) == "table" and #matrix[1] or 1
    
    for j = 1, cols do
        result[j] = {}
        for i = 1, rows do
            if type(matrix[i]) == "table" then
                result[j][i] = matrix[i][j]
            else
                result[j][i] = matrix[i]
            end
        end
    end
    
    return result
end

-- 矩阵乘法
function matrix_multiply(a, b)
    if type(a) ~= "table" or type(b) ~= "table" then
        return "错误：参数必须是表"
    end
    
    local rows_a = #a
    local cols_a = type(a[1]) == "table" and #a[1] or 1
    local rows_b = #b
    local cols_b = type(b[1]) == "table" and #b[1] or 1
    
    if cols_a ~= rows_b then
        return "错误：矩阵维度不匹配"
    end
    
    local result = {}
    for i = 1, rows_a do
        result[i] = {}
        for j = 1, cols_b do
            local sum_val = 0
            for k = 1, cols_a do
                local val_a = type(a[i]) == "table" and a[i][k] or a[i]
                local val_b = type(b[k]) == "table" and b[k][j] or b[k]
                sum_val = sum_val + (val_a or 0) * (val_b or 0)
            end
            result[i][j] = sum_val
        end
    end
    
    return result
end

-- 统计函数（返回多个值：总和、平均、最小、最大）
function stats(...)
    local total = 0
    local count = 0
    local min_val = nil
    local max_val = nil
    local args = {...}
    
    for _, v in ipairs(args) do
        if type(v) == "number" then
            total = total + v
            count = count + 1
            if min_val == nil or v < min_val then min_val = v end
            if max_val == nil or v > max_val then max_val = v end
        elseif type(v) == "table" then
            for _, row in ipairs(v) do
                if type(row) == "table" then
                    for _, cell in ipairs(row) do
                        if type(cell) == "number" then
                            total = total + cell
                            count = count + 1
                            if min_val == nil or cell < min_val then min_val = cell end
                            if max_val == nil or cell > max_val then max_val = cell end
                        end
                    end
                elseif type(row) == "number" then
                    total = total + row
                    count = count + 1
                    if min_val == nil or row < min_val then min_val = row end
                    if max_val == nil or row > max_val then max_val = row end
                end
            end
        end
    end
    
    local avg = count > 0 and total / count or 0
    return total, avg, min_val or 0, max_val or 0
end

-- 过滤大于阈值的值
function filter_greater(threshold, data)
    if type(data) ~= "table" then
        return {}
    end
    
    local result = {}
    for i, row in ipairs(data) do
        result[i] = {}
        if type(row) == "table" then
            for j, cell in ipairs(row) do
                if type(cell) == "number" and cell > threshold then
                    table.insert(result[i], cell)
                end
            end
        elseif type(row) == "number" and row > threshold then
            table.insert(result[i], row)
        end
    end
    
    return result
end

-- 自定义公式计算（支持字符串表达式）
function calc(expr, data)
    -- 简单示例：将 data 中的值应用到表达式
    if type(data) == "table" and type(data[1]) == "table" then
        local result = {}
        for i, row in ipairs(data) do
            result[i] = {}
            for j, cell in ipairs(row) do
                if type(cell) == "number" then
                    -- 用 x 代替单元格值
                    local formula = string.gsub(expr, "x", tostring(cell))
                    local func = load("return " .. formula)
                    if func then
                        result[i][j] = func()
                    else
                        result[i][j] = cell
                    end
                else
                    result[i][j] = cell
                end
            end
        end
        return result
    end
    
    return data
end

-- 生成序列
function sequence(start, stop, step)
    step = step or 1
    local result = {{}}
    local idx = 1
    
    for i = start, stop, step do
        result[1][idx] = i
        idx = idx + 1
    end
    
    return result
end

-- 生成随机矩阵
function random_matrix(rows, cols, min_val, max_val)
    min_val = min_val or 0
    max_val = max_val or 100
    
    local result = {}
    math.randomseed(os.time())
    
    for i = 1, rows do
        result[i] = {}
        for j = 1, cols do
            result[i][j] = math.random(min_val, max_val)
        end
    end
    
    return result
end

-- 字符串处理：连接所有参数
function concat_all(sep, ...)
    sep = sep or ", "
    local parts = {}
    local args = {...}
    
    for _, v in ipairs(args) do
        if type(v) == "table" then
            for _, row in ipairs(v) do
                if type(row) == "table" then
                    for _, cell in ipairs(row) do
                        table.insert(parts, tostring(cell))
                    end
                else
                    table.insert(parts, tostring(row))
                end
            end
        else
            table.insert(parts, tostring(v))
        end
    end
    
    return table.concat(parts, sep)
end

print("functions.lua 加载成功！")


function test()
    return {{1,3},{9,4},{4,12}}
end

function loop_test(n)
    while true do
        n = n + 1
        if n > 100000000000000 then
            break
        end
    end
    return 12
end

-- ============================================
-- functions.lua - Lua 协程测试示例
-- ============================================

-- ============================================
-- Lua 协程任务模板
-- 说明：
--   taskCell: Excel 调用的单元格地址（VBA端传入）
--   args: 可选的启动参数
-- 返回：
--   协程 yield / return 字典格式
-- ============================================

local function MyTask(taskCell, ...)
    -- 启动参数
    local args = {...}

    -- 初始化状态
    local progress = 0
    local message = "任务已启动"
    local value = nil

    -- yield 第一次状态（初始广播）
    coroutine.yield{
        status = "running",   -- running / yielded / done / error
        progress = progress,
        message = message,
        value = value
    }

    -- 模拟长计算或多步任务
    for i = 1, 10 do
        -- 假设这里是计算逻辑
        progress = i * 10
        value = i * i  -- 示例值，可替换为计算结果

        message = "正在计算，第 " .. i .. " 步"

        -- yield 中间状态
        coroutine.yield{
            status = "yielded",
            progress = progress,
            message = message,
            value = value,
        }
    end

    -- 任务完成，返回最终结果
    return {
        status = "done",
        progress = 100,
        message = "任务完成",
        value = value,
    }
end

---- 将函数注册为全局函数供 VBA 调用
--_G["MyTask"] = MyTask



-- 示例1：无限循环计数器（每次 yield 增加计数）
function counter_infinite(taskCell, startValue)
    local count = startValue or 0
    local step = 1
    
    while true do
        count = count + step
        
        -- 每次 yield 返回当前状态
        coroutine.yield({
            status = "running",
            progress = math.min(count, 100),  -- 进度条最多显示到100
            message = "计数: " .. count,
            value = {{count, count * 2, count * 3}},  -- 返回一行三列数据
        })
    end
end


-- 示例2：无限监控（读取单元格并处理）
function monitor_cell(taskCell, targetCell)
    local iteration = 0
    
    while true do
        iteration = iteration + 1
        
        -- 在 resume 时会获取到 targetCell 的最新值
        local cellValue = coroutine.yield({
            status = "running",
            progress = (iteration % 100),
            message = "监控中，第 " .. iteration .. " 次检查",
            value = {{iteration, "等待数据..."}},
        })
        
        -- 处理获取到的值
        if cellValue then
            local result = cellValue * 2  -- 简单处理：乘以2
            
            coroutine.yield({
                status = "running",
                progress = (iteration % 100),
                message = "处理: " .. cellValue .. " -> " .. result,
                value = {{iteration, cellValue, result}},
            })
        end
    end
end


-- 示例3：累加器（持续累加 resume 参数）
function accumulator(taskCell, initialSum)
    local sum = initialSum or 0
    local count = 0
    
    while true do
        count = count + 1
        
        -- yield 并等待新的值
        local newValue = coroutine.yield({
            status = "running",
            progress = math.min(count * 5, 100),
            message = "当前总和: " .. sum .. " (已累加 " .. count .. " 次)",
            value = {{count, sum, sum / count}},  -- 次数、总和、平均值
        })
        
        -- 累加新值
        if newValue and type(newValue) == "number" then
            sum = sum + newValue
        end
    end
end


-- 示例4：状态机（循环切换状态）
function state_machine(taskCell)
    local states = {"待机", "工作中", "暂停", "恢复"}
    local stateIndex = 1
    local iteration = 0
    
    while true do
        iteration = iteration + 1
        local currentState = states[stateIndex]
        
        coroutine.yield({
            status = "running",
            progress = (iteration % 100),
            message = "状态: " .. currentState,
            value = {{iteration, currentState, stateIndex}},
        })
        
        -- 切换到下一个状态
        stateIndex = stateIndex + 1
        if stateIndex > #states then
            stateIndex = 1
        end
    end
end


-- 示例5：进度模拟（无限循环，但有进度条）
function progress_simulator(taskCell, maxSteps)
    local steps = maxSteps or 100
    local currentStep = 0

    while true do
        currentStep = currentStep + 1
        if currentStep > steps then
            currentStep = 1
        end

        local progress = (currentStep / steps) * 100

        coroutine.yield({
            status = "yielded",
            progress = progress,
            message = string.format(
                "progress: %d/%d (%.1f%%)",
                currentStep, steps, progress
            ),
            value = {{currentStep, steps, progress}}
        })
    end
end



-- 示例6：数据流处理（持续接收和处理数据）
function data_stream(taskCell)
    local processedCount = 0
    local totalSum = 0
    
    while true do
        processedCount = processedCount + 1
        
        -- 接收新数据（从 resume 参数）
        local data = coroutine.yield({
            status = "running",
            progress = math.min(processedCount, 100),
            message = "已处理 " .. processedCount .. " 条数据",
            value = {{processedCount, totalSum, totalSum / processedCount}},
        })
        
        -- 处理数据
        if data and type(data) == "number" then
            totalSum = totalSum + data
        end
    end
end


-- 示例7：时间戳记录器
function timestamp_logger(taskCell)
    local logs = {}
    local logCount = 0
    
    while true do
        logCount = logCount + 1
        local timestamp = os.date("%H:%M:%S")
        
        -- 保留最近10条记录
        table.insert(logs, timestamp)
        if #logs > 10 then
            table.remove(logs, 1)
        end
        
        local logString = table.concat(logs, ", ")
        
        coroutine.yield({
            status = "running",
            progress = (logCount % 100),
            message = "记录 #" .. logCount .. ": " .. timestamp,
            value = {{logCount, timestamp}},
        })
    end
end


-- 示例8：条件控制的无限循环（可通过 resume 参数控制）
function controlled_loop(taskCell)
    local iteration = 0
    local running = true
    
    while true do
        iteration = iteration + 1
        
        -- 获取控制命令
        local command = coroutine.yield({
            status = running and "running" or "paused",
            progress = (iteration % 100),
            message = running and ("运行中: " .. iteration) or "已暂停",
            value = {{iteration, running and "运行" or "暂停"}},
        })
        
        -- 处理命令
        if command == "pause" then
            running = false
        elseif command == "resume" then
            running = true
        elseif command == "stop" then
            break  -- 退出循环
        end
        
        -- 只有运行时才继续
        if not running then
            -- 暂停状态，等待下一个命令
        end
    end
    
    -- 退出时返回
    return {
        value = "已停止",
    }
end


-- 示例9：简单的心跳检测
function heartbeat(taskCell)
    local beatCount = 0
    
    while true do
        beatCount = beatCount + 1
        local isAlive = (beatCount % 2 == 0) and "💚" or "🤍"
        
        coroutine.yield({
            status = "running",
            progress = 50,
            message = "心跳 " .. isAlive,
            value = {{beatCount, isAlive}},
        })
    end
end


-- 示例10：多值返回测试
function multi_value_test(taskCell)
    local count = 0
    
    while true do
        count = count + 1
        
        -- 返回一个大的二维数组
        local data = {}
        for i = 1, 5 do
            data[i] = {count + i - 1, (count + i - 1) * 2, (count + i - 1) * 3}
        end
        
        coroutine.yield({
            status = "running",
            progress = (count % 100),
            message = "生成 5x3 数据表",
            value = data,
        })
    end
end

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