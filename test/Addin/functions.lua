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

-- ============================================
-- functions.lua - Lua 协程测试示例
-- ============================================

-- ============================================
-- 协程测试函数：测试所有支持的参数类型
-- ============================================

-- 测试函数：全类型参数测试
-- 启动参数测试：数值、字符串、布尔、数组
-- Resume参数测试：字面量、单元格引用、动态字符串
function test_all_types(num_param, str_param, bool_param, arr_param)
    -- 第一次 yield：报告启动参数
    local start_report = {
        status = "yield",
        progress = 10,
        message = "启动参数接收完成",
        value = {
            {"参数类型", "参数值", "Lua类型"},
            {"数值参数", tostring(num_param), type(num_param)},
            {"字符串参数", tostring(str_param), type(str_param)},
            {"布尔参数", tostring(bool_param), type(bool_param)},
            {"数组参数", arr_param and "已接收" or "nil", type(arr_param)}
        }
    }
    
    -- 第一次 resume：接收字面量参数
    local literal_val = coroutine.yield(start_report)
    
    local resume1_report = {
        status = "yield",
        progress = 30,
        message = "Resume#1: 字面量参数",
        value = {
            {"Resume参数", "值", "类型"},
            {"字面量", tostring(literal_val), type(literal_val)}
        }
    }
    
    -- 第二次 resume：接收单元格引用
    local cell_val = coroutine.yield(resume1_report)
    
    local resume2_report = {
        status = "yield",
        progress = 50,
        message = "Resume#2: 单元格引用",
        value = {
            {"Resume参数", "值", "类型"},
            {"单元格值", tostring(cell_val), type(cell_val)}
        }
    }
    
    -- 第三次 resume：接收动态字符串引用
    local dynamic_val = coroutine.yield(resume2_report)
    
    local resume3_report = {
        status = "yield",
        progress = 70,
        message = "Resume#3: 动态字符串",
        value = {
            {"Resume参数", "值", "类型"},
            {"动态引用值", tostring(dynamic_val), type(dynamic_val)}
        }
    }
    
    -- 第四次 resume：接收多个混合参数
    local mix1, mix2, mix3 = coroutine.yield(resume3_report)
    
    local resume4_report = {
        status = "yield",
        progress = 90,
        message = "Resume#4: 多参数混合",
        value = {
            {"参数序号", "值", "类型"},
            {"参数1", tostring(mix1), type(mix1)},
            {"参数2", tostring(mix2), type(mix2)},
            {"参数3", tostring(mix3), type(mix3)}
        }
    }
    
    -- 最后一次 yield 后完成
    coroutine.yield(resume4_report)
    
    -- 返回最终结果
    return {
        status = "done",
        progress = 100,
        message = "全类型测试完成",
        value = {
            {"测试项", "结果"},
            {"启动参数", "通过"},
            {"字面量Resume", "通过"},
            {"单元格引用", "通过"},
            {"动态字符串", "通过"},
            {"多参数混合", "通过"},
        }
    }
end

-- 测试函数：边界情况测试 OK
function test_edge_cases(empty_val, zero_val, negative_val, long_str)
    local results = {
        {"参数", "值", "类型", "判定"},
        {"空值", tostring(empty_val), type(empty_val), empty_val == nil and "正确:nil" or "有值"},
        {"零值", tostring(zero_val), type(zero_val), zero_val == 0 and "正确:0" or "非零"},
        {"负数", tostring(negative_val), type(negative_val), (negative_val or 0) < 0 and "正确:负数" or "非负"},
        {"长字符串", string.len(tostring(long_str or "")) .. "字符", type(long_str), "已接收"}
    }

    local report = {
        status = "yield",
        progress = 50,
        message = "边界值分析",
        value = results
    }

    -- Resume 测试空值和特殊值
    local resume_empty, resume_zero, resume_bool = coroutine.yield(report)

    return {
        status = "done",
        progress = 100,
        message = "边界测试完成",
        value = {
            {"Resume参数", "值", "类型"},
            {"空参数", tostring(resume_empty), type(resume_empty)},
            {"零参数", tostring(resume_zero), type(resume_zero)},
            {"布尔参数", tostring(resume_bool), type(resume_bool)}
        }
    }
end

-- 测试函数：错误处理测试 OK
function test_error_handling(should_error)
    local report = {
        status = "yield",
        progress = 50,
        message = "准备测试错误处理",
        value = {{"should_error", tostring(should_error)}}
    }

    coroutine.yield(report)

    if should_error then
        error("这是一个预期的测试错误")
    end

    return {
        status = "done",
        progress = 100,
        message = "错误测试完成（无错误发生）",
        value = "正常完成"
    }
end

-- 测试函数：返回值类型测试 OK
function test_return_types()
    -- 测试不同返回类型
    -- 返回字符串
    coroutine.yield({
        status = "yield",
        progress = 0,
        message = "返回字符串",
        value = "这是测试字符串，支持中文！"
    })

    -- 返回数值
    coroutine.yield({
        status = "yield",
        progress = 20,
        message = "返回数值",
        value = 12345.678
    })

    -- 返回布尔
    coroutine.yield({
        status = "yield",
        progress = 40,
        message = "返回布尔",
        value = true
    })

    -- 返回字符串
    coroutine.yield({
        status = "yield",
        progress = 60,
        message = "返回nil",
        value = nil
    })

    -- 返回一维数组
    coroutine.yield({
        status = "yield",
        progress = 80,
        message = "返回一维数组",
        value = {1, 2}
    })
    -- 返回二维数组（表格数据）
    return {
        status = "done",
        progress = 100,
        message = "返回二维数组",
        value = {
            {"A", "B"},
            {1, nil},
        }
    }
end
-- ============================================
-- 重要说明
-- ============================================

--[[ 
1. 函数签名规则：
   - 后续参数对应 LuaTask 的启动参数（"|" 之前）
   - resume 参数通过 coroutine.yield() 的返回值接收

2. yield 返回格式（为一节或二阶列表、字典。列表默认作为为value，字典按如下规则）：
   {
       status = "yield",      -- 可选，应为yield、done、error，指挥VBA调度器接下来怎么处理此协程，yield：等待下一次resume；done：提前结束，被清理出协程队列；error：手动触发VBA调度错误，被清理出队列。如果省略此字段则默认视为yield
       progress = 50,           -- 可选，进度百分比
       message = "消息",        -- 可选，状态消息
       value = result_data      -- 可选，当前结果，单值、列表、无嵌套字典
   }

3. return 协程的return格式（单值、一阶或二阶列表、字典（会尝试按以下格式解析几种值）；普通函数return格式（单值、一阶或二阶列表、字典（不支持嵌套，将作为二阶列表返回excel））
   {
       status = "done",         -- 可选，此字段一般省略，字段会被自动设置为 "done"
       progress = 100,          -- 可选，最终进度
       message = "完成",        -- 可选，完成消息
       value = final_result     -- 可选，最终结果，单值、列表、无嵌套字典
   }

4. Excel 中读取结果：
   - =LuaGet(taskId, "status")   -> 获取状态
   - =LuaGet(taskId, "progress") -> 获取进度
   - =LuaGet(taskId, "message")  -> 获取消息
   - =LuaGet(taskId, "value")    -> 获取结果值
   - =LuaGet(taskId, "error")    -> 获取错误信息
]]


-- 去重累加函数：只有当值变化时才累加
-- 参数：
--   initialValue（可选）: 累加的起始值，默认为0
-- Resume参数：
--   要读取的数值（来自单元格引用）
function accumulate(initialValue)
    local sum = initialValue or 0
    local count = 0
    local lastValue = nil  -- 记录上一次的值

    -- 首次 yield，等待第一个值
    local input = coroutine.yield({
        status = "yield",
        progress = count,
        message = "等待输入...",
        value = sum
    })

    while true do
        local num = tonumber(input)

        -- 只有当值存在且与上次不同时才累加
        if num ~= nil and num ~= lastValue then
            sum = sum + num
            count = count + 1
            lastValue = num

            input = coroutine.yield({
                status = "yield",
                progress = count,
                message = string.format("新增: %s, 已累加 %d 次", num, count),
                value = sum
            })
        else
            -- 值未变化，不累加
            local msg
            if num == nil then
                msg = "无效输入，跳过"
            else
                msg = string.format("值未变化(%s)，跳过", num)
            end

            input = coroutine.yield({
                status = "yield",
                progress = count,
                message = msg,
                value = sum
            })
        end
    end
end

function accumulate_once(lastValue, newValue)
    return {
        value = lastValue + newValue
    }
end

-- functions.lua
-- 两个相同大小的区域，对应位置相加，结果写回第二个区域
function matrixAdd(matrix1, matrix2, rows, cols)
    local result = {}
    coroutine.yield()

    for i = 1, rows do
        result[i] = {}
        for j = 1, cols do
            -- 获取值，nil 当作 0
            local val1 = 0
            local val2 = 0
            -- 检查 matrix1[i] 是否存在
            if matrix1[i] then
                val1 = matrix1[i][j] or 0
            end
            -- 检查 matrix2[i] 是否存在
            if matrix2[i] then
                val2 = matrix2[i][j] or 0
            end
            result[i][j] = val1 + val2
        end
    end
    result.__size = {rows,cols}
    -- return {value = result}
    return {
        100, 200, 300,      -- 数组部分：索引 1, 2, 3
        name = "混合表",     -- 字典部分
        count = 3
    }
end

function matrixAddNil(matrix1, matrix2, rows, cols)
    local result = {}

    for i = 1, rows do
        result[i] = {}
        for j = 1, cols do
            -- 获取值，nil 当作 0
            local val1 = matrix1[i][j]
            local val2 = matrix2[i][j]
            -- 检查 matrix1[i] 是否存在
            if val1 == nil and val2 == nil then
                result[i][j] = nil
            elseif val1 == nil then
                result[i][j] = val2
            elseif val2 == nil then
                result[i][j] = val1
            else
                result[i][j] = val1+val2
            end
        end
    end
    result.__size = {rows,cols}
    return result
end
-- ============================================
-- 字典返回值测试函数集
-- ============================================

-- 1. 简单字典
function test_simple_dict()
    return {
        name = "张三",
        age = 25,
        city = "北京"
    }
end

-- 2. 混合类型值的字典
function test_mixed_dict()
    return {
        str_val = "hello",
        num_val = 3.14159,
        bool_val = true,
        nil_val = nil,  -- nil 会被忽略
        int_val = 42
    }
end

-- 3. 纯数组（对比用）
function test_pure_array()
    return {10, 20, 30, 40, 50}
end

-- 4. 混合 table（既有数组索引又有字符串键）
function test_mixed_table()
    return {
        100, 200, 300,      -- 数组部分：索引 1, 2, 3
        name = "混合表",     -- 字典部分
        count = 3
    }
end

-- 6. 空字典
function test_empty_dict()
    return {}
end

-- 7. 单键字典
function test_single_key()
    return {
        only_key = "only_value"
    }
end

-- 8. 数字键字典（非连续）
function test_sparse_dict()
    return {
        [1] = "一",
        [5] = "五",
        [10] = "十"
    }
end

-- 9. 中文键字典
function test_chinese_keys()
    return {
        ["姓名"] = "王五",
        ["年龄"] = 28,
        ["城市"] = "广州"
    }
end

-- 协程中使用（yield）
function test_coroutine_compound()
    for i = 1, 5 do
        coroutine.yield({
            progress = i * 20,
            message = "正在处理第 " .. i .. " 步",
            value = {step = i, data = "step_" .. i}
        })
    end
    
    return {
        status = "done",
        message = "全部完成",
        value = {total = 5, result = "success"}
    }
end

function test_coroutine_return_nested_compound()
    coroutine.yield({value = {name = "张六", score = 95}})
    return {value = {name = "张六", score = 95}}
end