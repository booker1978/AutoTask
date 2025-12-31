using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Conditions; // 必须添加这个引用
using FlaUI.Core.Tools; // 用于 Retry
using FlaUI.UIA3;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using static ToolsAutoTask.ScriptParser;
using FlaUIApp = FlaUI.Core.Application;

namespace ToolsAutoTask
{
    public class ScriptParser
    {
        public enum ActionType
        {
            ATTACH,         // 附加进程
            CLICK,          // 点击
            SELECT,         // 下拉框/列表选择
            INPUT,          // 输入文本
            TYPE,           // 模拟纯键盘输入，支持{ENTER}等特殊字符
            WAIT,           // 等待
            ASSERT_TEXT,    // 验证文本包含
            ASSERT_ENABLE,  // 验证按钮可用
            ASSERT_CHECKED, // 验证复选框状态
            LOG,            // 记录日志
            LOOP,           // 循环控制
            END_LOOP,
            // === 【新增】鼠标模拟指令 ===
            MOUSE_MOVE,      // 移动鼠标 (Target: "x,y")
            MOUSE_CLICK,     // 左键单击 (Target: "x,y")
            MOUSE_DBLCLICK,  // 左键双击 (Target: "x,y")
            MOUSE_RCLICK,    // 右键单击 (Target: "x,y")
            MOUSE_DRAG,       // 拖拽 (Target: "x1,y1", Value: "x2,y2")
            MOUSE_SCROLL,    // 滚轮指令 (Target: "x,y" ,Value:  "5" 或 "-5")

            // === 【新增】无效指令标记 ===
            INVALID // 用于标记无法识别的指令
        }

        public class ScriptLine
        {
            public int LineNumber { get; set; }
            public ActionType Action { get; set; }
            public string Target { get; set; } // 例如 "Name:登录"
            public string Value { get; set; }  // 例如 "admin"
            public string RawCommand { get; set; } // 【新增】记录原始指令字符串，方便报错
            public string ErrorMessage { get; set; } // 【新增】记录解析时的错误信息
        }

        public class LoopContext
        {
            public int StartLineIndex { get; set; } // 循环开始的那一行(LOOP指令所在行)
            public int RemainingCount { get; set; } // 剩余次数

            public LoopContext(int index, int count)
            {
                StartLineIndex = index;
                RemainingCount = count;
            }
        }
        public static List<ScriptLine> Parse(string scriptText)
        {
            var lines = new List<ScriptLine>();
            if (string.IsNullOrWhiteSpace(scriptText)) return lines;

            var rawLines = scriptText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < rawLines.Length; i++)
            {
                string currentLine = rawLines[i].Trim();
                // 跳过空行和注释
                if (string.IsNullOrEmpty(currentLine) || currentLine.StartsWith("//")) continue;

                var parts = currentLine.Split(new[] { ' ' }, 3); // 最多切成3段
                var cmd = new ScriptLine
                {
                    LineNumber = i + 1,
                    Target = parts.Length > 1 ? parts[1] : null,
                    Value = parts.Length > 2 ? parts[2] : null,
                    RawCommand = parts[0] // 记录原始指令
                };

                // === 【核心修改】使用 TryParse 避免崩溃 ===
                if (Enum.TryParse(parts[0], true, out ActionType parsedAction))
                {
                    cmd.Action = parsedAction;

                    // === 【进阶】简单的参数校验 ===
                    // 比如 CLICK 指令必须有 Target
                    if (IsTargetRequired(parsedAction) && string.IsNullOrEmpty(cmd.Target))
                    {
                        cmd.Action = ActionType.INVALID;
                        cmd.ErrorMessage = $"指令 {parsedAction} 缺少必要的定位符参数";
                    }
                }
                else
                {
                    // 如果解析失败，标记为无效
                    cmd.Action = ActionType.INVALID;
                    cmd.ErrorMessage = $"未知的指令: '{parts[0]}'";
                }

                lines.Add(cmd);
            }
            return lines;
        }

        // 辅助方法：判断哪些指令必须带参数
        private static bool IsTargetRequired(ActionType action)
        {
            switch (action)
            {
                case ActionType.CLICK:
                case ActionType.INPUT:
                case ActionType.SELECT:
                case ActionType.ASSERT_TEXT:
                case ActionType.ASSERT_ENABLE:
                case ActionType.MOUSE_CLICK:
                case ActionType.MOUSE_MOVE:
                    return true;
                default:
                    return false;
            }
        }
    }
    public class AutomationEngine
    {
        private AutomationElement _window;
        private UIA3Automation _automation;
        private string _lastProcessName = "";
        private FlaUI.Core.Application _application; // 必须有这个，用于 AttachToProcess
        // 【新增】鼠标控制器实例
        private MouseController _mouseController;

        public event Action<string> OnLog; // 用于向界面反馈日志
                                           // 【新增】停止标记 (volatile 确保多线程也能读到最新值)
        private volatile bool _isCancellationRequested = false;

        //public AutomationEngine()
        //{
        //    // 在构造函数或 ExecuteScript 开头初始化
        //    // 传递 Log 方法给它，让它也能打日志
        //    _mouseController = new MouseController(this.Log);
        //}
        // 【新增】外部调用的停止方法
        public void Stop()
        {
            _isCancellationRequested = true;
        }
        public bool IgnoreError { get; set; } = false;

        private void Log(string msg) => OnLog?.Invoke($"[{DateTime.Now:HH:mm:ss}] {msg}");
        // 在 AutomationEngine 类中添加
        private void EnsureWindowReady()
        {
            // 【修改】使用 IsOffscreenSafe 替代直接调用
            if (_window == null || IsOffscreenSafe(_window))
            {
                // 如果窗口对象丢了，或者在屏幕外，尝试重新连接
                if (!string.IsNullOrEmpty(_lastProcessName))
                {
                    try
                    {
                        AttachToProcess(_lastProcessName);
                    }
                    catch
                    {
                        // 重连失败暂时忽略，由后续步骤抛出异常
                    }
                }
            }

            if (_window != null)
            {
                try
                {
                    // 1. 检查是否最小化
                    if (_window.Patterns.Window.IsSupported)
                    {
                        var windowPattern = _window.Patterns.Window.Pattern;
                        if (windowPattern.WindowVisualState.Value == FlaUI.Core.Definitions.WindowVisualState.Minimized)
                        {
                            Log("[自动维护] 窗口已最小化，正在还原...");
                            windowPattern.SetWindowVisualState(FlaUI.Core.Definitions.WindowVisualState.Normal);
                            Thread.Sleep(500);
                        }
                    }

                    // 2. 强制置顶
                    _window.Focus();
                }
                catch (Exception ex)
                {
                    // 这里只记录警告，不要抛出异常中断流程
                    // Log($"[自动维护警告] 尝试唤醒窗口失败: {ex.Message}");
                }
            }
        }
        public void ExecuteScript(string script)
        {
            // =========================================================
            // 【核心修复】 必须在每次开始执行时，强制把红灯变绿灯！
            // =========================================================
            _isCancellationRequested = false;

            Stack<LoopContext> loopStack = new Stack<LoopContext>();
            var commands = ScriptParser.Parse(script);
            _automation = new FlaUI.UIA3.UIA3Automation();

            // 确保每次执行前初始化（如果不想在构造函数初始化）
            if (_mouseController == null) _mouseController = new MouseController(this.Log);

            Log($"=== 开始执行 (跳过错误: {IgnoreError}) ===");

            try
            {
                for (int i = 0; i < commands.Count; i++)
                {
                    // 1. 检查停止标记
                    if (_isCancellationRequested)
                    {
                        Log(">> 用户手动停止了脚本 <<");
                        break;
                    }

                    var cmd = commands[i];

                    try
                    {
                        // 2. 检查无效指令
                        if (cmd.Action == ScriptParser.ActionType.INVALID)
                        {
                            string errorMsg = $"[语法错误] 第 {cmd.LineNumber} 行: {cmd.ErrorMessage}";
                            if (IgnoreError) { Log(errorMsg); continue; }
                            else throw new Exception(errorMsg);
                        }

                        Log($"正在执行行 {cmd.LineNumber}: {cmd.Action}...");

                        // 3. 处理循环逻辑 (LOOP / END_LOOP)
                        if (cmd.Action == ScriptParser.ActionType.LOOP)
                        {
                            int count = int.Parse(cmd.Target);
                            if (loopStack.Count > 0 && loopStack.Peek().StartLineIndex == i) { }
                            else { loopStack.Push(new LoopContext(i, count)); }
                            continue; // LOOP 指令本身不执行 ExecuteCommand
                        }
                        else if (cmd.Action == ScriptParser.ActionType.END_LOOP)
                        {
                            if (loopStack.Count == 0) throw new Exception("发现多余的 END_LOOP");
                            var context = loopStack.Peek();
                            context.RemainingCount--;
                            if (context.RemainingCount > 0)
                            {
                                Log($"循环中... 剩余 {context.RemainingCount} 次");
                                i = context.StartLineIndex - 1;
                            }
                            else
                            {
                                loopStack.Pop();
                                Log("循环块结束");
                            }
                            continue; // END_LOOP 指令本身不执行 ExecuteCommand
                        }

                        // 4. 【关键】执行指令 (全场唯一入口)
                        ExecuteCommand(cmd);
                    }
                    catch (Exception ex)
                    {
                        if (IgnoreError)
                            Log($"[警告-已跳过] 第 {cmd.LineNumber} 行执行出错: {ex.Message}");
                        else
                            throw;
                    }
                }

                // 如果是因为停止而退出的，最后记录一下
                if (_isCancellationRequested)
                    Log("=== 脚本已终止 ===");
                else
                    Log("=== 脚本全部执行完毕 ===");
            }
            catch (Exception ex)
            {
                // 如果是 OperationCanceledException，说明是正常的停止
                if (ex is OperationCanceledException || _isCancellationRequested)
                {
                    Log("=== 脚本已终止 (用户停止) ===");
                }
                else
                {
                    Log($"!!! 执行被中断 !!! ... {ex.Message}");
                }
            }
            finally
            {
                _automation?.Dispose();
            }
        }
        // 【新增】统一的停止检查
        private void CheckCancellation()
        {
            if (_isCancellationRequested)
            {
                // 抛出一个特殊的异常，或者普通的异常，外层会捕获并识别为停止
                throw new OperationCanceledException("用户手动停止了脚本");
            }
        }

        private void ExecuteCommand(ScriptLine cmd)
        {
            try
            {
                switch (cmd.Action)
                {
                    case ActionType.ATTACH:
                        AttachToProcess(cmd.Target);
                        break;

                    case ActionType.WAIT:
                        int waitTime = int.Parse(cmd.Target);
                        Log($"等待 {waitTime} 毫秒...");

                        // 【优化】切片等待：每 100ms 醒来一次检查是否要停止
                        int elapsed = 0;
                        while (elapsed < waitTime)
                        {
                            // 如果发现停止标记，立刻抛出异常或退出
                            if (_isCancellationRequested)
                            {
                                // 这里可以直接 break，让外层的检查逻辑去处理停止
                                break;
                            }

                            Thread.Sleep(100); // 每次只睡 0.1 秒
                            elapsed += 100;
                        }
                        break;

                    // === 日志输出 (修复报错的关键) ===
                    case ActionType.LOG:
                        // ScriptParser 可能会把日志内容切成 Target 和 Value 两段
                        // 例如: LOG === 开始测试 ===
                        // Target: ===
                        // Value: 开始测试 ===
                        string logMsg = cmd.Target;
                        if (!string.IsNullOrEmpty(cmd.Value))
                        {
                            logMsg += " " + cmd.Value;
                        }
                        Log($"[脚本] {logMsg}");
                        break;

                    case ActionType.CLICK:
                        //var btn = FindElement(cmd.Target).AsButton();
                        //if (!btn.IsEnabled) throw new Exception("按钮不可用(Disabled)");
                        //btn.Invoke();
                        //break;
                        // 1. 查找控件 (这里的 FindElement 会调用升级后的 GetControlType)
                        var eleToClick = FindElement(cmd.Target);
                        if (!eleToClick.IsEnabled) throw new Exception("控件不可用(Disabled)");

                        // 2. 【修改】调用万能点击
                        PerformClick(eleToClick);
                        break;

                    case ActionType.INPUT:
                        //var txt = FindElement(cmd.Target).AsTextBox();
                        //txt.Text = cmd.Value;
                        //break;
                        // 1. 查找控件
                        var eleToInput = FindElement(cmd.Target);
                        if (!eleToInput.IsEnabled) throw new Exception("控件不可用(Disabled)");

                        // 2. 【修改】调用万能输入
                        PerformInput(eleToInput, cmd.Value);
                        break;

                    // === 纯键盘输入(新增) ===
                    case ActionType.TYPE:
                        // 1. 处理文本内容
                        // 因为解析器是按空格切分的，如果用户写 "TYPE hello world"
                        // Target 会是 "hello"，Value 会是 "world"
                        // 我们需要把它们拼回来，或者你只取 Target
                        string textToType = cmd.Target;
                        if (!string.IsNullOrEmpty(cmd.Value))
                        {
                            textToType += " " + cmd.Value;
                        }

                        if (string.IsNullOrEmpty(textToType)) return; // 空指令不执行

                        Log($"[键盘模拟] 输入: {textToType}");

                        // 2. 识别特殊按键 (支持大写或小写)
                        // FlaUI 的 Keyboard.Type 支持输入字符串，也支持输入 VirtualKeyShort
                        switch (textToType.ToUpper())
                        {
                            case "{ENTER}":
                                FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.ENTER);
                                break;
                            case "{TAB}":
                                FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.TAB);
                                break;
                            case "{ESC}":
                            case "{ESCAPE}":
                                FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.ESCAPE);
                                break;
                            case "{BACKSPACE}":
                            case "{BS}":
                                FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.BACK);
                                break;
                            case "{DELETE}":
                            case "{DEL}":
                                FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.DELETE);
                                break;
                            case "{SPACE}":
                                FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.SPACE);
                                break;

                            // ... 你可以根据需要扩展更多键，如 F1-F12, UP, DOWN 等 ...

                            default:
                                // 3. 普通文本输入
                                // 注意：Keyboard.Type 输入速度很快，如果不稳定可以在字符间加延迟，
                                // 但 FlaUI 默认实现通常足够好用。
                                FlaUI.Core.Input.Keyboard.Type(textToType);
                                break;
                        }
                        break;

                    case ActionType.SELECT:
                        HandleSelect(cmd.Target, cmd.Value);
                        break;

                    case ActionType.ASSERT_ENABLE:
                        var element = FindElement(cmd.Target);
                        if (!element.IsEnabled) throw new Exception($"验证失败: 控件 {cmd.Target} 处于禁用状态");
                        break;

                    case ActionType.ASSERT_TEXT:
                        CheckText(cmd.Target, cmd.Value);
                        break;

                    case ActionType.ASSERT_CHECKED:
                        // 解析参数：如果脚本写了 "true" 则期待选中，写 "false" 则期待未选中
                        // 默认为 true (即如果不写参数，就默认验证它是否被选中了)
                        bool expectChecked = true;
                        if (!string.IsNullOrEmpty(cmd.Value))
                        {
                            bool.TryParse(cmd.Value, out expectChecked);
                        }

                        CheckCheckBoxState(cmd.Target, expectChecked);
                        break;

                    // === 【新增】鼠标指令分发 ===
                    case ActionType.MOUSE_MOVE:
                    case ActionType.MOUSE_CLICK:
                    case ActionType.MOUSE_DBLCLICK:
                    case ActionType.MOUSE_RCLICK:
                    case ActionType.MOUSE_DRAG:
                    case ActionType.MOUSE_SCROLL:
                        // 1. 确保窗口存在且前置 (利用你现有的 EnsureWindowReady)
                        EnsureWindowReady();
                        // 2. 委托给 MouseController 处理
                        _mouseController.HandleMouseAction(cmd, _window);
                        break;

                    case ActionType.INVALID:
                        // 理论上上面已经拦截了，这里做个双重保险
                        throw new Exception($"无法执行无效指令: {cmd.RawCommand}");

                    default:
                        // 遇到未实现的指令
                        throw new Exception($"指令 {cmd.Action} 尚未实现功能逻辑");
                }
            }
            catch (Exception ex)
            {
                ex.Data["Line"] = cmd.LineNumber;
                throw; // 抛给上层处理
            }
        }

        // --- 核心工具方法 ---
        private void CheckCheckBoxState(string targetStr, bool expectChecked)
        {
            var element = FindElement(targetStr);

            // 转换为 CheckBox 控件
            var checkBox = element.AsCheckBox();

            // 轮询验证状态 (防止刚刚点击还在动画中)
            bool isMatch = FlaUI.Core.Tools.Retry.WhileFalse(() =>
            {
                try
                {
                    // 获取 Toggle 模式
                    if (checkBox.Patterns.Toggle.IsSupported)
                    {
                        var state = checkBox.Patterns.Toggle.Pattern.ToggleState.Value;

                        // FlaUI 的状态有: On(选中), Off(未选), Indeterminate(半选/不确定)
                        bool isChecked = (state == FlaUI.Core.Definitions.ToggleState.On);

                        return isChecked == expectChecked;
                    }
                    else
                    {
                        // 如果不支持 Toggle 模式，尝试降级用 Win32 的状态判断 (兼容旧程序)
                        // 或者直接抛错
                        throw new Exception("该控件不支持 Toggle 模式，无法判断选中状态");
                    }
                }
                catch
                {
                    return false;
                }
            }, TimeSpan.FromSeconds(3)).Result;

            if (!isMatch)
            {
                throw new Exception($"复选框状态验证失败！\r\n目标: {targetStr}\r\n预期状态: {(expectChecked ? "选中" : "未选中")}\r\n实际状态: {(expectChecked ? "未选中" : "选中")}");
            }

            Log($"验证通过: 复选框 {targetStr} 状态正确");
        }

        //private void AttachProcess(string processName)
        //{
        //    var process = Process.GetProcessesByName(processName).FirstOrDefault();
        //    if (process == null) throw new Exception($"找不到进程: {processName}");

        //    _app = FlaUIApp.Attach(process);
        //    _window = Retry.WhileNull(() => _app.GetMainWindow(_automation), TimeSpan.FromSeconds(3)).Result;

        //    if (_window == null) throw new Exception("无法获取主窗口");
        //    _window.SetForeground();
        //}

        private void AttachToProcess(string processName)
        {
            // =============================================================
            // 1. 安全清理旧连接
            // =============================================================
            if (_application != null)
            {
                try { _application.Dispose(); } catch { }
                _application = null;
                _window = null;
            }

            // =============================================================
            // 2. 查找新进程
            // =============================================================
            var targetProc = System.Diagnostics.Process.GetProcessesByName(processName).FirstOrDefault();
            if (targetProc == null)
            {
                throw new Exception($"未找到进程: {processName}。请确认程序已启动。");
            }

            int procId = -1;
            try { procId = targetProc.Id; } catch { }

            if (targetProc.HasExited)
            {
                throw new Exception($"进程 {processName} 刚启动就退出了，可能是启动器进程。");
            }

            // =============================================================
            // 3. 建立连接
            // =============================================================
            try
            {
                _application = FlaUI.Core.Application.Attach(targetProc);
            }
            catch (Exception ex)
            {
                throw new Exception($"无法附加到进程 {processName} (PID: {procId}): {ex.Message}");
            }

            // =============================================================
            // 4. 获取窗口 (初次尝试)
            // =============================================================
            _window = FlaUI.Core.Tools.Retry.WhileNull(() =>
            {
                try
                {
                    if (targetProc.HasExited) return null;
                    return _application.GetMainWindow(_automation);
                }
                catch { return null; }
            }, TimeSpan.FromSeconds(2)).Result; // 稍微缩短初次等待时间

            // =============================================================
            // 5. 【新增】托盘/最小化唤醒逻辑 (针对微信、QQ等)
            // =============================================================
            // 判断条件：窗口为空，或者窗口虽然拿到了但宽/高为0，或者处于Offscreen状态
            bool isWindowInvalid = (_window == null);

            if (!isWindowInvalid)
            {
                // 如果窗口不为 null，进一步检查是否是一个无效的隐藏窗口
                try
                {
                    // 使用安全方法检查 Offscreen
                    if (_window.BoundingRectangle.Width <= 0 || IsOffscreenSafe(_window))
                    {
                        isWindowInvalid = true;
                    }
                }
                catch { isWindowInvalid = true; }
            }

            if (isWindowInvalid)
            {
                Log($"[系统] 窗口状态异常 (Null/隐藏/托盘)，尝试唤醒进程: {processName}...");

                // 尝试通过再次运行 exe 来唤醒主窗口 (适用于微信、QQ等单实例应用)
                try
                {
                    if (targetProc.MainModule != null)
                    {
                        string exePath = targetProc.MainModule.FileName;
                        System.Diagnostics.Process.Start(exePath);

                        // 给它一点时间弹出窗口
                        System.Threading.Thread.Sleep(1500);

                        // 再次尝试获取
                        _window = _application.GetMainWindow(_automation);
                    }
                }
                catch (Exception ex)
                {
                    Log($"[警告] 尝试唤醒失败: {ex.Message}");
                }
            }

            // =============================================================
            // 6. 最终检查
            // =============================================================
            if (_window == null)
            {
                if (targetProc.HasExited)
                    throw new Exception("连接过程中目标进程已退出。");
                else
                    throw new Exception("连接到了进程，但无法获取主窗口（可能是最小化到了托盘，且唤醒失败）。");
            }

            // 记录名字方便自动重连
            _lastProcessName = processName;

            // 尝试置顶
            EnsureWindowReady();

            Log($"[系统] 已连接到进程: {processName} (PID: {procId})");
        }

        // 强大的元素查找器：支持 Name:xx, Id:xx, Index:ComboBox:0
        // ============================================================
        // 修改后的 FindElement 方法 (已合并 Type 到 Index)
        // ============================================================
        private AutomationElement FindElement(string targetStr)
        {
            // 1. 【新增】每次找控件前，先确保窗口是正常的
            EnsureWindowReady();

            if (_window == null) throw new Exception("未连接窗口");

            var parts = targetStr.Split(new[] { ':' }, 3);
            if (parts.Length < 2) throw new Exception($"定位符格式错误: {targetStr}");

            string locatorType = parts[0].ToLower();

            return FlaUI.Core.Tools.Retry.WhileNull(() =>
            {
                try
                {
                    switch (locatorType)
                    {
                        case "id":
                            return _window.FindFirstDescendant(cf => cf.ByAutomationId(parts[1]));

                        case "name":
                            // 解析参数: Name:类型:名称
                            // 例如: Name:Custom:Project
                            string nameValue = "";
                            FlaUI.Core.Definitions.ControlType? targetType = null;

                            if (parts.Length == 3)
                            {
                                // 脚本指定了类型
                                targetType = GetControlType(parts[1]);
                                nameValue = parts[2];
                            }
                            else
                            {
                                // 脚本没指定类型，默认为 null
                                nameValue = parts[1];
                            }

                            // === 阶段 1: 严格匹配 (按指定类型找) ===
                            AutomationElement foundElement = null;

                            if (targetType.HasValue)
                            {
                                var strictCandidates = _window.FindAllDescendants(cf => cf.ByControlType(targetType.Value));
                                foreach (var item in strictCandidates)
                                {
                                    if (item.Name?.IndexOf(nameValue, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        foundElement = item;
                                        break;
                                    }
                                }
                            }

                            // === 阶段 2: 降级匹配 (如果严格匹配没找到，尝试在所有控件里找) ===
                            // 很多时候 Spy 识别为 Custom，但实际上它是 Text 或 MenuItem，这里做兜底
                            if (foundElement == null)
                            {
                                // 如果阶段1没找到，或者脚本本来就没指定类型，就全量搜
                                // 记录个日志方便调试 (可选)
                                // if (targetType.HasValue) Log($"[兼容查找] 按类型 {targetType} 未找到 '{nameValue}'，正在全屏搜索...");

                                var allCandidates = _window.FindAllDescendants();
                                foreach (var item in allCandidates)
                                {
                                    if (item.Name?.IndexOf(nameValue, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        foundElement = item;
                                        break;
                                    }
                                }
                            }

                            return foundElement;

                        case "index":
                            // Index:Button:0
                            if (parts.Length < 3) throw new Exception($"Index 格式错误");

                            string controlTypeStr = parts[1];
                            int idx = int.Parse(parts[2]);
                            var controlType = GetControlType(controlTypeStr);

                            // 1. 获取所有该类型的控件 (树形顺序)
                            var elements = _window.FindAllDescendants(cf => cf.ByControlType(controlType));

                            // =========================================================
                            // 【核心升级】 使用通用过滤器清洗数据
                            //  自动剔除 ComboBox、滚动条、标题栏内部的 Button 和 Edit
                            // =========================================================
                            elements = elements.Where(e => !IsInternalStructure(e)).ToArray();
                            // =========================================================

                            if (elements.Length <= idx) return null;

                            return elements[idx];

                        default:
                            throw new Exception($"未知的定位方式: {parts[0]}");
                    }
                }
                catch (Exception ex)
                {
                    // 【新增】如果是拒绝访问，说明 COM 连接断了，尝试重新 Attach 一次
                    if (ex.Message.Contains("拒绝访问") || ex.Message.Contains("Access Denied"))
                    {
                        Log("[严重] COM 连接失效，正在尝试重新 ATTACH...");
                        AttachToProcess(_lastProcessName);
                        // 重新 Attach 后，_window 变了，下次循环 retry 就会用新的 window
                        return null;
                    }
                    return null;
                }
            }, TimeSpan.FromSeconds(2)).Result ?? throw new Exception($"找不到控件: {targetStr}");
        }
        /// <summary>
        /// 判断一个控件是否属于某个复合控件的内部结构（噪音）
        /// </summary>
        private bool IsInternalStructure(FlaUI.Core.AutomationElements.AutomationElement element)
        {
            var parent = element.Parent;
            if (parent == null) return false; // 没有父级肯定不是内部结构

            // 获取父级类型
            var parentType = parent.ControlType;

            // --- 针对 Button 的过滤规则 ---
            if (element.ControlType == FlaUI.Core.Definitions.ControlType.Button)
            {
                // 1. 过滤下拉框里的箭头按钮
                if (parentType == FlaUI.Core.Definitions.ControlType.ComboBox) return true;

                // 2. 过滤数字微调框(Spinner)里的上下箭头
                if (parentType == FlaUI.Core.Definitions.ControlType.Spinner) return true;

                // 3. 过滤滚动条两端的箭头
                if (parentType == FlaUI.Core.Definitions.ControlType.ScrollBar) return true;

                // 4. 过滤标题栏里的最小化、关闭按钮
                if (parentType == FlaUI.Core.Definitions.ControlType.TitleBar) return true;

                // 5. 过滤 SplitButton 内部的下拉触发器 (视情况而定，通常也是杂音)
                if (parentType == FlaUI.Core.Definitions.ControlType.SplitButton) return true;
            }

            // --- 针对 Edit 的过滤规则 ---
            if (element.ControlType == FlaUI.Core.Definitions.ControlType.Edit)
            {
                // 1. 过滤下拉框里的文本输入区
                if (parentType == FlaUI.Core.Definitions.ControlType.ComboBox) return true;

                // 2. 过滤数字控件里的输入区 (如果它被识别为 Edit)
                if (parentType == FlaUI.Core.Definitions.ControlType.Spinner) return true;
            }

            // 其他情况都视为正常控件
            return false;
        }
        // ============================================================
        // 完善后的 GetControlType (支持更多别名)
        // ============================================================
        private ControlType GetControlType(string typeName)
        {
            // 移除空白并转小写，防止格式误差
            string cleanName = typeName.Trim().ToLower();

            switch (cleanName)
            {
                // === 基础输入 ===
                case "edit":
                case "textbox":
                case "text":
                case "document": // 新增：常见的编辑器类型（如 Notepad++, Word）
                    return ControlType.Edit;

                // === 点击类 ===
                case "button":
                case "btn":
                    return ControlType.Button;
                case "hyperlink": // 新增：超链接
                case "link":
                    return ControlType.Hyperlink;
                case "menuitem":  // 新增：菜单项
                    return ControlType.MenuItem;
                case "tabitem":   // 新增：选项卡
                    return ControlType.TabItem;
                case "image":     // 新增：图片（有时当作按钮用）
                    return ControlType.Image;
                case "pane":      // 新增：面板（有时候整个区域可点）
                    return ControlType.Pane;
                case "group":     // 新增：组
                    return ControlType.Group;
                case "custom":    // 新增：自定义控件
                    return ControlType.Custom;

                // === 选择类 ===
                case "combobox":
                case "combo":
                    return ControlType.ComboBox;
                case "list":
                case "listbox":
                    return ControlType.List;
                case "datagrid":  // 新增：数据表格
                    return ControlType.DataGrid;
                case "checkbox":
                case "check":
                    return ControlType.CheckBox;
                case "radiobutton": // 新增：单选框
                case "radio":
                    return ControlType.RadioButton;
                case "tree":      // 新增：树形菜单
                    return ControlType.Tree;
                case "treeitem":
                    return ControlType.TreeItem;

                // === 兜底 ===
                default:
                    // 如果脚本里写了无法识别的类型，直接尝试用 Custom，不要报错
                    // 这样能最大程度兼容 Spy 抓回来的奇怪类型
                    return ControlType.Custom;
            }
        }
        private void PerformClick(FlaUI.Core.AutomationElements.AutomationElement element)
        {
            try
            {
                // === 1. 标准 Invoke (按钮) ===
                if (element.Patterns.Invoke.IsSupported)
                {
                    element.Patterns.Invoke.Pattern.Invoke();
                    return;
                }

                // === 2. Toggle (复选/开关) ===
                if (element.Patterns.Toggle.IsSupported)
                {
                    element.Patterns.Toggle.Pattern.Toggle();
                    return;
                }

                // === 3. SelectionItem (列表项/Tab) ===
                if (element.Patterns.SelectionItem.IsSupported)
                {
                    element.Patterns.SelectionItem.Pattern.Select();
                    return;
                }

                // === 4. ExpandCollapse (折叠菜单) ===
                if (element.Patterns.ExpandCollapse.IsSupported)
                {
                    var pattern = element.Patterns.ExpandCollapse.Pattern;
                    if (pattern.ExpandCollapseState.Value == FlaUI.Core.Definitions.ExpandCollapseState.Collapsed)
                        pattern.Expand();
                    else
                        pattern.Collapse(); // 如果你只想展开不想折叠，可以把这就注释掉
                    return;
                }

                // === 5. 【关键新增】LegacyIAccessible (菜单项/老旧控件专用) ===
                // 很多菜单项识别为 Custom，且没有 Invoke 模式，但支持这个“默认动作”
                if (element.Patterns.LegacyIAccessible.IsSupported)
                {
                    var pattern = element.Patterns.LegacyIAccessible.Pattern;
                    // 执行“默认动作”，通常就是点击
                    pattern.DoDefaultAction();
                    Log("[提示] 使用 LegacyIAccessible 模式点击成功");
                    return;
                }

                // === 6. 最后的兜底：鼠标物理点击 ===
                if (element.TryGetClickablePoint(out var point))
                {
                    Log("[提示] 控件无自动化模式，降级为鼠标物理点击...");
                    FlaUI.Core.Input.Mouse.Click(point);
                }
                else
                {
                    // 如果连坐标都拿不到（比如菜单没展开），尝试聚焦后回车
                    Log("[警告] 无法获取坐标，尝试 Focus + Enter...");
                    element.Focus();
                    System.Threading.Thread.Sleep(50);
                    FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.ENTER);
                }
            }
            catch (Exception ex)
            {
                // 捕获点击过程中的错误，防止脚本直接崩掉，允许上层决定是否忽略
                throw new Exception($"点击操作失败: {ex.Message}");
            }
        }
        // 【新增】安全获取 IsOffscreen 属性
        private bool IsOffscreenSafe(FlaUI.Core.AutomationElements.AutomationElement element)
        {
            if (element == null) return true; // 元素都没了，肯定算Offscreen
            try
            {
                return element.IsOffscreen;
            }
            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException)
            {
                // 如果控件不支持这个属性，通常说明它是那种很老或者很奇怪的控件
                // 我们默认它是在屏幕上的 (false)，以免阻碍后续操作
                return false;
            }
            catch (Exception)
            {
                // 其他错误也默认可见
                return false;
            }
        }
        private void PerformInput(FlaUI.Core.AutomationElements.AutomationElement element, string text)
        {
            // 1. 核心动作：先聚焦。对于 Excel，不聚焦几乎无法操作。
            try { element.Focus(); } catch { }
            System.Threading.Thread.Sleep(50);

            bool handled = false;

            // 2. 尝试标准赋值 (ValuePattern)
            // Excel 的坑：它有 ValuePattern，但 SetValue 经常报错。
            // 所以这里必须加 try-catch，一旦报错立即降级为键盘模拟。
            if (element.Patterns.Value.IsSupported)
            {
                try
                {
                    var valPattern = element.Patterns.Value.Pattern;
                    // 只有当它是普通输入框，且不只读时，才尝试 SetValue
                    if (!valPattern.IsReadOnly.Value)
                    {
                        valPattern.SetValue(text);
                        handled = true;
                    }
                }
                catch (Exception ex)
                {
                    Log($"[Excel兼容] 标准赋值失败 ({ex.Message})，转为键盘模拟...");
                    // handled = false; // 标记为未处理，让下面继续执行键盘模拟
                }
            }

            // 3. 如果标准赋值没成功，或者报错了，执行“键盘模拟”
            if (!handled)
            {
                // 特殊处理：如果是 Excel 单元格，通常需要双击进入编辑模式，或者直接打字覆盖
                // 这里我们简单粗暴：直接模拟打字

                // 既然上面已经 Focus 了，这里直接 Type
                FlaUI.Core.Input.Keyboard.Type(text);

                // 【关键】Excel 输入完通常需要回车确认，否则脚本继续跑，焦点还在编辑框里
                // 判断一下，如果是 Excel 进程，自动补一个回车（可选，看你需求）
                // FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.ENTER);

                Log($"已通过键盘模拟输入: {text}");
            }
        }

        // ============================================================
        // 1. 修改入口方法：根据控件类型分流
        // ============================================================
        private void HandleSelect(string targetStr, string valueOrIndex)
        {
            // 先找到基础元素
            var element = FindElement(targetStr);

            // 判断它是 ComboBox 还是 ListBox
            if (element.ControlType == FlaUI.Core.Definitions.ControlType.List)
            {
                // 如果是列表框，走列表专用逻辑
                HandleListBox(element.AsListBox(), valueOrIndex);
            }
            else
            {
                // 默认为下拉框，走之前的复杂逻辑
                HandleComboBox(element.AsComboBox(), valueOrIndex);
            }
        }

        // ============================================================
        // 2. 新增：专门处理 ListBox (列表框) 的方法
        // ============================================================
        private void HandleListBox(FlaUI.Core.AutomationElements.ListBox listBox, string valueOrIndex)
        {
            try
            {
                CheckCancellation(); // Check 1
                Log($"正在操作列表框: {listBox.Name ?? "无名List"}");

                bool itemsLoaded = FlaUI.Core.Tools.Retry.WhileFalse(() =>
                {
                    if (_isCancellationRequested) return true; // Check 2
                    try { return listBox.Items != null && listBox.Items.Length > 0; }
                    catch { return false; }
                }, TimeSpan.FromSeconds(3)).Result;

                CheckCancellation(); // Check 3

                if (!itemsLoaded) throw new Exception("列表框内容为空或加载超时");

                int index = -1;
                bool isIndex = int.TryParse(valueOrIndex, out index);

                if (isIndex)
                {
                    // 【核心修复】
                    if (listBox.Items.Length == 0 && _isCancellationRequested) CheckCancellation();

                    if (index < 0 || index >= listBox.Items.Length)
                    {
                        throw new IndexOutOfRangeException($"列表索引越界：总数 {listBox.Items.Length}, 目标 {index}");
                    }

                    var item = listBox.Items[index];
                    item.ScrollIntoView();
                    item.Select();
                }
                else
                {
                    var item = listBox.Items.FirstOrDefault(x => x.Name != null && x.Name.Contains(valueOrIndex));
                    if (item != null)
                    {
                        item.ScrollIntoView();
                        item.Select();
                    }
                    else
                    {
                        CheckCancellation();
                        throw new Exception($"列表框中找不到文本: {valueOrIndex}");
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex is OperationCanceledException) throw;

                if (ex is IndexOutOfRangeException || ex.Message.Contains("越界") || ex.Message.Contains("找不到"))
                {
                    if (_isCancellationRequested) throw new OperationCanceledException("用户手动停止");
                    throw;
                }

                CheckCancellation();
                Log($"[警告] ListBox 标准操作失败 ({ex.Message})，尝试键盘降级...");
                SimulateSelectWithKeyboard(listBox, int.TryParse(valueOrIndex, out int i) ? i : 0);
            }
        }

        // ============================================================
        // 3. 原有的 ComboBox 逻辑 (改名为 HandleComboBox)
        //    (逻辑保持上一版修复后的状态，只改方法名)
        // ============================================================
        private void HandleComboBox(FlaUI.Core.AutomationElements.ComboBox combo, string valueOrIndex)
        {
            try
            {
                CheckCancellation(); // 1. 进来先检查

                // 尝试展开
                if (combo.Patterns.ExpandCollapse.IsSupported &&
                    combo.Patterns.ExpandCollapse.Pattern.ExpandCollapseState.Value != FlaUI.Core.Definitions.ExpandCollapseState.Expanded)
                {
                    combo.Patterns.ExpandCollapse.Pattern.Expand();
                    Thread.Sleep(500);
                }

                // 等待子元素 (在重试循环里也要检查停止)
                bool childrenLoaded = FlaUI.Core.Tools.Retry.WhileFalse(() =>
                {
                    if (_isCancellationRequested) return true; // 2. 如果停止，骗过重试循环让它立刻结束

                    try { var children = combo.FindAllChildren(); return children != null && children.Length > 0; }
                    catch { return false; }
                }, TimeSpan.FromSeconds(2)).Result;

                CheckCancellation(); // 3. 循环结束后立刻检查

                if (!childrenLoaded) throw new Exception("下拉框未加载到选项");

                var items = combo.Items;
                if (items == null) throw new Exception("Items 集合为 null");

                int index = -1;
                bool isIndex = int.TryParse(valueOrIndex, out index);

                if (isIndex)
                {
                    // =========================================================
                    // 【核心修复】 在抛出越界错误前，先确认是不是因为用户停止导致的
                    // =========================================================
                    if (items.Length == 0 && _isCancellationRequested)
                    {
                        CheckCancellation(); // 这里会抛出 OperationCanceledException
                    }
                    // =========================================================

                    if (index < 0 || index >= items.Length)
                    {
                        throw new IndexOutOfRangeException($"下拉框索引越界：总数 {items.Length}, 目标 {index}");
                    }
                    items[index].Select();
                }
                else
                {
                    var item = items.FirstOrDefault(x => x.Name != null && x.Name.Contains(valueOrIndex));
                    if (item != null) item.Select();
                    else
                    {
                        CheckCancellation(); // 同样，抛出找不到前先检查
                        throw new Exception($"找不到下拉项: {valueOrIndex}");
                    }
                }

                if (combo.Patterns.ExpandCollapse.IsSupported) combo.Collapse();
            }
            catch (Exception ex)
            {
                // 如果是我们自己抛出的停止异常，直接往上抛，不要降级到键盘模拟
                if (ex is OperationCanceledException) throw;

                if (ex is IndexOutOfRangeException || ex.Message.Contains("越界") || ex.Message.Contains("找不到"))
                {
                    // 如果出错了，但同时发现用户按了停止，那优先算作停止，而不是错误
                    if (_isCancellationRequested) throw new OperationCanceledException("用户手动停止");

                    throw;
                }

                CheckCancellation(); // 键盘模拟前最后检查一次

                Log($"[警告] 标准操作受阻 ({ex.Message})，切换键盘盲操作...");
                SimulateSelectWithKeyboard(combo, int.TryParse(valueOrIndex, out int i) ? i : 0);
            }
        }

        // 【新增】键盘模拟辅助方法
        // 【修改点】参数类型从 ComboBox 改为 AutomationElement
        // 这样 ListBox 和 ComboBox 都能传进来了
        private void SimulateSelectWithKeyboard(FlaUI.Core.AutomationElements.AutomationElement element, int index)
        {
            try
            {
                // 1. 聚焦控件 (这是父类 AutomationElement 就有的方法)
                element.Focus();
                Thread.Sleep(200);

                // 2. 先按 Home 键，确保回到第一个选项
                FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.HOME);
                Thread.Sleep(50);

                // 3. 按 Down 键移动到指定索引
                if (index > 0)
                {
                    for (int i = 0; i < index; i++)
                    {
                        FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.DOWN);
                        Thread.Sleep(30); // 稍微停顿防止按太快
                    }
                }

                // 4. 按 Enter 确认 (对于 ListBox 其实按不按回车通常都行，但按了也没坏处)
                FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.ENTER);
                Thread.Sleep(200); // 等待生效

                Log($"已通过键盘模拟选择了第 {index} 项");
            }
            catch (Exception keyEx)
            {
                // 既然是盲操作，报错了也没办法处理了，只能记录
                Log($"键盘模拟操作发生错误: {keyEx.Message}");
            }
        }

        // ============================================================
        // 将 AutomationEngine 类中的 CheckText 方法替换为以下内容
        // ============================================================
        private void CheckText(string targetStr, string expectedText)
        {
            var element = FindElement(targetStr);
            string lastSeenText = "";

            // 1. 延长超时时间到 10 秒 (网络请求通常较慢)
            var success = FlaUI.Core.Tools.Retry.WhileFalse(() =>
            {
                try
                {
                    // 获取当前文本
                    lastSeenText = GetTextSafe(element);

                    // 打印调试日志 (可选，调试时打开)
                    // Debug.WriteLine($"正在检查: [{lastSeenText}] vs [{expectedText}]");

                    return lastSeenText.Contains(expectedText);
                }
                catch
                {
                    return false;
                }
            }, TimeSpan.FromSeconds(10)).Result;

            // 2. 如果失败，抛出带有详细信息的异常
            if (!success)
            {
                throw new Exception($"验证失败！\r\n目标控件: {targetStr}\r\n预期包含: '{expectedText}'\r\n实际内容: '{lastSeenText}'\r\n(请检查是否定位错了控件，或者文本稍有不同)");
                //throw new Exception($"验证失败！\r\n目标控件: {targetStr}\r\n预期包含: '{expectedText}'");
            }

            Log($"验证通过: 控件包含了 '{expectedText}'");
        }

        // ============================================================
        // 【新增】万能文本获取方法 (兼容 Label, TextBox, Document)
        // ============================================================
        private string GetTextSafe(FlaUI.Core.AutomationElements.AutomationElement element)
        {
            if (element == null) return "";

            // 策略 1: 如果是 TextBox/Edit，优先读 Text 属性 (ValuePattern)
            if (element.ControlType == FlaUI.Core.Definitions.ControlType.Edit ||
                element.ControlType == FlaUI.Core.Definitions.ControlType.Document)
            {
                var textPattern = element.Patterns.Value.PatternOrDefault;
                if (textPattern != null)
                {
                    return textPattern.Value.Value;
                }
                // 有些富文本框用的是 TextPattern
                var docPattern = element.Patterns.Text.PatternOrDefault;
                if (docPattern != null)
                {
                    return docPattern.DocumentRange.GetText(-1);
                }
            }

            // 策略 2: 对于 Label/Text，通常只有 Name 属性
            if (!string.IsNullOrEmpty(element.Name))
            {
                return element.Name;
            }

            // 策略 3: 最后的兜底，尝试读 HelpText 或 ItemStatus
            return element.HelpText ?? "";
        }
    }
}
