using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using System;
using System.Drawing; // 需要引用 System.Drawing.Primitives 或 System.Drawing.Common
using System.Threading;

namespace ToolsAutoTask
{
    /// <summary>
    /// 专门负责处理鼠标模拟操作的控制器
    /// </summary>
    public class MouseController
    {
        // 引用主引擎中的 Log 方法（可选，方便调试）
        private Action<string> _logger;

        public MouseController(Action<string> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// 处理鼠标指令的入口
        /// </summary>
        public void HandleMouseAction(ScriptParser.ScriptLine cmd, AutomationElement window)
        {
            if (window == null) throw new Exception("执行鼠标操作前必须先 ATTACH 窗口");

            switch (cmd.Action)
            {
                case ScriptParser.ActionType.MOUSE_MOVE:
                    var point = GetAbsolutePoint(window, cmd.Target);
                    Mouse.MoveTo(point);
                    _logger?.Invoke($"鼠标移动至相对坐标: {cmd.Target}");
                    break;

                case ScriptParser.ActionType.MOUSE_CLICK:
                    Click(window, cmd.Target, MouseButton.Left);
                    break;

                case ScriptParser.ActionType.MOUSE_DBLCLICK:
                    Click(window, cmd.Target, MouseButton.Left, isDoubleClick: true);
                    break;

                case ScriptParser.ActionType.MOUSE_RCLICK:
                    Click(window, cmd.Target, MouseButton.Right);
                    break;

                case ScriptParser.ActionType.MOUSE_DRAG:
                    Drag(window, cmd.Target, cmd.Value);
                    break;

                case ScriptParser.ActionType.MOUSE_SCROLL:
                    Scroll(window, cmd.Target, cmd.Value);
                    break;
            }
        }

        // === 【新增】具体的滚动实现方法 ===
        private void Scroll(AutomationElement window, string coordStr, string amountStr)
        {
            // 1. 解析坐标并移动鼠标
            // 滚轮操作前，必须先把鼠标指在目标区域上方，否则滚轮可能对错误的控件生效
            var point = GetAbsolutePoint(window, coordStr);
            Mouse.MoveTo(point);

            // 稍微等待一下，确保系统识别到 Hover 状态
            Thread.Sleep(100);

            // 2. 解析滚动数值
            if (!int.TryParse(amountStr, out int scrollAmount))
            {
                throw new Exception($"滚轮数值格式错误: {amountStr} (应为整数)");
            }

            // 3. 执行滚动
            // FlaUI 的 Mouse.Scroll 接受 double 类型
            // 正数通常是“向上/向前”滚，负数是“向下/向后”滚
            Mouse.Scroll(scrollAmount);

            _logger?.Invoke($"鼠标在 {coordStr} 处滚动了 {scrollAmount} 格");
        }
        private void Click(AutomationElement window, string coordStr, MouseButton button, bool isDoubleClick = false)
        {
            var point = GetAbsolutePoint(window, coordStr);

            // 1. 先移动过去
            Mouse.MoveTo(point);
            Thread.Sleep(50); // 稍微停顿，模拟人类

            // 2. 执行点击
            if (isDoubleClick)
            {
                Mouse.DoubleClick(button);
                _logger?.Invoke($"鼠标双击相对坐标: {coordStr}");
            }
            else
            {
                Mouse.Click(button);
                _logger?.Invoke($"鼠标{(button == MouseButton.Right ? "右" : "左")}键点击相对坐标: {coordStr}");
            }
        }

        private void Drag(AutomationElement window, string startCoordStr, string endCoordStr)
        {
            var startPoint = GetAbsolutePoint(window, startCoordStr);
            var endPoint = GetAbsolutePoint(window, endCoordStr);

            _logger?.Invoke($"开始拖拽: {startCoordStr} -> {endCoordStr}");

            // FlaUI 内置了 Drag 方法，但手动模拟更可控
            Mouse.Position = startPoint;
            Thread.Sleep(100);

            Mouse.Down(MouseButton.Left);
            Thread.Sleep(100);

            // 模拟平滑移动（可选，防反作弊或增加稳定性）
            // 如果不需要平滑，直接 Mouse.Position = endPoint 即可
            Mouse.MoveTo(endPoint);
            Thread.Sleep(200);

            Mouse.Up(MouseButton.Left);
            _logger?.Invoke("拖拽完成");
        }

        /// <summary>
        /// 将脚本中的 "x,y" 字符串解析并转换为屏幕绝对坐标
        /// </summary>
        private Point GetAbsolutePoint(AutomationElement window, string coordStr)
        {
            if (string.IsNullOrWhiteSpace(coordStr))
                throw new Exception("鼠标坐标参数为空");

            var parts = coordStr.Split(new[] { ',', '，' }); // 支持中英文逗号
            if (parts.Length != 2)
                throw new Exception($"坐标格式错误: {coordStr} (格式应为 x,y)");

            if (!int.TryParse(parts[0].Trim(), out int relX) ||
                !int.TryParse(parts[1].Trim(), out int relY))
            {
                throw new Exception($"坐标数值非法: {coordStr}");
            }

            // 获取窗口当前的屏幕位置
            // BoundingRectangle.Location 是窗口左上角在屏幕的绝对坐标
            var windowRect = window.BoundingRectangle;

            // 计算绝对坐标 = 窗口左上角 + 相对偏移
            int absX = (int)windowRect.X + relX;
            int absY = (int)windowRect.Y + relY;

            return new Point(absX, absY);
        }
    }
}
