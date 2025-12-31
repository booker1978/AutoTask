using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;
using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms; // 必须引用 System.Windows.Forms

namespace ToolsAutoTask
{
    public class ElementSpy
    {
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int VK_LBUTTON = 0x01; // 左键
        private const int VK_RBUTTON = 0x02; // 右键
        private const int VK_CONTROL = 0x11; // Ctrl键

        private volatile bool _isSpying = false;
        private UIA3Automation _automation;

        // === 1. 终极兼容版 SpyResult (包含所有历史属性) ===
        public class SpyResult
        {
            // 新版属性
            public string ScriptLine { get; set; }      // 最终生成的单行脚本

            // 旧版属性 (保留以兼容你的 UI 代码)
            public string SuggestedScript { get; set; } // 等同于 ScriptLine
            public string Selector { get; set; }        // 定位符 (Name:xx)
            public string RelativeCoord { get; set; }   // 坐标 (100,200)
            public string ProcessName { get; set; }     // 进程名
            public string Description { get; set; }     // 描述
            public string FullDescription { get; set; } // 详细描述文本
        }

        public ElementSpy()
        {
            _automation = new UIA3Automation();
        }

        // 兼容两个方法名：StartSpying 和 StartRecording
        // 你的代码调用哪个都不会报错
        public void StartSpying(Action<SpyResult> onRecord) => StartRecording(onRecord);

        public void StartRecording(Action<SpyResult> onRecord)
        {
            if (_isSpying) return;
            _isSpying = true;

            Task.Run(() =>
            {
                while (_isSpying)
                {
                    // 检测 Ctrl 是否按住
                    bool isCtrlDown = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;

                    if (isCtrlDown)
                    {
                        // 检测鼠标点击 (左键 或 右键)
                        bool isLeftClick = (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
                        bool isRightClick = (GetAsyncKeyState(VK_RBUTTON) & 0x8000) != 0;

                        if (isLeftClick || isRightClick)
                        {
                            try
                            {
                                var mousePoint = Cursor.Position;
                                var element = _automation.FromPoint(mousePoint);

                                if (element != null)
                                {
                                    // 高亮反馈
                                    try { element.DrawHighlight(Color.Red); } catch { }

                                    // 生成详细结果
                                    var result = AnalyzeAndGenerateScript(element, mousePoint, isRightClick);

                                    // 回调
                                    onRecord?.Invoke(result);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("录制异常: " + ex.Message);
                            }

                            // 防抖：等待鼠标松开
                            while ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0 ||
                                   (GetAsyncKeyState(VK_RBUTTON) & 0x8000) != 0)
                            {
                                Thread.Sleep(50);
                            }
                        }
                    }
                    Thread.Sleep(50);
                }
            });
        }

        public void Stop()
        {
            _isSpying = false;
        }

        // === 分析逻辑 ===
        // === 优化后的分析逻辑 ===
        private SpyResult AnalyzeAndGenerateScript(AutomationElement element, Point mousePoint, bool isRightClick)
        {
            var result = new SpyResult();

            // 1. 获取基本信息
            string autoId = GetSafeProperty(() => element.AutomationId);
            string name = GetSafeProperty(() => element.Name);
            var controlType = element.ControlType;
            string typeSimple = GetSimpleTypeName(controlType);

            // 优化：如果当前抓到的是 Pane/Custom 且没有名字，尝试看看它是不是包裹着一个 Text
            if ((element.ControlType == ControlType.Pane || element.ControlType == ControlType.Custom)
                && string.IsNullOrEmpty(element.Name)
                && string.IsNullOrEmpty(element.AutomationId))
            {
                // 尝试找子元素里的 Text
                var textChild = element.FindFirstDescendant(cf => cf.ByControlType(ControlType.Text));
                if (textChild != null && !string.IsNullOrEmpty(textChild.Name))
                {
                    // 偷梁换柱：把目标改成内部的 Text 元素
                    element = textChild;
                    // 重新获取一下类型
                    controlType = element.ControlType;
                    typeSimple = GetSimpleTypeName(controlType);
                }
            }

            // 2. 生成 Selector (智能精简版)
            string selector = "";

            // 规则A: 有 ID 永远优先用 ID (最短且最稳)
            if (!string.IsNullOrEmpty(autoId))
            {
                selector = $"Id:{autoId}";
            }
            // 规则B: 对于输入框，Name 往往是内容，容易导致脚本爆炸。
            // 所以如果是 Edit/Document，且 Name 很长(>20字符)，我们强制跳过 Name，改用 Index。
            else if (!string.IsNullOrEmpty(name) &&
                     !((controlType == ControlType.Edit || controlType == ControlType.Document) && name.Length > 20))
            {
                selector = $"Name:{typeSimple}:{name}";
            }
            // 规则C: 兜底使用 Index
            else
            {
                int index = CalculateIndex(element);
                selector = $"Index:{typeSimple}:{index}";
            }

            result.Selector = selector;

            // 3. 计算相对坐标
            string relCoord = "0,0";
            var root = GetRootWindow(element);
            if (root != null)
            {
                relCoord = $"{mousePoint.X - root.BoundingRectangle.X},{mousePoint.Y - root.BoundingRectangle.Y}";
                try { result.ProcessName = System.Diagnostics.Process.GetProcessById(root.Properties.ProcessId).ProcessName; } catch { }
            }

            result.RelativeCoord = relCoord;

            // 4. 生成 ScriptLine
            string action = "";

            // 定义一个“备用操作”字符串，用于追加到行尾
            // 使用 ScriptParser 支持的 MOUSE_CLICK 指令，这样用户取消注释就能直接运行
            string backupComment = $" // MOUSE_CLICK {relCoord}";

            if (isRightClick)
            {
                action = $"MOUSE_RCLICK {relCoord} // 右键";
            }
            else if (controlType == ControlType.Edit || controlType == ControlType.Document)
            {
                action = $"INPUT {selector} \"\"";
            }
            else if (controlType == ControlType.ComboBox)
            {
                action = $"SELECT {selector} \"\"";
            }
            else
            {
                // === 核心修改逻辑 ===
                bool isGoodSelector = selector.StartsWith("Id:") || selector.StartsWith("Name:");

                if (IsClickableType(controlType) && isGoodSelector)
                {
                    // 生成标准点击指令
                    action = $"CLICK {selector}";

                    // 【新增】如果 ID 看起来像是纯数字（不稳定），或者是为了保险起见，追加鼠标坐标建议
                    // 这里我们对所有 CLICK 都追加，作为一种稳健的容错
                    action += backupComment;
                }
                else
                {
                    // 如果连 ID/Name 都没有，直接生成鼠标点击
                    action = $"MOUSE_CLICK {relCoord} // {typeSimple}";
                }
            }

            // 填充属性
            result.ScriptLine = action;
            result.SuggestedScript = action;
            result.Description = $"录制: {action}";
            result.FullDescription = $"控件: {typeSimple}\r\n定位: {selector}";

            return result;
        }

        // --- 辅助方法 ---
        private string GetSafeProperty(Func<string> getter) { try { return getter(); } catch { return null; } }

        private bool IsClickableType(ControlType type)
        {
            // 如果这些类型有明确的 Name 或 ID，通常可以用 CLICK 指令，而不是非要用鼠标坐标
            return type == ControlType.Button ||
                   type == ControlType.CheckBox ||
                   type == ControlType.RadioButton ||
                   type == ControlType.Hyperlink ||
                   type == ControlType.MenuItem ||
                   type == ControlType.TabItem ||
                   type == ControlType.TreeItem ||
                   type == ControlType.ListItem ||
                   type == ControlType.Image ||
                   // === 新增以下类型 ===
                   type == ControlType.Pane ||    // 很多容器按钮是 Pane
                   type == ControlType.Text ||    // 很多标签按钮是 Text
                   type == ControlType.Group ||   // 分组框有时也是点击目标
                   type == ControlType.Custom;    // 自定义控件
        }

        private AutomationElement GetRootWindow(AutomationElement element)
        {
            var current = element;
            while (current != null)
            {
                if (current.ControlType == ControlType.Window) return current;
                current = current.Parent;
            }
            return null;
        }

        private int CalculateIndex(AutomationElement element)
        {
            var parent = element.Parent;
            if (parent == null) return 0;
            var siblings = parent.FindAllChildren(cf => cf.ByControlType(element.ControlType));
            for (int i = 0; i < siblings.Length; i++)
            {
                if (siblings[i].Properties.RuntimeId.Value.SequenceEqual(element.Properties.RuntimeId.Value)) return i;
            }
            return 0;
        }

        private string GetSimpleTypeName(ControlType type)
        {
            switch (type)
            {
                case ControlType.Edit: return "Edit";
                case ControlType.Button: return "Button";
                case ControlType.ComboBox: return "ComboBox";
                case ControlType.List: return "List";
                case ControlType.CheckBox: return "CheckBox";
                case ControlType.Document: return "Document";
                case ControlType.MenuItem: return "MenuItem";
                case ControlType.Pane: return "Pane";
                case ControlType.Image: return "Image";
                case ControlType.Group: return "Group";
                default: return "Custom";
            }
        }
    }
}