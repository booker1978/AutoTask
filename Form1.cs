using System;
using System.Diagnostics;
using System.IO;

namespace ToolsAutoTask
{
    public partial class Form1 : Form
    {
        /*
         * ToolsAutoTask v1.0.2025.1229 
         * 利用脚本自动驱动应用程序，完成对各类控件的操控（要求目标应用已经运行）
         * 目前支持对应用程序中的文本框、列表框、下拉框、按钮等控件进行选择、检验内容、输入内容等操作
         * 支持设定循环，支持在操作中间设定间隔（不锁死系统），支持忽略操作错误
         * 仅支持屏幕关闭但系统未休眠/未锁定的场景；一旦系统休眠、锁定或远程会话断开，任务将失败
         * 
         * ToolsAutoTask v1.1.2025.1230 
         * 修复：遇到错误指令，进行提示
         * 新增：支持鼠标操作脚本的识别（单击、双击、滚轮等）
         * 新增：间谍模式，可以识别目标程序的名称、控件Id、坐标位置等信息
         * 新增：支持间谍模式下的脚本长录模式，对于识别到的模糊控件提供鼠标模拟操作建议
         * 
         */

        private AutomationEngine _engine;
        private ToolsAutoTask.ElementSpy _spy;
        private bool _isRunning = false; // 标记当前是否正在运行侦测
        private string _lastAttachedProcess = "";

        public Form1()
        {
            InitializeComponent();
            _engine = new AutomationEngine();
            _spy = new ToolsAutoTask.ElementSpy();

            // 订阅日志事件，更新到界面
            _engine.OnLog += msg => this.Invoke((MethodInvoker)(() => textBox2.AppendText(msg + "\r\n")));
        }

        /// <summary>
        /// 执行自动化任务
        /// </summary>
        /// <returns>返回执行结果字符串，成功或具体的失败步骤</returns>
        private async void button1_Click(object sender, EventArgs e)
        {
            string script = txtScript.Text;
            if (string.IsNullOrWhiteSpace(script)) return;

            // 1. 切换按钮状态：禁用运行，启用停止
            button1.Enabled = false;
            button5.Enabled = true;

            // 传递设置
            _engine.IgnoreError = checkBox1.Checked;

            // 2. 异步执行脚本 (防止卡死界面)
            await Task.Run(() =>
            {
                _engine.ExecuteScript(script);
            });

            // 3. 执行结束后(不管是跑完还是被停止)，恢复按钮状态
            // 注意：Invoke用于确保在UI线程操作
            this.Invoke((Action)(() =>
            {
                button1.Enabled = true;
                button5.Enabled = false;
            }));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            txtScript.Text = "";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 1. 创建文件选择对话框
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // 2. 设置默认目录为程序当前运行的目录
            // AppDomain.CurrentDomain.BaseDirectory 获取的是 .exe 所在的文件夹
            openFileDialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;

            // 3. 设置文件过滤器 (只显示 txt 文件和所有文件)
            openFileDialog.Filter = "脚本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*";

            // 4. 设置对话框标题
            openFileDialog.Title = "请选择指令脚本文件";

            // 5. 显示对话框并判断用户是否点击了“确定”
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // 6. 获取选中的文件路径
                    string filePath = openFileDialog.FileName;

                    // 7. 读取文件内容
                    // 注意：默认使用 UTF-8 编码读取。
                    // 如果你的文件是记事本默认保存的 (GB2312/ANSI)，出现乱码的话，
                    // 请把下面这行改成: File.ReadAllText(filePath, System.Text.Encoding.Default);
                    string fileContent = File.ReadAllText(filePath);

                    // 8. 填充到任务文本框 (textbox1)
                    txtScript.Text = fileContent;

                    // (可选) 在日志里提示一下
                    // AppendLog($"已加载脚本: {filePath}"); 
                }
                catch (Exception ex)
                {
                    MessageBox.Show("读取文件失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            // 调用引擎的刹车方法
            _engine.Stop();

            // 界面反馈
            // (注意：这里不要直接把 btnRun 设为 true，因为后台线程还没完全退出来，
            //  最好等 Task.Run 彻底结束回到 btnRun_Click 的最后一步再恢复)
            button5.Enabled = false; // 防止用户狂点
        }

        private void btnSpy_Click(object sender, EventArgs e)
        {
            if (!_isRunning)
            {
                // === [启动阶段] ===
                _isRunning = true;

                // 【关键】每次重新开始录制时，重置进程记录，确保第一行脚本包含 ATTACH
                // 如果你希望在“暂停”后续录时不重复生成，可以将这行删掉
                _lastAttachedProcess = "";

                chkRecordMode.Enabled = false;

                // 界面样式更新
                if (chkRecordMode.Checked)
                {
                    btnSpy.Text = "Spy Stop";
                    btnSpy.BackColor = Color.LightCoral;
                    lblStatus.Text = "状态: 录制中... (按 Ctrl + 点击)";
                }
                else
                {
                    btnSpy.Text = "Spy Stop";
                    btnSpy.BackColor = Color.LightGreen;
                    lblStatus.Text = "状态: 侦测中... (按 Ctrl + 点击)";
                }

                this.TopMost = true;

                // 启动监听
                _spy.StartRecording((result) =>
                {
                    this.Invoke(new Action(() =>
                    {
                        // 1. Spy 功能：始终更新信息显示
                        if (txtSelector != null) txtSelector.Text = result.Selector;
                        if (txtCoordinate != null) txtCoordinate.Text = result.RelativeCoord;
                        lblStatus.Text = $"已识别: {result.Selector}";

                        // 2. 录制功能：写入脚本
                        if (chkRecordMode.Checked)
                        {
                            // ====================================================
                            // 【核心逻辑】智能 ATTACH 判断
                            // ====================================================

                            // 只有当 1. 抓取到了进程名 且 2. 进程名与上次不同 时，才写入 ATTACH
                            if (!string.IsNullOrEmpty(result.ProcessName) &&
                                result.ProcessName != _lastAttachedProcess)
                            {
                                txtScript.AppendText($"ATTACH {result.ProcessName}\r\n");

                                // 更新记录，下次如果是同一个进程就不写了
                                _lastAttachedProcess = result.ProcessName;
                            }

                            // 写入实际的操作指令
                            txtScript.AppendText(result.ScriptLine + "\r\n");
                            txtScript.ScrollToCaret();
                        }
                    }));
                });
            }
            else
            {
                // === [停止阶段] ===
                _spy.Stop();
                _isRunning = false;

                btnSpy.Text = "Spy...";
                btnSpy.BackColor = SystemColors.Control;
                chkRecordMode.Enabled = true;
                this.TopMost = false;
                lblStatus.Text = "状态: 就绪";
            }
        }

       
    }
}
