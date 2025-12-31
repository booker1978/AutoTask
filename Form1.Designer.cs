namespace ToolsAutoTask
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            button1 = new Button();
            txtScript = new TextBox();
            textBox2 = new TextBox();
            splitContainer1 = new SplitContainer();
            chkRecordMode = new CheckBox();
            lblStatus = new Label();
            txtCoordinate = new TextBox();
            label2 = new Label();
            txtSelector = new TextBox();
            label1 = new Label();
            btnSpy = new Button();
            button5 = new Button();
            checkBox1 = new CheckBox();
            button4 = new Button();
            button3 = new Button();
            button2 = new Button();
            splitContainer2 = new SplitContainer();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer2).BeginInit();
            splitContainer2.Panel1.SuspendLayout();
            splitContainer2.Panel2.SuspendLayout();
            splitContainer2.SuspendLayout();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(3, 5);
            button1.Name = "button1";
            button1.Size = new Size(88, 23);
            button1.TabIndex = 0;
            button1.Text = "Run";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // txtScript
            // 
            txtScript.Dock = DockStyle.Fill;
            txtScript.Location = new Point(0, 0);
            txtScript.Multiline = true;
            txtScript.Name = "txtScript";
            txtScript.ScrollBars = ScrollBars.Both;
            txtScript.Size = new Size(344, 371);
            txtScript.TabIndex = 1;
            // 
            // textBox2
            // 
            textBox2.Dock = DockStyle.Fill;
            textBox2.Location = new Point(0, 0);
            textBox2.Multiline = true;
            textBox2.Name = "textBox2";
            textBox2.ScrollBars = ScrollBars.Both;
            textBox2.Size = new Size(452, 371);
            textBox2.TabIndex = 2;
            // 
            // splitContainer1
            // 
            splitContainer1.Dock = DockStyle.Fill;
            splitContainer1.FixedPanel = FixedPanel.Panel1;
            splitContainer1.Location = new Point(0, 0);
            splitContainer1.Name = "splitContainer1";
            splitContainer1.Orientation = Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            splitContainer1.Panel1.Controls.Add(button1);
            splitContainer1.Panel1.Controls.Add(button5);
            splitContainer1.Panel1.Controls.Add(button4);
            splitContainer1.Panel1.Controls.Add(button2);
            splitContainer1.Panel1.Controls.Add(button3);
            splitContainer1.Panel1.Controls.Add(checkBox1);
            splitContainer1.Panel1.Controls.Add(btnSpy);
            splitContainer1.Panel1.Controls.Add(chkRecordMode);
            splitContainer1.Panel1.Controls.Add(label1);
            splitContainer1.Panel1.Controls.Add(txtSelector);
            splitContainer1.Panel1.Controls.Add(label2);
            splitContainer1.Panel1.Controls.Add(txtCoordinate);
            splitContainer1.Panel1.Controls.Add(lblStatus);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(splitContainer2);
            splitContainer1.Size = new Size(800, 450);
            splitContainer1.SplitterDistance = 75;
            splitContainer1.TabIndex = 3;
            // 
            // chkRecordMode
            // 
            chkRecordMode.AutoSize = true;
            chkRecordMode.Location = new Point(96, 36);
            chkRecordMode.Name = "chkRecordMode";
            chkRecordMode.Size = new Size(50, 21);
            chkRecordMode.TabIndex = 12;
            chkRecordMode.Text = "REC";
            chkRecordMode.UseVisualStyleBackColor = true;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(448, 37);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(43, 17);
            lblStatus.TabIndex = 11;
            lblStatus.Text = "Status";
            // 
            // txtCoordinate
            // 
            txtCoordinate.Location = new Point(372, 34);
            txtCoordinate.Name = "txtCoordinate";
            txtCoordinate.Size = new Size(70, 23);
            txtCoordinate.TabIndex = 10;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(340, 37);
            label2.Name = "label2";
            label2.Size = new Size(26, 17);
            label2.TabIndex = 9;
            label2.Text = "XY:";
            // 
            // txtSelector
            // 
            txtSelector.Location = new Point(216, 34);
            txtSelector.Name = "txtSelector";
            txtSelector.Size = new Size(118, 23);
            txtSelector.TabIndex = 8;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(152, 37);
            label1.Name = "label1";
            label1.Size = new Size(58, 17);
            label1.TabIndex = 7;
            label1.Text = "Selector:";
            // 
            // btnSpy
            // 
            btnSpy.Location = new Point(3, 34);
            btnSpy.Name = "btnSpy";
            btnSpy.Size = new Size(88, 23);
            btnSpy.TabIndex = 6;
            btnSpy.Text = "Spy...";
            btnSpy.UseVisualStyleBackColor = true;
            btnSpy.Click += btnSpy_Click;
            // 
            // button5
            // 
            button5.Location = new Point(97, 5);
            button5.Name = "button5";
            button5.Size = new Size(88, 23);
            button5.TabIndex = 5;
            button5.Text = "Stop";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // checkBox1
            // 
            checkBox1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            checkBox1.AutoSize = true;
            checkBox1.Location = new Point(697, 5);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(100, 21);
            checkBox1.TabIndex = 4;
            checkBox1.Text = "Ignore Error";
            checkBox1.UseVisualStyleBackColor = true;
            // 
            // button4
            // 
            button4.Location = new Point(191, 5);
            button4.Name = "button4";
            button4.Size = new Size(88, 23);
            button4.TabIndex = 3;
            button4.Text = "Load Task";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // button3
            // 
            button3.Location = new Point(379, 5);
            button3.Name = "button3";
            button3.Size = new Size(88, 23);
            button3.TabIndex = 2;
            button3.Text = "Clear Logs";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button2
            // 
            button2.Location = new Point(285, 5);
            button2.Name = "button2";
            button2.Size = new Size(88, 23);
            button2.TabIndex = 1;
            button2.Text = "Clear Task";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // splitContainer2
            // 
            splitContainer2.Dock = DockStyle.Fill;
            splitContainer2.Location = new Point(0, 0);
            splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            splitContainer2.Panel1.Controls.Add(txtScript);
            // 
            // splitContainer2.Panel2
            // 
            splitContainer2.Panel2.Controls.Add(textBox2);
            splitContainer2.Size = new Size(800, 371);
            splitContainer2.SplitterDistance = 344;
            splitContainer2.TabIndex = 0;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(splitContainer1);
            Name = "Form1";
            Text = "AutoTask";
            splitContainer1.Panel1.ResumeLayout(false);
            splitContainer1.Panel1.PerformLayout();
            splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
            splitContainer1.ResumeLayout(false);
            splitContainer2.Panel1.ResumeLayout(false);
            splitContainer2.Panel1.PerformLayout();
            splitContainer2.Panel2.ResumeLayout(false);
            splitContainer2.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer2).EndInit();
            splitContainer2.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Button button1;
        private TextBox txtScript;
        private TextBox textBox2;
        private SplitContainer splitContainer1;
        private SplitContainer splitContainer2;
        private Button button3;
        private Button button2;
        private Button button4;
        private CheckBox checkBox1;
        private Button button5;
        private Button btnSpy;
        private Label label2;
        private TextBox txtSelector;
        private Label label1;
        private Label lblStatus;
        private TextBox txtCoordinate;
        private CheckBox chkRecordMode;
    }
}
