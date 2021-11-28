using System.Windows.Forms;

namespace ExcelConverter
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.singleMode = new System.Windows.Forms.RadioButton();
            this.batchmode = new System.Windows.Forms.RadioButton();
            this.startBtn = new System.Windows.Forms.Button();
            this.openPathBtn = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.modeSelection = new System.Windows.Forms.ComboBox();
            this.pathTextBox = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.singleMode);
            this.panel1.Controls.Add(this.batchmode);
            this.panel1.Controls.Add(this.startBtn);
            this.panel1.Controls.Add(this.openPathBtn);
            this.panel1.Controls.Add(this.progressBar1);
            this.panel1.Controls.Add(this.modeSelection);
            this.panel1.Controls.Add(this.pathTextBox);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1295, 291);
            this.panel1.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(85, 145);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(138, 25);
            this.label1.TabIndex = 9;
            this.label1.Text = "选择目标格式";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // singleMode
            // 
            this.singleMode.AutoSize = true;
            this.singleMode.Location = new System.Drawing.Point(841, 141);
            this.singleMode.Name = "singleMode";
            this.singleMode.Size = new System.Drawing.Size(148, 29);
            this.singleMode.TabIndex = 8;
            this.singleMode.Text = "单文件转换";
            this.singleMode.UseVisualStyleBackColor = true;
            // 
            // batchmode
            // 
            this.batchmode.AutoSize = true;
            this.batchmode.Checked = true;
            this.batchmode.Location = new System.Drawing.Point(662, 141);
            this.batchmode.Name = "batchmode";
            this.batchmode.Size = new System.Drawing.Size(127, 29);
            this.batchmode.TabIndex = 7;
            this.batchmode.TabStop = true;
            this.batchmode.Text = "批量转换";
            this.batchmode.UseVisualStyleBackColor = true;
            this.batchmode.CheckedChanged += new System.EventHandler(this.batchmode_CheckedChanged);
            // 
            // startBtn
            // 
            this.startBtn.Location = new System.Drawing.Point(1046, 220);
            this.startBtn.Name = "startBtn";
            this.startBtn.Size = new System.Drawing.Size(218, 47);
            this.startBtn.TabIndex = 6;
            this.startBtn.Text = "开始执行";
            this.startBtn.UseVisualStyleBackColor = true;
            this.startBtn.Click += new System.EventHandler(this.startBtn_Click);
            // 
            // openPathBtn
            // 
            this.openPathBtn.Location = new System.Drawing.Point(43, 43);
            this.openPathBtn.Name = "openPathBtn";
            this.openPathBtn.Size = new System.Drawing.Size(218, 48);
            this.openPathBtn.TabIndex = 4;
            this.openPathBtn.Text = "选择文件";
            this.openPathBtn.UseVisualStyleBackColor = true;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(43, 220);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(946, 47);
            this.progressBar1.TabIndex = 3;
            // 
            // modeSelection
            // 
            this.modeSelection.FormattingEnabled = true;
            this.modeSelection.Location = new System.Drawing.Point(318, 140);
            this.modeSelection.Name = "modeSelection";
            this.modeSelection.Size = new System.Drawing.Size(218, 33);
            this.modeSelection.TabIndex = 2;
            this.modeSelection.SelectedIndexChanged += new System.EventHandler(this.modeSelection_SelectedIndexChanged);
            // 
            // pathTextBox
            // 
            this.pathTextBox.Location = new System.Drawing.Point(318, 52);
            this.pathTextBox.Name = "pathTextBox";
            this.pathTextBox.Size = new System.Drawing.Size(946, 31);
            this.pathTextBox.TabIndex = 0;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1295, 291);
            this.Controls.Add(this.panel1);
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel Converter";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button startBtn;
        private System.Windows.Forms.Button openPathBtn;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ComboBox modeSelection;
        private System.Windows.Forms.TextBox pathTextBox;
        private System.Windows.Forms.RadioButton singleMode;
        private System.Windows.Forms.RadioButton batchmode;
        private System.Windows.Forms.Label label1;
    }
}

