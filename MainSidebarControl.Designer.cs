namespace PPTExtensionPanel
{
    partial class MainSidebarControl
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

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnAddTool = new System.Windows.Forms.Button();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.SuspendLayout();
            // 
            // btnAddTool
            // 
            this.btnAddTool.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.btnAddTool.Location = new System.Drawing.Point(0, 967);
            this.btnAddTool.MaximumSize = new System.Drawing.Size(100, 30);
            this.btnAddTool.MinimumSize = new System.Drawing.Size(100, 30);
            this.btnAddTool.Name = "btnAddTool";
            this.btnAddTool.Size = new System.Drawing.Size(100, 30);
            this.btnAddTool.TabIndex = 0;
            this.btnAddTool.Text = "添加";
            this.btnAddTool.UseVisualStyleBackColor = true;
            this.btnAddTool.Click += new System.EventHandler(this.btnAddTool_Click_1);
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(532, 967);
            this.flowLayoutPanel1.TabIndex = 1;
            // 
            // MainSidebarControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.flowLayoutPanel1);
            this.Controls.Add(this.btnAddTool);
            this.Name = "MainSidebarControl";
            this.Size = new System.Drawing.Size(532, 997);
            this.Load += new System.EventHandler(this.MainSidebarControl_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnAddTool;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
    }
}
