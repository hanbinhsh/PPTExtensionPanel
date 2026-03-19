using System;

namespace PPTExtensionPanel
{
    partial class ToolSelectionDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnSave = new System.Windows.Forms.Button();
            this.chkListTools = new System.Windows.Forms.CheckedListBox();
            this.SuspendLayout();
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(674, 522);
            this.btnSave.MaximumSize = new System.Drawing.Size(100, 30);
            this.btnSave.MinimumSize = new System.Drawing.Size(100, 30);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(100, 30);
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "保存";
            this.btnSave.UseVisualStyleBackColor = true;
            // 
            // chkListTools
            // 
            this.chkListTools.FormattingEnabled = true;
            this.chkListTools.Location = new System.Drawing.Point(12, 12);
            this.chkListTools.Name = "chkListTools";
            this.chkListTools.Size = new System.Drawing.Size(762, 504);
            this.chkListTools.TabIndex = 2;
            // 
            // ToolSelectionDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(786, 555);
            this.Controls.Add(this.chkListTools);
            this.Controls.Add(this.btnSave);
            this.Name = "ToolSelectionDialog";
            this.Text = "ToolSelectionDialog";
            this.Load += new System.EventHandler(this.ToolSelectionDialog_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.CheckedListBox chkListTools;

        private void ToolSelectionDialog_Load(object sender, EventArgs e)
        {
        }
    }
}