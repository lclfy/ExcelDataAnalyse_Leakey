namespace ExcelDataAnalyse_Leakey
{
    partial class Main
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
            this.components = new System.ComponentModel.Container();
            this.GetFile_btn = new CCWin.SkinControl.SkinButton();
            this.Start_btn = new CCWin.SkinControl.SkinButton();
            this.filePathLBL = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // GetFile_btn
            // 
            this.GetFile_btn.BackColor = System.Drawing.Color.Transparent;
            this.GetFile_btn.BaseColor = System.Drawing.Color.DodgerBlue;
            this.GetFile_btn.BorderColor = System.Drawing.Color.DodgerBlue;
            this.GetFile_btn.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.GetFile_btn.DownBack = null;
            this.GetFile_btn.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.GetFile_btn.ForeColor = System.Drawing.Color.White;
            this.GetFile_btn.Location = new System.Drawing.Point(132, 298);
            this.GetFile_btn.MouseBack = null;
            this.GetFile_btn.Name = "GetFile_btn";
            this.GetFile_btn.NormlBack = null;
            this.GetFile_btn.Size = new System.Drawing.Size(195, 77);
            this.GetFile_btn.TabIndex = 0;
            this.GetFile_btn.Text = "导入所有文件";
            this.GetFile_btn.UseVisualStyleBackColor = false;
            this.GetFile_btn.Click += new System.EventHandler(this.GetFile_btn_Click);
            // 
            // Start_btn
            // 
            this.Start_btn.BackColor = System.Drawing.Color.Transparent;
            this.Start_btn.BaseColor = System.Drawing.Color.OrangeRed;
            this.Start_btn.BorderColor = System.Drawing.Color.OrangeRed;
            this.Start_btn.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.Start_btn.DownBack = null;
            this.Start_btn.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Start_btn.ForeColor = System.Drawing.Color.White;
            this.Start_btn.Location = new System.Drawing.Point(412, 298);
            this.Start_btn.MouseBack = null;
            this.Start_btn.Name = "Start_btn";
            this.Start_btn.NormlBack = null;
            this.Start_btn.Size = new System.Drawing.Size(195, 77);
            this.Start_btn.TabIndex = 1;
            this.Start_btn.Text = "执行";
            this.Start_btn.UseVisualStyleBackColor = false;
            this.Start_btn.Click += new System.EventHandler(this.Start_btn_Click);
            // 
            // filePathLBL
            // 
            this.filePathLBL.AutoSize = true;
            this.filePathLBL.Location = new System.Drawing.Point(128, 160);
            this.filePathLBL.Name = "filePathLBL";
            this.filePathLBL.Size = new System.Drawing.Size(106, 24);
            this.filePathLBL.TabIndex = 2;
            this.filePathLBL.Text = "已选择：";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(722, 440);
            this.Controls.Add(this.filePathLBL);
            this.Controls.Add(this.Start_btn);
            this.Controls.Add(this.GetFile_btn);
            this.Name = "Main";
            this.Text = "ExamDA_Leakey";
            this.Load += new System.EventHandler(this.Main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private CCWin.SkinControl.SkinButton GetFile_btn;
        private CCWin.SkinControl.SkinButton Start_btn;
        private System.Windows.Forms.Label filePathLBL;
    }
}

