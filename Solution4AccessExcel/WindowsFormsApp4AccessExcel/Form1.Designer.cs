namespace WindowsFormsApp4AccessExcel
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.BtnCreate2k3Excel = new System.Windows.Forms.Button();
            this.BtnInport = new System.Windows.Forms.Button();
            this.DgwResult = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.BtnCreate2k7Excel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DgwResult)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnCreate2k3Excel
            // 
            this.BtnCreate2k3Excel.Font = new System.Drawing.Font("新細明體", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.BtnCreate2k3Excel.Location = new System.Drawing.Point(12, 12);
            this.BtnCreate2k3Excel.Name = "BtnCreate2k3Excel";
            this.BtnCreate2k3Excel.Size = new System.Drawing.Size(224, 40);
            this.BtnCreate2k3Excel.TabIndex = 0;
            this.BtnCreate2k3Excel.Text = "新增 2003 EXCEL";
            this.BtnCreate2k3Excel.UseVisualStyleBackColor = true;
            this.BtnCreate2k3Excel.Click += new System.EventHandler(this.BtnCreate2k3Excel_Click);
            // 
            // BtnInport
            // 
            this.BtnInport.Font = new System.Drawing.Font("新細明體", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.BtnInport.Location = new System.Drawing.Point(472, 12);
            this.BtnInport.Name = "BtnInport";
            this.BtnInport.Size = new System.Drawing.Size(159, 40);
            this.BtnInport.TabIndex = 1;
            this.BtnInport.Text = "載入EXCEL";
            this.BtnInport.UseVisualStyleBackColor = true;
            this.BtnInport.Click += new System.EventHandler(this.BtnInport_Click);
            // 
            // DgwResult
            // 
            this.DgwResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgwResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DgwResult.Location = new System.Drawing.Point(0, 0);
            this.DgwResult.Name = "DgwResult";
            this.DgwResult.RowTemplate.Height = 24;
            this.DgwResult.Size = new System.Drawing.Size(728, 176);
            this.DgwResult.TabIndex = 2;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.BtnCreate2k7Excel);
            this.panel1.Controls.Add(this.BtnCreate2k3Excel);
            this.panel1.Controls.Add(this.BtnInport);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(728, 64);
            this.panel1.TabIndex = 3;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.DgwResult);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 64);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(728, 176);
            this.panel2.TabIndex = 4;
            // 
            // BtnCreate2k7Excel
            // 
            this.BtnCreate2k7Excel.Font = new System.Drawing.Font("新細明體", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.BtnCreate2k7Excel.Location = new System.Drawing.Point(242, 12);
            this.BtnCreate2k7Excel.Name = "BtnCreate2k7Excel";
            this.BtnCreate2k7Excel.Size = new System.Drawing.Size(224, 40);
            this.BtnCreate2k7Excel.TabIndex = 2;
            this.BtnCreate2k7Excel.Text = "新增 2007 EXCEL";
            this.BtnCreate2k7Excel.UseVisualStyleBackColor = true;
            this.BtnCreate2k7Excel.Click += new System.EventHandler(this.BtnCreate2k7Excel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(728, 240);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "Form1";
            this.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.DgwResult)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button BtnCreate2k3Excel;
        private System.Windows.Forms.Button BtnInport;
        private System.Windows.Forms.DataGridView DgwResult;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button BtnCreate2k7Excel;
    }
}

