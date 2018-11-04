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
            this.BtnCreate2k7Excel = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.BtnUpdateExcelWithoutNewCell = new System.Windows.Forms.Button();
            this.BtnUpdateExcelWithNewCell = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DgwResult)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnCreate2k3Excel
            // 
            this.BtnCreate2k3Excel.Font = new System.Drawing.Font("新細明體", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.BtnCreate2k3Excel.Location = new System.Drawing.Point(18, 18);
            this.BtnCreate2k3Excel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.BtnCreate2k3Excel.Name = "BtnCreate2k3Excel";
            this.BtnCreate2k3Excel.Size = new System.Drawing.Size(336, 60);
            this.BtnCreate2k3Excel.TabIndex = 0;
            this.BtnCreate2k3Excel.Text = "新增 2003 EXCEL";
            this.BtnCreate2k3Excel.UseVisualStyleBackColor = true;
            this.BtnCreate2k3Excel.Click += new System.EventHandler(this.BtnCreate2k3Excel_Click);
            // 
            // BtnInport
            // 
            this.BtnInport.Font = new System.Drawing.Font("新細明體", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.BtnInport.Location = new System.Drawing.Point(708, 18);
            this.BtnInport.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.BtnInport.Name = "BtnInport";
            this.BtnInport.Size = new System.Drawing.Size(238, 60);
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
            this.DgwResult.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.DgwResult.Name = "DgwResult";
            this.DgwResult.RowTemplate.Height = 24;
            this.DgwResult.Size = new System.Drawing.Size(1092, 173);
            this.DgwResult.TabIndex = 2;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.BtnUpdateExcelWithNewCell);
            this.panel1.Controls.Add(this.BtnUpdateExcelWithoutNewCell);
            this.panel1.Controls.Add(this.BtnCreate2k7Excel);
            this.panel1.Controls.Add(this.BtnCreate2k3Excel);
            this.panel1.Controls.Add(this.BtnInport);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1092, 187);
            this.panel1.TabIndex = 3;
            // 
            // BtnCreate2k7Excel
            // 
            this.BtnCreate2k7Excel.Font = new System.Drawing.Font("新細明體", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.BtnCreate2k7Excel.Location = new System.Drawing.Point(363, 18);
            this.BtnCreate2k7Excel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.BtnCreate2k7Excel.Name = "BtnCreate2k7Excel";
            this.BtnCreate2k7Excel.Size = new System.Drawing.Size(336, 60);
            this.BtnCreate2k7Excel.TabIndex = 2;
            this.BtnCreate2k7Excel.Text = "新增 2007 EXCEL";
            this.BtnCreate2k7Excel.UseVisualStyleBackColor = true;
            this.BtnCreate2k7Excel.Click += new System.EventHandler(this.BtnCreate2k7Excel_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.DgwResult);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 187);
            this.panel2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1092, 173);
            this.panel2.TabIndex = 4;
            // 
            // BtnUpdateExcelWithoutNewCell
            // 
            this.BtnUpdateExcelWithoutNewCell.Font = new System.Drawing.Font("新細明體", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.BtnUpdateExcelWithoutNewCell.Location = new System.Drawing.Point(13, 95);
            this.BtnUpdateExcelWithoutNewCell.Margin = new System.Windows.Forms.Padding(4);
            this.BtnUpdateExcelWithoutNewCell.Name = "BtnUpdateExcelWithoutNewCell";
            this.BtnUpdateExcelWithoutNewCell.Size = new System.Drawing.Size(454, 60);
            this.BtnUpdateExcelWithoutNewCell.TabIndex = 3;
            this.BtnUpdateExcelWithoutNewCell.Text = "修改Excel(不新增欄位)";
            this.BtnUpdateExcelWithoutNewCell.UseVisualStyleBackColor = true;
            this.BtnUpdateExcelWithoutNewCell.Click += new System.EventHandler(this.BtnUpdateExcelWithoutNewCell_Click);
            // 
            // BtnUpdateExcelWithNewCell
            // 
            this.BtnUpdateExcelWithNewCell.Font = new System.Drawing.Font("新細明體", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.BtnUpdateExcelWithNewCell.Location = new System.Drawing.Point(492, 95);
            this.BtnUpdateExcelWithNewCell.Margin = new System.Windows.Forms.Padding(4);
            this.BtnUpdateExcelWithNewCell.Name = "BtnUpdateExcelWithNewCell";
            this.BtnUpdateExcelWithNewCell.Size = new System.Drawing.Size(454, 60);
            this.BtnUpdateExcelWithNewCell.TabIndex = 4;
            this.BtnUpdateExcelWithNewCell.Text = "修改Excel(新增欄位)";
            this.BtnUpdateExcelWithNewCell.UseVisualStyleBackColor = true;
            this.BtnUpdateExcelWithNewCell.Click += new System.EventHandler(this.BtnUpdateExcelWithNewCell_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1092, 360);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
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
        private System.Windows.Forms.Button BtnUpdateExcelWithoutNewCell;
        private System.Windows.Forms.Button BtnUpdateExcelWithNewCell;
    }
}

