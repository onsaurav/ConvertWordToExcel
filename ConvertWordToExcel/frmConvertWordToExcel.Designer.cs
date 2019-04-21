namespace ConvertWordToExcel
{
    partial class frmConvertWordToExcel
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
            this.btnConvert = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnClose = new System.Windows.Forms.Button();
            this.txtFrom = new System.Windows.Forms.TextBox();
            this.txtTo = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnConvertTo = new System.Windows.Forms.Button();
            this.btnConvertFrom = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnConvert
            // 
            this.btnConvert.Location = new System.Drawing.Point(219, 87);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(75, 23);
            this.btnConvert.TabIndex = 0;
            this.btnConvert.Text = "Convert";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 69);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(363, 13);
            this.progressBar1.TabIndex = 1;
            this.progressBar1.Visible = false;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(300, 87);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 3;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // txtFrom
            // 
            this.txtFrom.Enabled = false;
            this.txtFrom.Location = new System.Drawing.Point(97, 18);
            this.txtFrom.Name = "txtFrom";
            this.txtFrom.Size = new System.Drawing.Size(254, 20);
            this.txtFrom.TabIndex = 4;
            // 
            // txtTo
            // 
            this.txtTo.Enabled = false;
            this.txtTo.Location = new System.Drawing.Point(97, 41);
            this.txtTo.Name = "txtTo";
            this.txtTo.Size = new System.Drawing.Size(254, 20);
            this.txtTo.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(9, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Convert From";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(9, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Convert To";
            // 
            // btnConvertTo
            // 
            this.btnConvertTo.Location = new System.Drawing.Point(357, 41);
            this.btnConvertTo.Name = "btnConvertTo";
            this.btnConvertTo.Size = new System.Drawing.Size(18, 21);
            this.btnConvertTo.TabIndex = 8;
            this.btnConvertTo.UseVisualStyleBackColor = true;
            this.btnConvertTo.Click += new System.EventHandler(this.btnConvertTo_Click);
            // 
            // btnConvertFrom
            // 
            this.btnConvertFrom.Location = new System.Drawing.Point(357, 18);
            this.btnConvertFrom.Name = "btnConvertFrom";
            this.btnConvertFrom.Size = new System.Drawing.Size(18, 21);
            this.btnConvertFrom.TabIndex = 9;
            this.btnConvertFrom.UseVisualStyleBackColor = true;
            this.btnConvertFrom.Click += new System.EventHandler(this.btnConvertFrom_Click);
            // 
            // frmConvertWordToExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(387, 123);
            this.Controls.Add(this.btnConvertFrom);
            this.Controls.Add(this.btnConvertTo);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtTo);
            this.Controls.Add(this.txtFrom);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnConvert);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmConvertWordToExcel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Convert Word to Excel";
            this.Load += new System.EventHandler(this.frmConvertWordToExcel_Load);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmConvertWordToExcel_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.TextBox txtFrom;
        private System.Windows.Forms.TextBox txtTo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnConvertTo;
        private System.Windows.Forms.Button btnConvertFrom;
    }
}

