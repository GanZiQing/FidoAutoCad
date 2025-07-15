namespace FidoAutoCad.Forms
{
    partial class FidoAutoCadMain
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
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.launchExcel = new System.Windows.Forms.Button();
            this.DispExcelStatus = new System.Windows.Forms.Label();
            this.detachExcel = new System.Windows.Forms.Button();
            this.attachRunningExcel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(72, 344);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(149, 38);
            this.button1.TabIndex = 0;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(72, 309);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(390, 29);
            this.textBox1.TabIndex = 1;
            // 
            // launchExcel
            // 
            this.launchExcel.Location = new System.Drawing.Point(6, 28);
            this.launchExcel.Name = "launchExcel";
            this.launchExcel.Size = new System.Drawing.Size(214, 38);
            this.launchExcel.TabIndex = 2;
            this.launchExcel.Text = "Launch New Instance";
            this.launchExcel.UseVisualStyleBackColor = true;
            this.launchExcel.Click += new System.EventHandler(this.launchExcel_Click);
            // 
            // DispExcelStatus
            // 
            this.DispExcelStatus.AutoSize = true;
            this.DispExcelStatus.Location = new System.Drawing.Point(6, 69);
            this.DispExcelStatus.Name = "DispExcelStatus";
            this.DispExcelStatus.Size = new System.Drawing.Size(167, 25);
            this.DispExcelStatus.TabIndex = 6;
            this.DispExcelStatus.Text = "Application: False";
            // 
            // detachExcel
            // 
            this.detachExcel.Location = new System.Drawing.Point(469, 28);
            this.detachExcel.Name = "detachExcel";
            this.detachExcel.Size = new System.Drawing.Size(149, 38);
            this.detachExcel.TabIndex = 7;
            this.detachExcel.Text = "Detach Excel";
            this.detachExcel.UseVisualStyleBackColor = true;
            this.detachExcel.Click += new System.EventHandler(this.detachExcel_Click);
            // 
            // attachRunningExcel
            // 
            this.attachRunningExcel.Location = new System.Drawing.Point(226, 28);
            this.attachRunningExcel.Name = "attachRunningExcel";
            this.attachRunningExcel.Size = new System.Drawing.Size(237, 38);
            this.attachRunningExcel.TabIndex = 8;
            this.attachRunningExcel.Text = "Attach to Existing Excel";
            this.attachRunningExcel.UseVisualStyleBackColor = true;
            this.attachRunningExcel.Click += new System.EventHandler(this.attachRunningExcel_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.launchExcel);
            this.groupBox1.Controls.Add(this.DispExcelStatus);
            this.groupBox1.Controls.Add(this.attachRunningExcel);
            this.groupBox1.Controls.Add(this.detachExcel);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(673, 126);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Excel Attachment";
            // 
            // FidoAutoCadMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button1);
            this.Name = "FidoAutoCadMain";
            this.Text = "Fido AutoCAD";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button launchExcel;
        private System.Windows.Forms.Label DispExcelStatus;
        private System.Windows.Forms.Button detachExcel;
        private System.Windows.Forms.Button attachRunningExcel;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}