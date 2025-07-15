namespace FidoAutoCad.Forms
{
    partial class FidoAutocadDock
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.showPrintOptions = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.printStartCheck = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dispLineProperties = new System.Windows.Forms.ComboBox();
            this.GetPropButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.launchExcel = new System.Windows.Forms.Button();
            this.DispExcelStatus = new System.Windows.Forms.Label();
            this.attachRunningExcel = new System.Windows.Forms.Button();
            this.detachExcel = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.testButton = new System.Windows.Forms.Button();
            this.tabPage1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabPage1
            // 
            this.tabPage1.AutoScroll = true;
            this.tabPage1.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage1.Controls.Add(this.groupBox2);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Location = new System.Drawing.Point(4, 33);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage1.Size = new System.Drawing.Size(531, 819);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Excel";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.testButton);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.showPrintOptions);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.checkBox2);
            this.groupBox2.Controls.Add(this.checkBox1);
            this.groupBox2.Controls.Add(this.printStartCheck);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.dispLineProperties);
            this.groupBox2.Controls.Add(this.GetPropButton);
            this.groupBox2.Location = new System.Drawing.Point(15, 186);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox2.Size = new System.Drawing.Size(508, 569);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Line Functions";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(7, 308);
            this.label3.Margin = new System.Windows.Forms.Padding(4);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(244, 31);
            this.label3.TabIndex = 51;
            this.label3.Text = "Area Conversion";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // showPrintOptions
            // 
            this.showPrintOptions.Location = new System.Drawing.Point(256, 27);
            this.showPrintOptions.Margin = new System.Windows.Forms.Padding(4);
            this.showPrintOptions.Name = "showPrintOptions";
            this.showPrintOptions.Size = new System.Drawing.Size(238, 39);
            this.showPrintOptions.TabIndex = 50;
            this.showPrintOptions.Text = "Set Print Output";
            this.showPrintOptions.UseVisualStyleBackColor = true;
            this.showPrintOptions.Click += new System.EventHandler(this.showPrintOptions_Click);
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(7, 270);
            this.label2.Margin = new System.Windows.Forms.Padding(4);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(244, 31);
            this.label2.TabIndex = 49;
            this.label2.Text = "Distance Conversion";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkBox2.Location = new System.Drawing.Point(216, 443);
            this.checkBox2.Margin = new System.Windows.Forms.Padding(4);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(170, 29);
            this.checkBox2.TabIndex = 48;
            this.checkBox2.Text = "Show Msg Box";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkBox1.Location = new System.Drawing.Point(216, 406);
            this.checkBox1.Margin = new System.Windows.Forms.Padding(4);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(190, 29);
            this.checkBox1.TabIndex = 47;
            this.checkBox1.Text = "Copy to clipboard";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // printStartCheck
            // 
            this.printStartCheck.AutoSize = true;
            this.printStartCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.printStartCheck.Location = new System.Drawing.Point(216, 369);
            this.printStartCheck.Margin = new System.Windows.Forms.Padding(4);
            this.printStartCheck.Name = "printStartCheck";
            this.printStartCheck.Size = new System.Drawing.Size(158, 29);
            this.printStartCheck.TabIndex = 46;
            this.printStartCheck.Text = "Write to Excel";
            this.printStartCheck.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(7, 233);
            this.label1.Margin = new System.Windows.Forms.Padding(4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(244, 31);
            this.label1.TabIndex = 12;
            this.label1.Text = "Rounding Options";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dispLineProperties
            // 
            this.dispLineProperties.AutoCompleteCustomSource.AddRange(new string[] {
            "No Rounding",
            "1000",
            "100",
            "10",
            "0",
            "0.0",
            "0.00",
            "0.000"});
            this.dispLineProperties.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.dispLineProperties.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.dispLineProperties.FormattingEnabled = true;
            this.dispLineProperties.Items.AddRange(new object[] {
            "No Rounding",
            "1000",
            "100",
            "10",
            "0",
            "0.0",
            "0.00",
            "0.000"});
            this.dispLineProperties.Location = new System.Drawing.Point(262, 233);
            this.dispLineProperties.Margin = new System.Windows.Forms.Padding(6);
            this.dispLineProperties.Name = "dispLineProperties";
            this.dispLineProperties.Size = new System.Drawing.Size(242, 32);
            this.dispLineProperties.TabIndex = 11;
            this.dispLineProperties.Text = "No Rounding";
            // 
            // GetPropButton
            // 
            this.GetPropButton.Location = new System.Drawing.Point(0, 27);
            this.GetPropButton.Margin = new System.Windows.Forms.Padding(4);
            this.GetPropButton.Name = "GetPropButton";
            this.GetPropButton.Size = new System.Drawing.Size(238, 39);
            this.GetPropButton.TabIndex = 10;
            this.GetPropButton.Text = "Get Properties";
            this.GetPropButton.UseVisualStyleBackColor = true;
            this.GetPropButton.Click += new System.EventHandler(this.GetPropButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.launchExcel);
            this.groupBox1.Controls.Add(this.DispExcelStatus);
            this.groupBox1.Controls.Add(this.attachRunningExcel);
            this.groupBox1.Controls.Add(this.detachExcel);
            this.groupBox1.Location = new System.Drawing.Point(15, 11);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(508, 170);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Excel Attachment";
            // 
            // launchExcel
            // 
            this.launchExcel.Location = new System.Drawing.Point(6, 28);
            this.launchExcel.Margin = new System.Windows.Forms.Padding(4);
            this.launchExcel.Name = "launchExcel";
            this.launchExcel.Size = new System.Drawing.Size(244, 39);
            this.launchExcel.TabIndex = 2;
            this.launchExcel.Text = "Launch New Instance";
            this.launchExcel.UseVisualStyleBackColor = true;
            this.launchExcel.Click += new System.EventHandler(this.launchExcel_Click);
            // 
            // DispExcelStatus
            // 
            this.DispExcelStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DispExcelStatus.Location = new System.Drawing.Point(6, 113);
            this.DispExcelStatus.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.DispExcelStatus.Name = "DispExcelStatus";
            this.DispExcelStatus.Size = new System.Drawing.Size(495, 50);
            this.DispExcelStatus.TabIndex = 6;
            this.DispExcelStatus.Text = "Application attached: False\r\nActive Workbook: NA\r\n";
            // 
            // attachRunningExcel
            // 
            this.attachRunningExcel.Location = new System.Drawing.Point(256, 28);
            this.attachRunningExcel.Margin = new System.Windows.Forms.Padding(4);
            this.attachRunningExcel.Name = "attachRunningExcel";
            this.attachRunningExcel.Size = new System.Drawing.Size(244, 39);
            this.attachRunningExcel.TabIndex = 8;
            this.attachRunningExcel.Text = "Attach to Existing Excel";
            this.attachRunningExcel.UseVisualStyleBackColor = true;
            this.attachRunningExcel.Click += new System.EventHandler(this.attachRunningExcel_Click);
            // 
            // detachExcel
            // 
            this.detachExcel.Location = new System.Drawing.Point(6, 72);
            this.detachExcel.Margin = new System.Windows.Forms.Padding(4);
            this.detachExcel.Name = "detachExcel";
            this.detachExcel.Size = new System.Drawing.Size(244, 39);
            this.detachExcel.TabIndex = 7;
            this.detachExcel.Text = "Detach Excel";
            this.detachExcel.UseVisualStyleBackColor = true;
            this.detachExcel.Click += new System.EventHandler(this.detachExcel_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(6, 6);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(6);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(539, 856);
            this.tabControl1.TabIndex = 13;
            // 
            // testButton
            // 
            this.testButton.Location = new System.Drawing.Point(0, 88);
            this.testButton.Margin = new System.Windows.Forms.Padding(4);
            this.testButton.Name = "testButton";
            this.testButton.Size = new System.Drawing.Size(238, 39);
            this.testButton.TabIndex = 52;
            this.testButton.Text = "Test Button";
            this.testButton.UseVisualStyleBackColor = true;
            this.testButton.Click += new System.EventHandler(this.testButton_Click);
            // 
            // FidoAutocadDock
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "FidoAutocadDock";
            this.Size = new System.Drawing.Size(550, 867);
            this.tabPage1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button GetPropButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button launchExcel;
        private System.Windows.Forms.Label DispExcelStatus;
        private System.Windows.Forms.Button attachRunningExcel;
        private System.Windows.Forms.Button detachExcel;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox dispLineProperties;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.CheckBox printStartCheck;
        private System.Windows.Forms.Button showPrintOptions;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button testButton;
    }
}
