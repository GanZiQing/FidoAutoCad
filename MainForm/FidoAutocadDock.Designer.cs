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
            Autodesk.AutoCAD.ApplicationServices.Core.Application.QuitWillStart -= detachExcel_Auto;

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
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.printStartCheck = new System.Windows.Forms.CheckBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.Settings = new System.Windows.Forms.GroupBox();
            this.roundTypeComboBox = new System.Windows.Forms.ComboBox();
            this.translateByUcsCheck = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.skipInvalidCheck = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.skipLockCheck = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dispAreaConv = new System.Windows.Forms.TextBox();
            this.dispRoundOpt = new System.Windows.Forms.TextBox();
            this.dispDistConv = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.getMidPointButt = new System.Windows.Forms.Button();
            this.getCoordinatesButt = new System.Windows.Forms.Button();
            this.showPrintOptions = new System.Windows.Forms.Button();
            this.GetPropButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.launchExcel = new System.Windows.Forms.Button();
            this.DispExcelStatus = new System.Windows.Forms.Label();
            this.attachRunningExcel = new System.Windows.Forms.Button();
            this.detachExcel = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.button1 = new System.Windows.Forms.Button();
            this.tabPage1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.Settings.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabPage1
            // 
            this.tabPage1.AutoScroll = true;
            this.tabPage1.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.groupBox3);
            this.tabPage1.Controls.Add(this.Settings);
            this.tabPage1.Controls.Add(this.groupBox2);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Location = new System.Drawing.Point(4, 33);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage1.Size = new System.Drawing.Size(531, 1057);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Excel";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(123, 649);
            this.label4.Margin = new System.Windows.Forms.Padding(4);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(356, 132);
            this.label4.TabIndex = 61;
            this.label4.Text = "Rounding to ceiling and floor not implemented, Mid way through refractoring print" +
    " coordinates and mid points";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.printStartCheck);
            this.groupBox3.Controls.Add(this.checkBox1);
            this.groupBox3.Controls.Add(this.checkBox2);
            this.groupBox3.Location = new System.Drawing.Point(15, 760);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(508, 143);
            this.groupBox3.TabIndex = 15;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Output Options (not implemented)";
            // 
            // printStartCheck
            // 
            this.printStartCheck.AutoSize = true;
            this.printStartCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.printStartCheck.Location = new System.Drawing.Point(11, 29);
            this.printStartCheck.Margin = new System.Windows.Forms.Padding(4);
            this.printStartCheck.Name = "printStartCheck";
            this.printStartCheck.Size = new System.Drawing.Size(158, 29);
            this.printStartCheck.TabIndex = 46;
            this.printStartCheck.Text = "Write to Excel";
            this.printStartCheck.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkBox1.Location = new System.Drawing.Point(11, 66);
            this.checkBox1.Margin = new System.Windows.Forms.Padding(4);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(190, 29);
            this.checkBox1.TabIndex = 47;
            this.checkBox1.Text = "Copy to clipboard";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkBox2.Location = new System.Drawing.Point(11, 103);
            this.checkBox2.Margin = new System.Windows.Forms.Padding(4);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(170, 29);
            this.checkBox2.TabIndex = 48;
            this.checkBox2.Text = "Show Msg Box";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // Settings
            // 
            this.Settings.Controls.Add(this.button1);
            this.Settings.Controls.Add(this.roundTypeComboBox);
            this.Settings.Controls.Add(this.translateByUcsCheck);
            this.Settings.Controls.Add(this.label1);
            this.Settings.Controls.Add(this.skipInvalidCheck);
            this.Settings.Controls.Add(this.label2);
            this.Settings.Controls.Add(this.skipLockCheck);
            this.Settings.Controls.Add(this.label3);
            this.Settings.Controls.Add(this.dispAreaConv);
            this.Settings.Controls.Add(this.dispRoundOpt);
            this.Settings.Controls.Add(this.dispDistConv);
            this.Settings.Location = new System.Drawing.Point(15, 322);
            this.Settings.Name = "Settings";
            this.Settings.Size = new System.Drawing.Size(508, 292);
            this.Settings.TabIndex = 14;
            this.Settings.TabStop = false;
            this.Settings.Text = "Settings";
            // 
            // roundTypeComboBox
            // 
            this.roundTypeComboBox.FormattingEnabled = true;
            this.roundTypeComboBox.Items.AddRange(new object[] {
            "nearest",
            "ceiling",
            "floor"});
            this.roundTypeComboBox.Location = new System.Drawing.Point(188, 33);
            this.roundTypeComboBox.Name = "roundTypeComboBox";
            this.roundTypeComboBox.Size = new System.Drawing.Size(121, 32);
            this.roundTypeComboBox.TabIndex = 60;
            // 
            // translateByUcsCheck
            // 
            this.translateByUcsCheck.AutoSize = true;
            this.translateByUcsCheck.Checked = true;
            this.translateByUcsCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.translateByUcsCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.translateByUcsCheck.Location = new System.Drawing.Point(17, 248);
            this.translateByUcsCheck.Margin = new System.Windows.Forms.Padding(4);
            this.translateByUcsCheck.Name = "translateByUcsCheck";
            this.translateByUcsCheck.Size = new System.Drawing.Size(194, 29);
            this.translateByUcsCheck.TabIndex = 59;
            this.translateByUcsCheck.Text = "Translate by UCS";
            this.translateByUcsCheck.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(12, 29);
            this.label1.Margin = new System.Windows.Forms.Padding(4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(169, 39);
            this.label1.TabIndex = 12;
            this.label1.Text = "Round Value To";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // skipInvalidCheck
            // 
            this.skipInvalidCheck.AutoSize = true;
            this.skipInvalidCheck.Checked = true;
            this.skipInvalidCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.skipInvalidCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.skipInvalidCheck.Location = new System.Drawing.Point(17, 209);
            this.skipInvalidCheck.Margin = new System.Windows.Forms.Padding(4);
            this.skipInvalidCheck.Name = "skipInvalidCheck";
            this.skipInvalidCheck.Size = new System.Drawing.Size(210, 29);
            this.skipInvalidCheck.TabIndex = 57;
            this.skipInvalidCheck.Text = "Skip Invalid Objects";
            this.skipInvalidCheck.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(12, 76);
            this.label2.Margin = new System.Windows.Forms.Padding(4);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(238, 39);
            this.label2.TabIndex = 49;
            this.label2.Text = "Dist. Conversion";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // skipLockCheck
            // 
            this.skipLockCheck.AutoSize = true;
            this.skipLockCheck.Checked = true;
            this.skipLockCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.skipLockCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.skipLockCheck.Location = new System.Drawing.Point(17, 170);
            this.skipLockCheck.Margin = new System.Windows.Forms.Padding(4);
            this.skipLockCheck.Name = "skipLockCheck";
            this.skipLockCheck.Size = new System.Drawing.Size(309, 29);
            this.skipLockCheck.TabIndex = 56;
            this.skipLockCheck.Text = "Skip Objects on Locked Layers";
            this.skipLockCheck.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(12, 123);
            this.label3.Margin = new System.Windows.Forms.Padding(4);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(238, 39);
            this.label3.TabIndex = 51;
            this.label3.Text = "Area Conversion";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // dispAreaConv
            // 
            this.dispAreaConv.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispAreaConv.Location = new System.Drawing.Point(260, 127);
            this.dispAreaConv.Margin = new System.Windows.Forms.Padding(6);
            this.dispAreaConv.Name = "dispAreaConv";
            this.dispAreaConv.Size = new System.Drawing.Size(239, 29);
            this.dispAreaConv.TabIndex = 55;
            this.dispAreaConv.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispAreaConv.WordWrap = false;
            // 
            // dispRoundOpt
            // 
            this.dispRoundOpt.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispRoundOpt.Location = new System.Drawing.Point(318, 33);
            this.dispRoundOpt.Margin = new System.Windows.Forms.Padding(6);
            this.dispRoundOpt.Name = "dispRoundOpt";
            this.dispRoundOpt.Size = new System.Drawing.Size(182, 29);
            this.dispRoundOpt.TabIndex = 53;
            this.dispRoundOpt.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispRoundOpt.WordWrap = false;
            // 
            // dispDistConv
            // 
            this.dispDistConv.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispDistConv.Location = new System.Drawing.Point(260, 80);
            this.dispDistConv.Margin = new System.Windows.Forms.Padding(6);
            this.dispDistConv.Name = "dispDistConv";
            this.dispDistConv.Size = new System.Drawing.Size(239, 29);
            this.dispDistConv.TabIndex = 54;
            this.dispDistConv.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispDistConv.WordWrap = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.getMidPointButt);
            this.groupBox2.Controls.Add(this.getCoordinatesButt);
            this.groupBox2.Controls.Add(this.showPrintOptions);
            this.groupBox2.Controls.Add(this.GetPropButton);
            this.groupBox2.Location = new System.Drawing.Point(15, 188);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox2.Size = new System.Drawing.Size(508, 127);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Line Functions";
            // 
            // getMidPointButt
            // 
            this.getMidPointButt.Location = new System.Drawing.Point(262, 74);
            this.getMidPointButt.Margin = new System.Windows.Forms.Padding(4);
            this.getMidPointButt.Name = "getMidPointButt";
            this.getMidPointButt.Size = new System.Drawing.Size(238, 39);
            this.getMidPointButt.TabIndex = 53;
            this.getMidPointButt.Text = "Get Mid Points";
            this.getMidPointButt.UseVisualStyleBackColor = true;
            this.getMidPointButt.Click += new System.EventHandler(this.getMidPointButt_Click);
            // 
            // getCoordinatesButt
            // 
            this.getCoordinatesButt.Location = new System.Drawing.Point(11, 74);
            this.getCoordinatesButt.Margin = new System.Windows.Forms.Padding(4);
            this.getCoordinatesButt.Name = "getCoordinatesButt";
            this.getCoordinatesButt.Size = new System.Drawing.Size(238, 39);
            this.getCoordinatesButt.TabIndex = 52;
            this.getCoordinatesButt.Text = "Get Coordinates";
            this.getCoordinatesButt.UseVisualStyleBackColor = true;
            this.getCoordinatesButt.Click += new System.EventHandler(this.getCoordinatesButt_Click);
            // 
            // showPrintOptions
            // 
            this.showPrintOptions.Location = new System.Drawing.Point(262, 28);
            this.showPrintOptions.Margin = new System.Windows.Forms.Padding(4);
            this.showPrintOptions.Name = "showPrintOptions";
            this.showPrintOptions.Size = new System.Drawing.Size(238, 39);
            this.showPrintOptions.TabIndex = 50;
            this.showPrintOptions.Text = "Set Print Output";
            this.showPrintOptions.UseVisualStyleBackColor = true;
            this.showPrintOptions.Click += new System.EventHandler(this.showPrintOptions_Click);
            // 
            // GetPropButton
            // 
            this.GetPropButton.Location = new System.Drawing.Point(11, 28);
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
            this.attachRunningExcel.Location = new System.Drawing.Point(257, 28);
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
            this.tabControl1.Size = new System.Drawing.Size(539, 1094);
            this.tabControl1.TabIndex = 13;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(261, 221);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(238, 39);
            this.button1.TabIndex = 61;
            this.button1.Text = "Get Coordinates";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.getCoordinatesButt_Click2);
            // 
            // FidoAutocadDock
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "FidoAutocadDock";
            this.Size = new System.Drawing.Size(550, 1105);
            this.tabPage1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.Settings.ResumeLayout(false);
            this.Settings.PerformLayout();
            this.groupBox2.ResumeLayout(false);
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
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.CheckBox printStartCheck;
        private System.Windows.Forms.Button showPrintOptions;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button getCoordinatesButt;
        private System.Windows.Forms.TextBox dispAreaConv;
        private System.Windows.Forms.TextBox dispDistConv;
        private System.Windows.Forms.TextBox dispRoundOpt;
        private System.Windows.Forms.CheckBox skipLockCheck;
        private System.Windows.Forms.CheckBox skipInvalidCheck;
        private System.Windows.Forms.CheckBox translateByUcsCheck;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox Settings;
        private System.Windows.Forms.ComboBox roundTypeComboBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button getMidPointButt;
        private System.Windows.Forms.Button button1;
    }
}
