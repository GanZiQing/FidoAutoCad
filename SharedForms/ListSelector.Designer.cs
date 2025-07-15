namespace FidoAutoCad.SharedForms
{
    partial class ListSelector
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
        //private void InitializeComponent()
        //{
        //    this.components = new System.ComponentModel.Container();
        //    this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        //    this.ClientSize = new System.Drawing.Size(800, 450);
        //    this.Text = "ListSelector";
        //}
        private void InitializeComponent()
        {
            this.descriptionLabel = new System.Windows.Forms.Label();
            this.LeftLabel = new System.Windows.Forms.Label();
            this.RightLabel = new System.Windows.Forms.Label();
            this.LeftListBox = new System.Windows.Forms.ListBox();
            this.MyCancelButton = new System.Windows.Forms.Button();
            this.ConfirmationButton = new System.Windows.Forms.Button();
            this.RightListBox = new System.Windows.Forms.ListBox();
            this.MoveAllRight = new System.Windows.Forms.Button();
            this.MoveSelectionRight = new System.Windows.Forms.Button();
            this.MoveSelectionLeft = new System.Windows.Forms.Button();
            this.MoveAllLeft = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // descriptionLabel
            // 
            this.descriptionLabel.AutoSize = true;
            this.descriptionLabel.Location = new System.Drawing.Point(22, 17);
            this.descriptionLabel.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.descriptionLabel.Name = "descriptionLabel";
            this.descriptionLabel.Size = new System.Drawing.Size(109, 25);
            this.descriptionLabel.TabIndex = 0;
            this.descriptionLabel.Text = "Description";
            // 
            // LeftLabel
            // 
            this.LeftLabel.AutoSize = true;
            this.LeftLabel.Location = new System.Drawing.Point(22, 66);
            this.LeftLabel.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.LeftLabel.Name = "LeftLabel";
            this.LeftLabel.Size = new System.Drawing.Size(110, 25);
            this.LeftLabel.TabIndex = 1;
            this.LeftLabel.Text = "Unselected";
            // 
            // RightLabel
            // 
            this.RightLabel.AutoSize = true;
            this.RightLabel.Location = new System.Drawing.Point(686, 66);
            this.RightLabel.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.RightLabel.Name = "RightLabel";
            this.RightLabel.Size = new System.Drawing.Size(89, 25);
            this.RightLabel.TabIndex = 2;
            this.RightLabel.Text = "Selected";
            // 
            // LeftListBox
            // 
            this.LeftListBox.FormattingEnabled = true;
            this.LeftListBox.ItemHeight = 24;
            this.LeftListBox.Location = new System.Drawing.Point(28, 96);
            this.LeftListBox.Margin = new System.Windows.Forms.Padding(6);
            this.LeftListBox.Name = "LeftListBox";
            this.LeftListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.LeftListBox.Size = new System.Drawing.Size(583, 556);
            this.LeftListBox.TabIndex = 3;
            // 
            // MyCancelButton
            // 
            this.MyCancelButton.Location = new System.Drawing.Point(1113, 668);
            this.MyCancelButton.Margin = new System.Windows.Forms.Padding(6);
            this.MyCancelButton.Name = "MyCancelButton";
            this.MyCancelButton.Size = new System.Drawing.Size(165, 46);
            this.MyCancelButton.TabIndex = 5;
            this.MyCancelButton.Text = "Cancel";
            this.MyCancelButton.UseVisualStyleBackColor = true;
            this.MyCancelButton.Click += new System.EventHandler(this.MyCancelButton_Click);
            // 
            // ConfirmationButton
            // 
            this.ConfirmationButton.Location = new System.Drawing.Point(937, 668);
            this.ConfirmationButton.Margin = new System.Windows.Forms.Padding(6);
            this.ConfirmationButton.Name = "ConfirmationButton";
            this.ConfirmationButton.Size = new System.Drawing.Size(165, 46);
            this.ConfirmationButton.TabIndex = 6;
            this.ConfirmationButton.Text = "OK";
            this.ConfirmationButton.UseVisualStyleBackColor = true;
            this.ConfirmationButton.Click += new System.EventHandler(this.ConfirmationButton_Click);
            // 
            // RightListBox
            // 
            this.RightListBox.FormattingEnabled = true;
            this.RightListBox.ItemHeight = 24;
            this.RightListBox.Location = new System.Drawing.Point(691, 96);
            this.RightListBox.Margin = new System.Windows.Forms.Padding(6);
            this.RightListBox.Name = "RightListBox";
            this.RightListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.RightListBox.Size = new System.Drawing.Size(583, 556);
            this.RightListBox.TabIndex = 7;
            // 
            // MoveAllRight
            // 
            this.MoveAllRight.Location = new System.Drawing.Point(625, 198);
            this.MoveAllRight.Margin = new System.Windows.Forms.Padding(6);
            this.MoveAllRight.Name = "MoveAllRight";
            this.MoveAllRight.Size = new System.Drawing.Size(55, 46);
            this.MoveAllRight.TabIndex = 8;
            this.MoveAllRight.Text = ">>";
            this.MoveAllRight.UseVisualStyleBackColor = true;
            this.MoveAllRight.Click += new System.EventHandler(this.MoveAllRight_Click);
            // 
            // MoveSelectionRight
            // 
            this.MoveSelectionRight.Location = new System.Drawing.Point(625, 255);
            this.MoveSelectionRight.Margin = new System.Windows.Forms.Padding(6);
            this.MoveSelectionRight.Name = "MoveSelectionRight";
            this.MoveSelectionRight.Size = new System.Drawing.Size(55, 46);
            this.MoveSelectionRight.TabIndex = 9;
            this.MoveSelectionRight.Text = ">";
            this.MoveSelectionRight.UseVisualStyleBackColor = true;
            this.MoveSelectionRight.Click += new System.EventHandler(this.MoveSelectionRight_Click);
            // 
            // MoveSelectionLeft
            // 
            this.MoveSelectionLeft.Location = new System.Drawing.Point(625, 388);
            this.MoveSelectionLeft.Margin = new System.Windows.Forms.Padding(6);
            this.MoveSelectionLeft.Name = "MoveSelectionLeft";
            this.MoveSelectionLeft.Size = new System.Drawing.Size(55, 46);
            this.MoveSelectionLeft.TabIndex = 10;
            this.MoveSelectionLeft.Text = "<";
            this.MoveSelectionLeft.UseVisualStyleBackColor = true;
            this.MoveSelectionLeft.Click += new System.EventHandler(this.MoveSelectionLeft_Click);
            // 
            // MoveAllLeft
            // 
            this.MoveAllLeft.Location = new System.Drawing.Point(625, 445);
            this.MoveAllLeft.Margin = new System.Windows.Forms.Padding(6);
            this.MoveAllLeft.Name = "MoveAllLeft";
            this.MoveAllLeft.Size = new System.Drawing.Size(55, 46);
            this.MoveAllLeft.TabIndex = 11;
            this.MoveAllLeft.Text = "<<";
            this.MoveAllLeft.UseVisualStyleBackColor = true;
            this.MoveAllLeft.Click += new System.EventHandler(this.MoveAllLeft_Click);
            // 
            // ListSelector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1291, 737);
            this.Controls.Add(this.MoveAllLeft);
            this.Controls.Add(this.MoveSelectionLeft);
            this.Controls.Add(this.MoveSelectionRight);
            this.Controls.Add(this.MoveAllRight);
            this.Controls.Add(this.RightListBox);
            this.Controls.Add(this.ConfirmationButton);
            this.Controls.Add(this.MyCancelButton);
            this.Controls.Add(this.LeftListBox);
            this.Controls.Add(this.RightLabel);
            this.Controls.Add(this.LeftLabel);
            this.Controls.Add(this.descriptionLabel);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "ListSelector";
            this.Text = "Item Selector";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.Label descriptionLabel;
        private System.Windows.Forms.Label LeftLabel;
        private System.Windows.Forms.Label RightLabel;
        private System.Windows.Forms.ListBox LeftListBox;
        private System.Windows.Forms.Button MyCancelButton;
        private System.Windows.Forms.Button ConfirmationButton;
        private System.Windows.Forms.ListBox RightListBox;
        private System.Windows.Forms.Button MoveAllRight;
        private System.Windows.Forms.Button MoveSelectionRight;
        private System.Windows.Forms.Button MoveSelectionLeft;
        private System.Windows.Forms.Button MoveAllLeft;
    }
}