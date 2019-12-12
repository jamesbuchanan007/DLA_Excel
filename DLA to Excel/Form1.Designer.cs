namespace DLA_to_Excel
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.cmbTableName = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.txtFolderDestination = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnFileLocation = new System.Windows.Forms.Button();
            this.btnFolderDestination = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProgress = new System.Windows.Forms.Label();
            this.lblStart = new System.Windows.Forms.Label();
            this.lblStop = new System.Windows.Forms.Label();
            this.lblElapsed = new System.Windows.Forms.Label();
            this.lblCurrentDoc = new System.Windows.Forms.Label();
            this.btnRowCount = new System.Windows.Forms.Button();
            this.lblRowCount = new System.Windows.Forms.Label();
            this.btnSample = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // cmbTableName
            // 
            this.cmbTableName.FormattingEnabled = true;
            this.cmbTableName.Items.AddRange(new object[] {
            "AMMO",
            "CAGECDS",
            "CAGE-DateEstAndChgd",
            "CHARDAT",
            "COLXREF",
            "ENAC",
            "FCAGE",
            "FCAN-SEGK",
            "FLISFOIA",
            "FLISV",
            "FSC",
            "FSG",
            "H4H8",
            "H5",
            "MRD06P1",
            "MRD06P2",
            "MRD0107",
            "MRD0300",
            "MRD0500",
            "NAME",
            "SCHEDLB-FOIA",
            "SEGKFOI"});
            this.cmbTableName.Location = new System.Drawing.Point(51, 100);
            this.cmbTableName.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.cmbTableName.Name = "cmbTableName";
            this.cmbTableName.Size = new System.Drawing.Size(289, 39);
            this.cmbTableName.TabIndex = 0;
            this.toolTip1.SetToolTip(this.cmbTableName, "Choose Table");
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(43, 62);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(192, 32);
            this.label1.TabIndex = 1;
            this.label1.Text = "Choose Table";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(43, 186);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(283, 32);
            this.label2.TabIndex = 3;
            this.label2.Text = "Choose File Location";
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(51, 222);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(1001, 38);
            this.txtFilePath.TabIndex = 4;
            // 
            // txtFolderDestination
            // 
            this.txtFolderDestination.Location = new System.Drawing.Point(51, 348);
            this.txtFolderDestination.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtFolderDestination.Name = "txtFolderDestination";
            this.txtFolderDestination.Size = new System.Drawing.Size(1001, 38);
            this.txtFolderDestination.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(43, 312);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(352, 32);
            this.label3.TabIndex = 6;
            this.label3.Text = "Choose Folder Destination";
            // 
            // btnFileLocation
            // 
            this.btnFileLocation.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnFileLocation.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFileLocation.ForeColor = System.Drawing.Color.White;
            this.btnFileLocation.Location = new System.Drawing.Point(1096, 217);
            this.btnFileLocation.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnFileLocation.Name = "btnFileLocation";
            this.btnFileLocation.Size = new System.Drawing.Size(141, 55);
            this.btnFileLocation.TabIndex = 7;
            this.btnFileLocation.Text = "Open";
            this.toolTip1.SetToolTip(this.btnFileLocation, "Get File Location");
            this.btnFileLocation.UseVisualStyleBackColor = false;
            this.btnFileLocation.Click += new System.EventHandler(this.BtnFileLocation_Click);
            // 
            // btnFolderDestination
            // 
            this.btnFolderDestination.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnFolderDestination.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFolderDestination.ForeColor = System.Drawing.Color.White;
            this.btnFolderDestination.Location = new System.Drawing.Point(1096, 343);
            this.btnFolderDestination.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnFolderDestination.Name = "btnFolderDestination";
            this.btnFolderDestination.Size = new System.Drawing.Size(141, 55);
            this.btnFolderDestination.TabIndex = 8;
            this.btnFolderDestination.Text = "Open";
            this.toolTip1.SetToolTip(this.btnFolderDestination, "Get Folder Path");
            this.btnFolderDestination.UseVisualStyleBackColor = false;
            this.btnFolderDestination.Click += new System.EventHandler(this.BtnFolderDestination_Click);
            // 
            // btnStart
            // 
            this.btnStart.BackColor = System.Drawing.Color.Green;
            this.btnStart.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnStart.ForeColor = System.Drawing.Color.White;
            this.btnStart.Location = new System.Drawing.Point(1032, 625);
            this.btnStart.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(221, 60);
            this.btnStart.TabIndex = 10;
            this.btnStart.Text = "Convert File";
            this.toolTip1.SetToolTip(this.btnStart, "Begin Conversion");
            this.btnStart.UseVisualStyleBackColor = false;
            this.btnStart.Click += new System.EventHandler(this.BtnStart_Click);
            // 
            // btnExit
            // 
            this.btnExit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExit.ForeColor = System.Drawing.Color.White;
            this.btnExit.Location = new System.Drawing.Point(51, 625);
            this.btnExit.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(165, 60);
            this.btnExit.TabIndex = 12;
            this.btnExit.Text = "Exit";
            this.toolTip1.SetToolTip(this.btnExit, "Exit Application");
            this.btnExit.UseVisualStyleBackColor = false;
            this.btnExit.Click += new System.EventHandler(this.BtnExit_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(432, 98);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(805, 52);
            this.progressBar1.TabIndex = 13;
            this.toolTip1.SetToolTip(this.progressBar1, "Progress Bar");
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(427, 62);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(136, 32);
            this.lblProgress.TabIndex = 14;
            this.lblProgress.Text = "Progress:";
            // 
            // lblStart
            // 
            this.lblStart.AutoSize = true;
            this.lblStart.Location = new System.Drawing.Point(280, 539);
            this.lblStart.Name = "lblStart";
            this.lblStart.Size = new System.Drawing.Size(83, 32);
            this.lblStart.TabIndex = 15;
            this.lblStart.Text = "Start:";
            // 
            // lblStop
            // 
            this.lblStop.AutoSize = true;
            this.lblStop.Location = new System.Drawing.Point(280, 577);
            this.lblStop.Name = "lblStop";
            this.lblStop.Size = new System.Drawing.Size(82, 32);
            this.lblStop.TabIndex = 16;
            this.lblStop.Text = "Stop:";
            // 
            // lblElapsed
            // 
            this.lblElapsed.AutoSize = true;
            this.lblElapsed.Location = new System.Drawing.Point(283, 610);
            this.lblElapsed.Name = "lblElapsed";
            this.lblElapsed.Size = new System.Drawing.Size(127, 32);
            this.lblElapsed.TabIndex = 17;
            this.lblElapsed.Text = "Elapsed:";
            // 
            // lblCurrentDoc
            // 
            this.lblCurrentDoc.AutoSize = true;
            this.lblCurrentDoc.Location = new System.Drawing.Point(283, 501);
            this.lblCurrentDoc.Name = "lblCurrentDoc";
            this.lblCurrentDoc.Size = new System.Drawing.Size(172, 32);
            this.lblCurrentDoc.TabIndex = 18;
            this.lblCurrentDoc.Text = "Working On:";
            // 
            // btnRowCount
            // 
            this.btnRowCount.BackColor = System.Drawing.Color.Green;
            this.btnRowCount.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRowCount.ForeColor = System.Drawing.Color.White;
            this.btnRowCount.Location = new System.Drawing.Point(51, 444);
            this.btnRowCount.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnRowCount.Name = "btnRowCount";
            this.btnRowCount.Size = new System.Drawing.Size(171, 55);
            this.btnRowCount.TabIndex = 19;
            this.btnRowCount.Text = "Row Count";
            this.toolTip1.SetToolTip(this.btnRowCount, "Get Total Number of Rows and Excel Books");
            this.btnRowCount.UseVisualStyleBackColor = false;
            this.btnRowCount.Click += new System.EventHandler(this.BtnRowCount_Click);
            // 
            // lblRowCount
            // 
            this.lblRowCount.AutoSize = true;
            this.lblRowCount.Location = new System.Drawing.Point(283, 455);
            this.lblRowCount.Name = "lblRowCount";
            this.lblRowCount.Size = new System.Drawing.Size(162, 32);
            this.lblRowCount.TabIndex = 20;
            this.lblRowCount.Text = "Row Count:";
            // 
            // btnSample
            // 
            this.btnSample.BackColor = System.Drawing.Color.Green;
            this.btnSample.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSample.ForeColor = System.Drawing.Color.White;
            this.btnSample.Location = new System.Drawing.Point(51, 529);
            this.btnSample.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnSample.Name = "btnSample";
            this.btnSample.Size = new System.Drawing.Size(171, 55);
            this.btnSample.TabIndex = 21;
            this.btnSample.Text = "Sample";
            this.toolTip1.SetToolTip(this.btnSample, "See 1st 20 Lines in Text Form");
            this.btnSample.UseVisualStyleBackColor = false;
            this.btnSample.Click += new System.EventHandler(this.BtnSample_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(16F, 31F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Navy;
            this.ClientSize = new System.Drawing.Size(1280, 730);
            this.Controls.Add(this.btnSample);
            this.Controls.Add(this.lblRowCount);
            this.Controls.Add(this.btnRowCount);
            this.Controls.Add(this.lblCurrentDoc);
            this.Controls.Add(this.lblElapsed);
            this.Controls.Add(this.lblStop);
            this.Controls.Add(this.lblStart);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnFolderDestination);
            this.Controls.Add(this.btnFileLocation);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtFolderDestination);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbTableName);
            this.ForeColor = System.Drawing.Color.White;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "DLA to Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbTableName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.TextBox txtFolderDestination;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnFileLocation;
        private System.Windows.Forms.Button btnFolderDestination;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.Label lblStart;
        private System.Windows.Forms.Label lblStop;
        private System.Windows.Forms.Label lblElapsed;
        private System.Windows.Forms.Label lblCurrentDoc;
		private System.Windows.Forms.Button btnRowCount;
		private System.Windows.Forms.Label lblRowCount;
        private System.Windows.Forms.Button btnSample;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}

