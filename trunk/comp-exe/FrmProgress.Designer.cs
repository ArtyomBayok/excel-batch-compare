namespace compare_exe
{
    partial class FrmProgress
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
            this.pbWorkbooks = new System.Windows.Forms.ProgressBar();
            this.btCancel = new System.Windows.Forms.Button();
            this.txtInfo = new System.Windows.Forms.TextBox();
            this.pbWorksheets = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btOpen = new System.Windows.Forms.Button();
            this.btReportHTML = new System.Windows.Forms.Button();
            this.lbWB = new System.Windows.Forms.Label();
            this.lbWS = new System.Windows.Forms.Label();
            this.btReportXML = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // pbWorkbooks
            // 
            this.pbWorkbooks.Location = new System.Drawing.Point(129, 23);
            this.pbWorkbooks.Name = "pbWorkbooks";
            this.pbWorkbooks.Size = new System.Drawing.Size(257, 17);
            this.pbWorkbooks.TabIndex = 0;
            // 
            // btCancel
            // 
            this.btCancel.Location = new System.Drawing.Point(288, 193);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(84, 28);
            this.btCancel.TabIndex = 1;
            this.btCancel.Text = "Close";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // txtInfo
            // 
            this.txtInfo.AcceptsReturn = true;
            this.txtInfo.Location = new System.Drawing.Point(5, 46);
            this.txtInfo.Multiline = true;
            this.txtInfo.Name = "txtInfo";
            this.txtInfo.ReadOnly = true;
            this.txtInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtInfo.Size = new System.Drawing.Size(381, 140);
            this.txtInfo.TabIndex = 2;
            this.txtInfo.WordWrap = false;
            // 
            // pbWorksheets
            // 
            this.pbWorksheets.Location = new System.Drawing.Point(129, 3);
            this.pbWorksheets.Name = "pbWorksheets";
            this.pbWorksheets.Size = new System.Drawing.Size(257, 17);
            this.pbWorksheets.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(2, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Workbooks :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(2, 4);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Worksheets :";
            // 
            // btOpen
            // 
            this.btOpen.Enabled = false;
            this.btOpen.Location = new System.Drawing.Point(17, 192);
            this.btOpen.Name = "btOpen";
            this.btOpen.Size = new System.Drawing.Size(84, 28);
            this.btOpen.TabIndex = 1;
            this.btOpen.Text = "Open Folder";
            this.btOpen.UseVisualStyleBackColor = true;
            this.btOpen.Click += new System.EventHandler(this.btOpen_Click);
            // 
            // btReportHTML
            // 
            this.btReportHTML.Enabled = false;
            this.btReportHTML.Location = new System.Drawing.Point(107, 192);
            this.btReportHTML.Name = "btReportHTML";
            this.btReportHTML.Size = new System.Drawing.Size(84, 28);
            this.btReportHTML.TabIndex = 1;
            this.btReportHTML.Text = "Diff. Report";
            this.btReportHTML.UseVisualStyleBackColor = true;
            this.btReportHTML.Click += new System.EventHandler(this.btReportHTML_Click);
            // 
            // lbWB
            // 
            this.lbWB.AutoSize = true;
            this.lbWB.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbWB.Location = new System.Drawing.Point(85, 25);
            this.lbWB.Name = "lbWB";
            this.lbWB.Size = new System.Drawing.Size(28, 13);
            this.lbWB.TabIndex = 4;
            this.lbWB.Text = "0 / 0";
            // 
            // lbWS
            // 
            this.lbWS.AutoSize = true;
            this.lbWS.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbWS.Location = new System.Drawing.Point(85, 4);
            this.lbWS.Name = "lbWS";
            this.lbWS.Size = new System.Drawing.Size(28, 13);
            this.lbWS.TabIndex = 4;
            this.lbWS.Text = "0 / 0";
            // 
            // btReportXML
            // 
            this.btReportXML.Enabled = false;
            this.btReportXML.Location = new System.Drawing.Point(197, 193);
            this.btReportXML.Name = "btReportXML";
            this.btReportXML.Size = new System.Drawing.Size(85, 27);
            this.btReportXML.TabIndex = 5;
            this.btReportXML.Text = "Stat. Report";
            this.btReportXML.UseVisualStyleBackColor = true;
            this.btReportXML.Click += new System.EventHandler(this.btReportXML_Click);
            // 
            // fProgress
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(390, 227);
            this.Controls.Add(this.btReportXML);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lbWS);
            this.Controls.Add(this.lbWB);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pbWorksheets);
            this.Controls.Add(this.txtInfo);
            this.Controls.Add(this.btReportHTML);
            this.Controls.Add(this.btOpen);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.pbWorkbooks);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "fProgress";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Progress";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar pbWorkbooks;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.TextBox txtInfo;
        private System.Windows.Forms.ProgressBar pbWorksheets;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btOpen;
        private System.Windows.Forms.Button btReportHTML;
        private System.Windows.Forms.Label lbWB;
        private System.Windows.Forms.Label lbWS;
        private System.Windows.Forms.Button btReportXML;
    }
}