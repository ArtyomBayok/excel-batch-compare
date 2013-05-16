namespace compare_exe
{
    partial class FrmCompare
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmCompare));
            this.btSelectA = new System.Windows.Forms.Button();
            this.btClearA = new System.Windows.Forms.Button();
            this.lsFilesA = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.btClearB = new System.Windows.Forms.Button();
            this.btSelectB = new System.Windows.Forms.Button();
            this.lsFilesB = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btCompare = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.cbOptStyle = new System.Windows.Forms.CheckBox();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.cbOptShape = new System.Windows.Forms.CheckBox();
            this.cbOptValue = new System.Windows.Forms.CheckBox();
            this.btSelFolderA = new System.Windows.Forms.Button();
            this.btSelFolderB = new System.Windows.Forms.Button();
            this.pictureBox5 = new System.Windows.Forms.PictureBox();
            this.label4 = new System.Windows.Forms.Label();
            this.ctrlFolder = new System.Windows.Forms.TextBox();
            this.ctrlBrowse = new System.Windows.Forms.Button();
            this.ctrlHelp = new System.Windows.Forms.Button();
            this.ctrlAbout = new System.Windows.Forms.Button();
            this.ctrlCreateXml = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btSelectA
            // 
            this.btSelectA.Location = new System.Drawing.Point(444, 42);
            this.btSelectA.Name = "btSelectA";
            this.btSelectA.Size = new System.Drawing.Size(49, 38);
            this.btSelectA.TabIndex = 0;
            this.btSelectA.Text = "Add Files";
            this.btSelectA.UseVisualStyleBackColor = true;
            this.btSelectA.Click += new System.EventHandler(this.btSelectA_Click);
            // 
            // btClearA
            // 
            this.btClearA.Location = new System.Drawing.Point(444, 131);
            this.btClearA.Name = "btClearA";
            this.btClearA.Size = new System.Drawing.Size(50, 32);
            this.btClearA.TabIndex = 0;
            this.btClearA.Text = "Clear";
            this.btClearA.UseVisualStyleBackColor = true;
            this.btClearA.Click += new System.EventHandler(this.btClearA_Click);
            // 
            // lsFilesA
            // 
            this.lsFilesA.AllowDrop = true;
            this.lsFilesA.FormattingEnabled = true;
            this.lsFilesA.Location = new System.Drawing.Point(7, 42);
            this.lsFilesA.Name = "lsFilesA";
            this.lsFilesA.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lsFilesA.Size = new System.Drawing.Size(436, 121);
            this.lsFilesA.TabIndex = 1;
            this.lsFilesA.DragDrop += new System.Windows.Forms.DragEventHandler(this.lsFilesA_DragDrop);
            this.lsFilesA.DragOver += new System.Windows.Forms.DragEventHandler(this.lsFilesA_DragOver);
            this.lsFilesA.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lsFilesA_KeyDown);
            this.lsFilesA.MouseDown += new System.Windows.Forms.MouseEventHandler(this.lsFilesA_MouseDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.label1.Location = new System.Drawing.Point(0, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(109, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Excel files selection A";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Silver;
            this.pictureBox1.Location = new System.Drawing.Point(3, 34);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(491, 1);
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.Silver;
            this.pictureBox2.Location = new System.Drawing.Point(3, 174);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(491, 1);
            this.pictureBox2.TabIndex = 3;
            this.pictureBox2.TabStop = false;
            // 
            // btClearB
            // 
            this.btClearB.Location = new System.Drawing.Point(444, 273);
            this.btClearB.Name = "btClearB";
            this.btClearB.Size = new System.Drawing.Size(49, 30);
            this.btClearB.TabIndex = 0;
            this.btClearB.Text = "Clear";
            this.btClearB.UseVisualStyleBackColor = true;
            this.btClearB.Click += new System.EventHandler(this.btClearB_Click);
            // 
            // btSelectB
            // 
            this.btSelectB.Location = new System.Drawing.Point(444, 182);
            this.btSelectB.Name = "btSelectB";
            this.btSelectB.Size = new System.Drawing.Size(49, 39);
            this.btSelectB.TabIndex = 0;
            this.btSelectB.Text = "Add Files";
            this.btSelectB.UseVisualStyleBackColor = true;
            this.btSelectB.Click += new System.EventHandler(this.btSelectB_Click);
            // 
            // lsFilesB
            // 
            this.lsFilesB.AllowDrop = true;
            this.lsFilesB.FormattingEnabled = true;
            this.lsFilesB.Location = new System.Drawing.Point(7, 182);
            this.lsFilesB.Name = "lsFilesB";
            this.lsFilesB.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lsFilesB.Size = new System.Drawing.Size(436, 121);
            this.lsFilesB.TabIndex = 1;
            this.lsFilesB.DragDrop += new System.Windows.Forms.DragEventHandler(this.lsFilesB_DragDrop);
            this.lsFilesB.DragOver += new System.Windows.Forms.DragEventHandler(this.lsFilesB_DragOver);
            this.lsFilesB.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lsFilesB_KeyDown);
            this.lsFilesB.MouseDown += new System.Windows.Forms.MouseEventHandler(this.lsFilesB_MouseDown);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.label2.Location = new System.Drawing.Point(0, 166);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(109, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Excel files selection B";
            // 
            // pictureBox3
            // 
            this.pictureBox3.BackColor = System.Drawing.Color.Silver;
            this.pictureBox3.Location = new System.Drawing.Point(4, 354);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(491, 1);
            this.pictureBox3.TabIndex = 4;
            this.pictureBox3.TabStop = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.label3.Location = new System.Drawing.Point(1, 347);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(43, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Options";
            // 
            // btCompare
            // 
            this.btCompare.Location = new System.Drawing.Point(155, 402);
            this.btCompare.Name = "btCompare";
            this.btCompare.Size = new System.Drawing.Size(89, 26);
            this.btCompare.TabIndex = 6;
            this.btCompare.Text = "Compare";
            this.btCompare.UseVisualStyleBackColor = true;
            this.btCompare.Click += new System.EventHandler(this.btCompare_Click);
            // 
            // btCancel
            // 
            this.btCancel.Location = new System.Drawing.Point(259, 402);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(88, 26);
            this.btCancel.TabIndex = 7;
            this.btCancel.Text = "Exit";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // cbOptStyle
            // 
            this.cbOptStyle.AutoSize = true;
            this.cbOptStyle.Location = new System.Drawing.Point(139, 367);
            this.cbOptStyle.Name = "cbOptStyle";
            this.cbOptStyle.Size = new System.Drawing.Size(140, 17);
            this.cbOptStyle.TabIndex = 8;
            this.cbOptStyle.Text = "Compare Styles (Slower)";
            this.cbOptStyle.UseVisualStyleBackColor = true;
            // 
            // pictureBox4
            // 
            this.pictureBox4.BackColor = System.Drawing.Color.Silver;
            this.pictureBox4.Location = new System.Drawing.Point(4, 395);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(491, 1);
            this.pictureBox4.TabIndex = 4;
            this.pictureBox4.TabStop = false;
            // 
            // cbOptShape
            // 
            this.cbOptShape.AutoSize = true;
            this.cbOptShape.Location = new System.Drawing.Point(307, 367);
            this.cbOptShape.Name = "cbOptShape";
            this.cbOptShape.Size = new System.Drawing.Size(107, 17);
            this.cbOptShape.TabIndex = 8;
            this.cbOptShape.Text = "Compare Shapes";
            this.cbOptShape.UseVisualStyleBackColor = true;
            // 
            // cbOptValue
            // 
            this.cbOptValue.AutoSize = true;
            this.cbOptValue.Checked = true;
            this.cbOptValue.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbOptValue.Location = new System.Drawing.Point(12, 367);
            this.cbOptValue.Name = "cbOptValue";
            this.cbOptValue.Size = new System.Drawing.Size(103, 17);
            this.cbOptValue.TabIndex = 9;
            this.cbOptValue.Text = "Compare Values";
            this.cbOptValue.UseVisualStyleBackColor = true;
            // 
            // btSelFolderA
            // 
            this.btSelFolderA.Location = new System.Drawing.Point(445, 86);
            this.btSelFolderA.Name = "btSelFolderA";
            this.btSelFolderA.Size = new System.Drawing.Size(48, 39);
            this.btSelFolderA.TabIndex = 10;
            this.btSelFolderA.Text = "Select Folder";
            this.btSelFolderA.UseVisualStyleBackColor = true;
            this.btSelFolderA.Click += new System.EventHandler(this.btSelFolderA_Click);
            // 
            // btSelFolderB
            // 
            this.btSelFolderB.Location = new System.Drawing.Point(444, 227);
            this.btSelFolderB.Name = "btSelFolderB";
            this.btSelFolderB.Size = new System.Drawing.Size(49, 40);
            this.btSelFolderB.TabIndex = 11;
            this.btSelFolderB.Text = "Select Folder";
            this.btSelFolderB.UseVisualStyleBackColor = true;
            this.btSelFolderB.Click += new System.EventHandler(this.btSelFolderB_Click);
            // 
            // pictureBox5
            // 
            this.pictureBox5.BackColor = System.Drawing.Color.Silver;
            this.pictureBox5.Location = new System.Drawing.Point(4, 313);
            this.pictureBox5.Name = "pictureBox5";
            this.pictureBox5.Size = new System.Drawing.Size(491, 1);
            this.pictureBox5.TabIndex = 4;
            this.pictureBox5.TabStop = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.label4.Location = new System.Drawing.Point(1, 306);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(90, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Working directory";
            // 
            // ctrlFolder
            // 
            this.ctrlFolder.Location = new System.Drawing.Point(7, 323);
            this.ctrlFolder.Name = "ctrlFolder";
            this.ctrlFolder.ReadOnly = true;
            this.ctrlFolder.Size = new System.Drawing.Size(436, 20);
            this.ctrlFolder.TabIndex = 13;
            this.ctrlFolder.DoubleClick += new System.EventHandler(this.ctrlBrowse_Click);
            // 
            // ctrlBrowse
            // 
            this.ctrlBrowse.Image = ((System.Drawing.Image)(resources.GetObject("ctrlBrowse.Image")));
            this.ctrlBrowse.Location = new System.Drawing.Point(444, 321);
            this.ctrlBrowse.Margin = new System.Windows.Forms.Padding(0);
            this.ctrlBrowse.Name = "ctrlBrowse";
            this.ctrlBrowse.Size = new System.Drawing.Size(49, 24);
            this.ctrlBrowse.TabIndex = 38;
            this.ctrlBrowse.UseVisualStyleBackColor = true;
            this.ctrlBrowse.Click += new System.EventHandler(this.ctrlBrowse_Click);
            // 
            // ctrlHelp
            // 
            this.ctrlHelp.FlatAppearance.BorderSize = 0;
            this.ctrlHelp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ctrlHelp.Location = new System.Drawing.Point(405, 1);
            this.ctrlHelp.Margin = new System.Windows.Forms.Padding(0);
            this.ctrlHelp.Name = "ctrlHelp";
            this.ctrlHelp.Size = new System.Drawing.Size(38, 21);
            this.ctrlHelp.TabIndex = 72;
            this.ctrlHelp.Text = "Help";
            this.ctrlHelp.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ctrlHelp.UseVisualStyleBackColor = true;
            this.ctrlHelp.Click += new System.EventHandler(this.ctrlHomePage_Click);
            // 
            // ctrlAbout
            // 
            this.ctrlAbout.FlatAppearance.BorderSize = 0;
            this.ctrlAbout.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ctrlAbout.Location = new System.Drawing.Point(448, 1);
            this.ctrlAbout.Margin = new System.Windows.Forms.Padding(0);
            this.ctrlAbout.Name = "ctrlAbout";
            this.ctrlAbout.Size = new System.Drawing.Size(45, 21);
            this.ctrlAbout.TabIndex = 72;
            this.ctrlAbout.Text = "About";
            this.ctrlAbout.UseVisualStyleBackColor = true;
            this.ctrlAbout.Click += new System.EventHandler(this.ctrlAbout_Click);
            // 
            // ctrlCreateXml
            // 
            this.ctrlCreateXml.FlatAppearance.BorderSize = 0;
            this.ctrlCreateXml.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ctrlCreateXml.Location = new System.Drawing.Point(0, 1);
            this.ctrlCreateXml.Margin = new System.Windows.Forms.Padding(0);
            this.ctrlCreateXml.Name = "ctrlCreateXml";
            this.ctrlCreateXml.Size = new System.Drawing.Size(115, 21);
            this.ctrlCreateXml.TabIndex = 72;
            this.ctrlCreateXml.Text = "Save as Xml plan...";
            this.ctrlCreateXml.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ctrlCreateXml.UseVisualStyleBackColor = true;
            this.ctrlCreateXml.Click += new System.EventHandler(this.ctrlCreateXml_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.ctrlCreateXml);
            this.panel1.Controls.Add(this.ctrlAbout);
            this.panel1.Controls.Add(this.ctrlHelp);
            this.panel1.Location = new System.Drawing.Point(-2, -2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(497, 26);
            this.panel1.TabIndex = 73;
            // 
            // FrmCompare
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(496, 432);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.ctrlBrowse);
            this.Controls.Add(this.ctrlFolder);
            this.Controls.Add(this.btSelFolderB);
            this.Controls.Add(this.btSelFolderA);
            this.Controls.Add(this.cbOptValue);
            this.Controls.Add(this.cbOptShape);
            this.Controls.Add(this.cbOptStyle);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.btCompare);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.pictureBox5);
            this.Controls.Add(this.pictureBox4);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lsFilesB);
            this.Controls.Add(this.btSelectB);
            this.Controls.Add(this.lsFilesA);
            this.Controls.Add(this.btClearB);
            this.Controls.Add(this.btSelectA);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.btClearA);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "FrmCompare";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel Batch Compare";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btSelectA;
        private System.Windows.Forms.Button btClearA;
        private System.Windows.Forms.ListBox lsFilesA;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Button btClearB;
        private System.Windows.Forms.Button btSelectB;
        private System.Windows.Forms.ListBox lsFilesB;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btCompare;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.CheckBox cbOptStyle;
        private System.Windows.Forms.PictureBox pictureBox4;
        private System.Windows.Forms.CheckBox cbOptShape;
        private System.Windows.Forms.CheckBox cbOptValue;
        private System.Windows.Forms.Button btSelFolderA;
        private System.Windows.Forms.Button btSelFolderB;
        private System.Windows.Forms.PictureBox pictureBox5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox ctrlFolder;
        private System.Windows.Forms.Button ctrlBrowse;
        private System.Windows.Forms.Button ctrlHelp;
        private System.Windows.Forms.Button ctrlAbout;
        private System.Windows.Forms.Button ctrlCreateXml;
        private System.Windows.Forms.Panel panel1;
    }
}

