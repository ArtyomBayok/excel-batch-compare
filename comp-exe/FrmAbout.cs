using System;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;


namespace compare_exe
{
    public partial class FrmAbout : Form
    {
        public FrmAbout(){
             InitializeComponent();
             //txtVersion.Text = Assembly.GetExecutingAssembly().GetName().Version.ToString();
             txtVersion.Text = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            System.Diagnostics.Process.Start(((LinkLabel)sender).Text);
        }

        private void btOk_Click(object sender, EventArgs e) {
            this.Close();
        }

        private void FrmAbout_Shown(object sender, EventArgs e) {
            btOk.Focus();
        }
    }
}
