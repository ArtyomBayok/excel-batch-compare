using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using compare_lib;

namespace compare_exe
{
    public partial class About : Form
    {
        public About()
        {
            Application.EnableVisualStyles();
            InitializeComponent();
            this.txtVersion.Text = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            SetLabel();
            new ToolTip().SetToolTip(this.btAddKey, "Set a new license key");
        }

        private void SetLabel()
        {
            this.txtProductId.Text = Func.id();
            RegistryKey or = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\0" + Func.id(), true); if (or != null)
            {
                this.txtLicense.Text = (string)or.GetValue("l");
                or.Close();
                int dl = Func.Dl();
                if(dl==0x09280){
                    this.txtStatus.ForeColor = System.Drawing.Color.Blue;
                    this.txtStatus.Text =  "Unlimited license on the current domain";
                }else if((0xFFF0|dl)==0xFFF0 && (0xFFF0&dl)>0){
                    this.txtStatus.ForeColor = System.Drawing.Color.Blue;
                    this.txtStatus.Text = "This license will expire in " + (dl>>4) + " days ! ";
                }else{
                    this.txtStatus.ForeColor = System.Drawing.Color.Red;
                    this.txtStatus.Text =  "This license has expired ! "; 
                }
            }
        }

        private void btOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public new void Show()
        {
            this.ShowDialog();
        }

        private void btAddKey_Click(object sender, EventArgs e)
        {
            string res= null;
            if( InputBox.Show("License key", "Enter your license key : ",ref res) == System.Windows.Forms.DialogResult.OK){
                int ret = Func.sDl(res);
                if((0x0FFF0|ret)!=0xFFF0 || (0xFFF0&ret)<1 ){
                    MsgBox.ShowWarn("This license key is incorrect !    ");
                    return;
                }else{
                    SetLabel();
                }
            }

        }

    }
}
