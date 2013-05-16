using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace compare_exe
{
    public partial class FrmProgress : Form
    {
        MethodInvoker cancelCompare;
        MethodInvoker exploreFolder;
        MethodInvoker openReportHTML;
        MethodInvoker openReportXML;

        public FrmProgress(MethodInvoker pCancelComapre, MethodInvoker pExploreFolder, MethodInvoker pOpenReportHTML, MethodInvoker pOpenReportXML)
        {
            InitializeComponent();
            this.cancelCompare = pCancelComapre;
            this.exploreFolder = pExploreFolder;
            this.openReportHTML = pOpenReportHTML;
            this.openReportXML = pOpenReportXML;
        }

        private void btCancel_Click(object sender, EventArgs e){
            cancelCompare();
            this.Close();
        }

        private void btOpen_Click(object sender, EventArgs e){
            exploreFolder();
        }

        public void UpdateWbProgress(int pProgress, int pTotal){
            if (this.pbWorkbooks.InvokeRequired){
                this.Invoke( (MethodInvoker)delegate{ this.UpdateWbProgress( pProgress, pTotal); });
            }else{
                this.pbWorkbooks.Value = pProgress * 100 / pTotal;
                this.lbWB.Text = pProgress + " / " + pTotal;
            }
        }

        public void UpdateWsProgress(int pProgress, int pTotal){
            if (this.pbWorksheets.InvokeRequired){
                this.Invoke( (MethodInvoker)delegate{ this.UpdateWsProgress( pProgress, pTotal); });
            }else{
                this.pbWorksheets.Value = pProgress * 100 / pTotal;
                this.lbWS.Text = pProgress + " / " + pTotal;
            }
        }
        
        public void UpdateInfo(string pInfo){
            if (this.txtInfo.InvokeRequired){
                this.Invoke(new Action<string>(this.UpdateInfo), new object[] { pInfo });
            }else{
                this.txtInfo.Text += pInfo + "\r\n";
            }
        }

        public void UpdateFinished(){
            if (this.btOpen.InvokeRequired){
                this.Invoke((MethodInvoker)delegate { this.UpdateFinished(); });
            }else{
                this.btOpen.Enabled = true;
                this.btReportHTML.Enabled = true;
                this.btReportXML.Enabled = true;
            }
        }

        private void btReportHTML_Click(object sender, EventArgs e)
        {
            openReportHTML();
        }

        private void btReportXML_Click(object sender, EventArgs e)
        {
            openReportXML();
        }

    }
}
