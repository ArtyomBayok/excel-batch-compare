using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using compare_lib;

namespace compare_exe
{
    public partial class FrmCompare : Form
    {
        private string DefaultDir;

        public FrmCompare(){
            InitializeComponent();
            this.DefaultDir = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            this.ctrlFolder.Text = this.DefaultDir;
        }

        private void btSelectA_Click(object sender, EventArgs e){
           string[] ret = SelectFiles(ref this.DefaultDir);
           if(ret!=null){
                foreach (string lFileName in ret)
                    this.lsFilesA.Items.Add(lFileName);
           }
        }

        private void btSelectB_Click(object sender, EventArgs e){
           string[] ret = SelectFiles(ref this.DefaultDir);
           if(ret!=null){
                foreach (string lFileName in ret)
                    this.lsFilesB.Items.Add(lFileName);
           }
        }

        private void btSelFolderA_Click(object sender, EventArgs e){
            string[] ret = SelectFolder(ref this.DefaultDir);
            if(ret==null) return;
            foreach (string lFileName in ret)
                this.lsFilesA.Items.Add(lFileName);
        }

        private void btSelFolderB_Click(object sender, EventArgs e){
            string[] ret = SelectFolder(ref this.DefaultDir);
            if (ret == null) return;
            foreach (string lFileName in ret)
                this.lsFilesB.Items.Add(lFileName);
        }

        private string[] SelectFolder(ref string pDefautDirectory){
            using (FolderBrowserDialog dlgOpen = new FolderBrowserDialog()){
                dlgOpen.SelectedPath = pDefautDirectory;
                dlgOpen.Description = "Select a folder";
                if (dlgOpen.ShowDialog() == DialogResult.OK){
                    List<string> files = new  List<string>();
                    foreach (string file in Directory.GetFiles(dlgOpen.SelectedPath)){
                        if (file.EndsWith(".xls") || file.EndsWith(".xlsx")){
                            files.Add(file);
                        }
                    }
                    return files.ToArray();
                }else{
                    return null;
                }
            }
        }

        private string[] SelectFiles(ref string pDefautDirectory){
            using( OpenFileDialog dlgOpen = new OpenFileDialog()){
                dlgOpen.Title = "Excel files selection";
                dlgOpen.ShowReadOnly = true;
                dlgOpen.Multiselect = true;
                dlgOpen.AutoUpgradeEnabled = false;
                dlgOpen.RestoreDirectory = true;
                dlgOpen.Filter = "Excel Workbook (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
                dlgOpen.InitialDirectory = pDefautDirectory;
                if(dlgOpen.ShowDialog() == DialogResult.OK){
                    System.IO.DirectoryInfo directory = new System.IO.FileInfo(dlgOpen.FileNames[0]).Directory;
                    if (directory.Parent != null){
                        pDefautDirectory = directory.Parent.FullName;
                    }else{
                        pDefautDirectory = directory.FullName;
                    }
                    return dlgOpen.FileNames;
                }else{
                    return null;
                }
            }
        }

        private void btClearA_Click(object sender, EventArgs e){
            this.lsFilesA.Items.Clear();
        }

        private void btClearB_Click(object sender, EventArgs e){
            this.lsFilesB.Items.Clear();
        }

        private void btCompare_Click(object sender, EventArgs e){
            if (this.lsFilesA.Items.Count == 0){
                MessageBox.Show("Selection A is empty !  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (this.lsFilesB.Items.Count == 0){
                MessageBox.Show("Selection B is empty !  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //put files path in array
            string[] lstA = new string[this.lsFilesA.Items.Count];
            this.lsFilesA.Items.CopyTo(lstA, 0);
            string[] lstB = new string[this.lsFilesB.Items.Count];
            this.lsFilesB.Items.CopyTo(lstB, 0);

            Compare compare = new Compare();
            
            FrmProgress progress = new FrmProgress(
                (MethodInvoker)delegate() { compare.Stop(); },
                (MethodInvoker)delegate() { compare.OpenFolder(); },
                (MethodInvoker)delegate() { compare.ShowReportHTML(); },
                (MethodInvoker)delegate() { compare.ShowReportXML(); }
            );
            compare.WbProgressEvent += new ProgressUpdateEventHandler(progress.UpdateWbProgress);
            compare.WsProgressEvent += new ProgressUpdateEventHandler(progress.UpdateWsProgress);
            compare.InfoEvent += new InfoUpdateEventHandler(progress.UpdateInfo);
            compare.OnCompleteEvent += new OnCompleteEventHandler(progress.UpdateFinished);
            compare.CompareFiles(lstA, lstB, @".*", this.ctrlFolder.Text, this.cbOptValue.Checked, this.cbOptStyle.Checked, this.cbOptShape.Checked, false);
            progress.ShowDialog();
        }

        private void btCancel_Click(object sender, EventArgs e){
            this.Close();
        }

        private string[] get_listbox_items( ListBox.ObjectCollection pItems){
            string[] lst = new string[pItems.Count];
            for(int i = 0; i < pItems.Count; i++)
                lst[i] = pItems[i].ToString();
            return lst;
        }

        private void lsFilesA_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete){
                for (int i = lsFilesA.SelectedItems.Count - 1; i != -1; i--)
                    lsFilesA.Items.RemoveAt(i);
            }else if(e.KeyCode == Keys.Up){
                int index = this.lsFilesA.SelectedIndex;
                if(index == 0) return;
                object data = this.lsFilesA.SelectedItem;
                this.lsFilesA.Items.RemoveAt(index);
                this.lsFilesA.Items.Insert(index-1, data);
                this.lsFilesA.SelectedIndex = index;
            }else if(e.KeyCode == Keys.Down){
                int index = this.lsFilesA.SelectedIndex;
                if(index == (this.lsFilesA.Items.Count-1)) return;
                object data = this.lsFilesA.SelectedItem;
                this.lsFilesA.Items.RemoveAt(index);
                this.lsFilesA.Items.Insert(index + 1, data);
                this.lsFilesA.SelectedIndex = index;
            }
        }

        private void lsFilesB_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete){
               for(int i = lsFilesB.SelectedItems.Count-1; i!=-1; i--)
                   lsFilesA.Items.RemoveAt(i);
            }else if(e.KeyCode == Keys.Up){
                int index = this.lsFilesB.SelectedIndex;
                if(index == 0) return;
                object data = this.lsFilesB.SelectedItem;
                this.lsFilesB.Items.RemoveAt(index);
                this.lsFilesB.Items.Insert(index-1, data);
                this.lsFilesB.SelectedIndex = index;
            }else if(e.KeyCode == Keys.Down){
                int index = this.lsFilesB.SelectedIndex;
                if(index == (this.lsFilesB.Items.Count-1)) return;
                object data = this.lsFilesB.SelectedItem;
                this.lsFilesB.Items.RemoveAt(index);
                this.lsFilesB.Items.Insert(index + 1, data);
                this.lsFilesB.SelectedIndex = index;
            }
        }

        private void ctrlBrowse_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dlgOpen = new FolderBrowserDialog()){
                dlgOpen.SelectedPath = this.DefaultDir;
                dlgOpen.Description = "Select a folder";
                if (dlgOpen.ShowDialog() == DialogResult.OK){
                    this.ctrlFolder.Text = dlgOpen.SelectedPath;
                }
            }
        }

        private void ctrlCreateXml_Click(object sender, EventArgs e)
        {
            string lFilename=String.Empty;
            using(SaveFileDialog saveDialog = new SaveFileDialog()){
                saveDialog.Filter = "XML FIle|*.xml";
                saveDialog.Title = "Save XML File";
                saveDialog.ShowDialog();
                lFilename = saveDialog.FileName;
            }
            if (lFilename != String.Empty){
                try{
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    string[] lstA = new string[this.lsFilesA.Items.Count];
                    this.lsFilesA.Items.CopyTo(lstA, 0);
                    string[] lstB = new string[this.lsFilesB.Items.Count];
                    this.lsFilesB.Items.CopyTo(lstB, 0);
                    RunnerPlan data = new RunnerPlan{
                        FilesA = lstA,
                        FilesB = lstB,
                        ReportFolder = this.ctrlFolder.Text,
                        CleanRegEx = @".*",
                        CompValue = this.cbOptValue.Checked,
                        CompStyle = this.cbOptStyle.Checked,
                        CompShape = this.cbOptShape.Checked
                    };
                    data.ParseToXml(lFilename);
                    string exePath = Func.GetRunnerPath();
                    MsgBox.Show("Command line :", "\"" + exePath + "\" \"" + lFilename + "\" \"" + lFilename.Replace(".xml", ".log") + "\"");
                }catch(Exception ex){ 
                    throw ex;
                }finally{
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                }
            }

        }

        private void ctrlHomePage_Click(object sender, EventArgs e)
        {
            try{
                System.Diagnostics.Process.Start(Directory.GetParent(System.Reflection.Assembly.GetExecutingAssembly().Location).FullName + @"\help.chm");
            }catch(Exception){}
        }

        private void ctrlAbout_Click(object sender, EventArgs e)
        {
            new FrmAbout().ShowDialog(this);
        }

        private void lsFilesA_DragDrop(object sender, DragEventArgs e)
        {
            if (this.lsFilesA.SelectedItems.Count >1) return;
            Point point = lsFilesA.PointToClient(new Point(e.X, e.Y));
            int index = this.lsFilesA.IndexFromPoint(point);
            if (index < 0) index = this.lsFilesA.Items.Count - 1;
            object data = e.Data.GetData(typeof(String));
            this.lsFilesA.Items.Remove(data);
            this.lsFilesA.Items.Insert(index, data);
            this.lsFilesA.SelectedIndex = index;
        }

        private void lsFilesA_MouseDown(object sender, MouseEventArgs e)
        {
            if (this.lsFilesA.SelectedItem == null) return;
            this.lsFilesA.DoDragDrop(this.lsFilesA.SelectedItem, DragDropEffects.Move);
        }

        private void lsFilesA_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void lsFilesB_DragDrop(object sender, DragEventArgs e)
        {
            if (this.lsFilesB.SelectedItems.Count >1) return;
            Point point = lsFilesB.PointToClient(new Point(e.X, e.Y));
            int index = this.lsFilesB.IndexFromPoint(point);
            if (index < 0) index = this.lsFilesB.Items.Count - 1;
            object data = e.Data.GetData(typeof(String));
            this.lsFilesB.Items.Remove(data);
            this.lsFilesB.Items.Insert(index, data);
            this.lsFilesB.SelectedIndex = index;
        }

        private void lsFilesB_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void lsFilesB_MouseDown(object sender, MouseEventArgs e)
        {
            if (this.lsFilesB.SelectedItem == null) return;
            this.lsFilesB.DoDragDrop(this.lsFilesB.SelectedItem, DragDropEffects.Move);
        }

    }
}
