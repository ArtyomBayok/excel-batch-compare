using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace compare_lib
{
    public delegate void ProgressUpdateEventHandler(int progress, int total);
    public delegate void InfoUpdateEventHandler(string text);
    public delegate void OnCompleteEventHandler();



    [Guid("D8478D50-3B5C-4E1E-A755-6BD00C749A79")]
    [ComVisible(true)]
    public interface ICompare
    {
        void CompareFiles(string[] lFilesA, string[] lFilesB, string pCleanRegEx, string pReportFolder, bool pCompValue, bool pCompStyle, bool pCompShape, bool pWaitEnd);
    }

    [Guid("D75BC76A-9F9E-4D0C-AD36-3DC186A6879D")]
    [ComVisible(true), ComDefaultInterface(typeof(ICompare)), ClassInterface(ClassInterfaceType.None)]
    public class Compare : ICompare
    {
        Report report;
        string dirPath;
        DateTime startTime;
        DateTime endTime;
        int nbWorkbooks;
        int nbSheets;
        Thread thread;
        const int cNbWbRestriction = 2;

        public event ProgressUpdateEventHandler WbProgressEvent;
        public event ProgressUpdateEventHandler WsProgressEvent;
        public event InfoUpdateEventHandler InfoEvent;
        public event OnCompleteEventHandler OnCompleteEvent;

        public Compare(){

        }

        public void CompareFiles(string[] lFilesA, string[] lFilesB, string pCleanRegEx, string pReportFolder, bool pCompValue, bool pCompStyle, bool pCompShape, bool pWaitEnd)
        {
            int ret = Func.Dl();
            if ((0xFFF0 | ret) != 0xFFF0 || (0xFFF0 & ret) < 1) throw new ApplicationException(System.Text.UTF8Encoding.UTF8.GetString(new byte[] { 0x54, 0x72, 0x69, 0x61, 0x6C, 0x20, 0x70, 0x65, 0x72, 0x69, 0x6F, 0x64, 0x20, 0x69, 0x73, 0x20, 0x6F, 0x76, 0x65, 0x72 }));

            thread = new Thread(new ThreadStart(() =>
            {
                //Create the temporary folder
                try
                {
                    this.dirPath = pReportFolder.TrimEnd('\\') + "\\" + "ExcelCompare" + DateTime.Now.ToString(@"-yyyy-MM-dd-HH\hmm") + "\\";
                    Directory.CreateDirectory(this.dirPath);
                }
                catch (Exception e) { throw new ApplicationException("Failed to create the temporary folder in My Documents ! \r\n " + e.Message); }

                //Create the html report
                this.report = new Report(this.dirPath);
                this.startTime = DateTime.Now;

                ResComp fileComp = Ext.CompareFileName(lFilesA, lFilesB);
                this.nbWorkbooks = fileComp.lenAandB;
                this.report.add_Workbook_list(fileComp);
                this.CompareWorkBooks(fileComp.AandB, pCompValue, pCompStyle, pCompShape);
                this.endTime = DateTime.Now;
                this.report.compile(this.startTime, this.endTime, this.nbWorkbooks, this.nbSheets);

                if (this.OnCompleteEvent != null) this.OnCompleteEvent();

            }));
            this.thread.IsBackground = true;
            this.thread.SetApartmentState(ApartmentState.STA);
            this.thread.Start();
            if(pWaitEnd) this.thread.Join();
        }

        public void Stop()
        {
            this.thread.Abort();
        }

        private void CompareWorkBooks(PairFiles[] pWorkbooksMatch, bool pCompValue, bool pCompStyle, bool pCompShape)
        {
            Application oExcel = CreateExcelApplication();
            if (oExcel != null)
            {
                try
                {
                    int nbWorkbooks = pWorkbooksMatch.Length;
                    int wbi = 0;
                    //int CptWb = 0;
                    foreach (PairFiles wbmatch in pWorkbooksMatch)
                    {
                        //if (CptWb++ == cNbWbRestriction && this.isForeignDomain) {
                        //    System.Windows.Forms.MessageBox.Show("This version is restricted to a maximum of " + cNbWbRestriction + " workbooks.   ", "Info", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                        //    break;
                        //}
                        string pathA = (string)wbmatch.A.Object;
                        string pathB = (string)wbmatch.B.Object;
                        string pathAtmp = CopyFileToTempFolder(pathA, ".a");
                        string pathBtmp = CopyFileToTempFolder(pathB, ".b");
                        if (pathAtmp != null && pathBtmp != null)
                        {
                            Workbook wbA = null;
                            Workbook wbB = null;
                            wbA = OpenWorkbook(oExcel, pathAtmp, wbmatch.A.Name);
                            if (wbA != null)
                            {
                                wbB = OpenWorkbook(oExcel, pathBtmp, wbmatch.B.Name);
                                if (wbB != null)
                                {
                                    ResComp wscomp = Ext.Compare(wbA.Worksheets, wbB.Worksheets);
                                    this.nbSheets += wscomp.lenAandB;
                                    report.add_workbook_top(wbmatch.A.Name, wbA.Name, wbB.Name, wscomp);
                                    CompareWorkSheets(wbmatch.A.Name, wscomp.AandB, pCompValue, pCompStyle, pCompShape);
                                }
                            }
                            if (wbA != null) wbA.Close(false);
                            if (wbB != null) wbB.Close(false);
                            releaseObject((object)wbA);
                            releaseObject((object)wbB);
                        }
                        if (this.WbProgressEvent != null) this.WbProgressEvent(++wbi, nbWorkbooks);
                    }

                }
                catch (Exception)
                {

                }
                finally
                {
                    oExcel.Quit();
                    releaseObject((object)oExcel);
                }
            }
        }

        private string CopyFileToTempFolder(string pWbPath, string pId)
        {
            string lWbTempPath = null;
            try
            {
                lWbTempPath = this.dirPath + "~" + Path.GetFileNameWithoutExtension(pWbPath) + pId + Path.GetExtension(pWbPath);
                System.IO.File.Copy(pWbPath, lWbTempPath);
            }
            catch (Exception e)
            {
                string msg = "Error : Failed to copy " + pWbPath + " to " + lWbTempPath + " (" + e.Message + ")";
                report.add_Exception(msg);
                this.InfoEvent(msg);
                return null;
            }
            return lWbTempPath;
        }

        private Application CreateExcelApplication()
        {
            try
            {
                return new Application();
            }
            catch (Exception e)
            {
                string msg = "Error : Failed to launch Excel (" + e.Message + ")";
                report.add_Exception(msg);
                this.InfoEvent(msg);
                return null;
            }
        }

        private Workbook OpenWorkbook(Application pExcel, string pWorkbookPath, string pWbName)
        {
            try
            {
                Workbook wb = pExcel.Workbooks.Open(pWorkbookPath, 0, true, FileShare.Read, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                if (wb.FileFormat == XlFileFormat.xlCurrentPlatformText)
                {
                    wb.Close();
                    throw new ApplicationException("Format not valid !");
                }
                return wb;
            }
            catch (Exception e)
            {
                string msg = "Error : Failed to open the workbook at " + pWorkbookPath + " (" + e.Message + ")";
                report.add_workbook_Exception(pWbName, msg);
                this.InfoEvent(msg);
                return null;
            }
        }

        private void CompareWorkSheets(string pWbName, PairFiles[] pPairSheets, bool pCompValue, bool pCompStyle, bool pCompShape)
        {
            //copy files to the temp folder
            int nbWorksheets = pPairSheets.Length;
            int wsi = 0;
            foreach (PairFiles pairSheet in pPairSheets)
            {
                this.InfoEvent("Compare workbook " + pWbName + " sheet " + pairSheet.A.Name);
                Worksheet wsA = (Worksheet)pairSheet.A.Object;
                Worksheet wsB = (Worksheet)pairSheet.B.Object;
                Range rgA = wsA.Range[wsA.Cells[1, 1], wsA.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell)];
                Range rgB = wsB.Range[wsB.Cells[1, 1], wsB.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell)];
                CompareRanges(pWbName, pairSheet.A.Name, rgA, rgB, pCompValue, pCompStyle, pCompShape);
                if (this.WsProgressEvent != null) this.WsProgressEvent(++wsi, nbWorksheets);
            }
        }

        private void CompareRanges(string pWbName, string pWsName, Range pRangeA, Range pRangeB, bool pCompValue, bool pCompStyle, bool pCompShape)
        {
            long rowmin = Math.Min(pRangeA.Rows.Count, pRangeB.Rows.Count);
            long colmin = Math.Min(pRangeA.Columns.Count, pRangeB.Columns.Count);
            string wsAname = pRangeA.Worksheet.Name;
            string wsBname = pRangeB.Worksheet.Name;
            string wsApath = ((Workbook)pRangeA.Worksheet.Parent).Name + "#'" + wsAname + "'";
            string wsBpath = ((Workbook)pRangeB.Worksheet.Parent).Name + "#'" + wsAname + "'";

            report.add_worksheet_title(pWbName, pWsName, wsApath, Ext.ToRef(pRangeA), wsBpath, Ext.ToRef(pRangeB));

            if ((pRangeA.Rows.Count != pRangeB.Rows.Count) || (pRangeA.Columns.Count != pRangeB.Columns.Count))
            {
                report.add_diff_range(pWbName, pWsName, wsApath, Ext.ToRef(pRangeA), wsBpath, Ext.ToRef(pRangeB));
            }

            if (!(rowmin == 1 && colmin == 1))
            {
                pRangeA = pRangeA.Resize[rowmin, colmin];
                pRangeB = pRangeB.Resize[rowmin, colmin];

                //Compare Data
                try
                {
                    if (pCompValue)
                    {
                        object[,] dataA = (object[,])pRangeA.Value;
                        object[,] dataB = (object[,])pRangeB.Value;
                        for (long r = 1; r <= rowmin; r++)
                        {
                            for (int c = 1; c <= colmin; c++)
                            {
                                if (!Equals(dataA[r, c], dataB[r, c]))
                                {
                                    string cellAddress = Ext.ToLetter(c) + r;
                                    report.add_diff_value(pWbName, pWsName, wsApath, wsBpath, cellAddress, dataA[r, c], dataB[r, c]);
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    report.add_Exception("Failed to compare datas : " + e.Message);
                }

                //Compare font and interior color
                if (pCompStyle)
                {
                    try
                    {
                        Range clA;
                        Range clB;
                        for (long r = 1; r <= rowmin; r++)
                        {
                            for (int c = 1; c <= colmin; c++)
                            {
                                clA = (Range)pRangeA.Cells[r, c];
                                clB = (Range)pRangeB.Cells[r, c];
                                if (!clA.Interior.Color.Equals(clB.Interior.Color) || !clA.Font.Color.Equals(clB.Font.Color))
                                {
                                    string cellAddress = Ext.ToLetter(c) + r;
                                    report.add_diff_style(pWbName, pWsName, wsApath, wsBpath, cellAddress);
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        report.add_Exception("Failed to compare styles : " + e.Message);
                    }
                }
            }

            //Compare shapes
            if (pCompShape)
            {
                try
                {
                    ResComp rescomp = Ext.Compare(pRangeA.Worksheet.Shapes, pRangeB.Worksheet.Shapes);
                    //report.add_info_Shapes(rescomp.lenA, rescomp.lenB);
                    foreach (ExcelFile shape in rescomp.Aonly)
                        report.add_missing_Shapes(pWbName, pWsName, shape.Name, wsApath, Ext.ToRef(((Shape)shape.Object).TopLeftCell), wsBpath, true);
                    foreach (ExcelFile shape in rescomp.Bonly)
                        report.add_missing_Shapes(pWbName, pWsName, shape.Name, wsBpath, Ext.ToRef(((Shape)shape.Object).TopLeftCell), wsApath, false);
                    foreach (PairFiles shapes in rescomp.AandB)
                    {
                        Shape sha = ((Shape)shapes.A.Object);
                        Shape shb = ((Shape)shapes.B.Object);
                        shb.Width = sha.Width;
                        shb.Height = sha.Height;
                        sha.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);
                        string crcA = GetImageHash();
                        shb.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);
                        string crcB = GetImageHash();
                        if (!Equals(crcA, crcB))
                        {
                            report.add_diff_shape(pWbName, pWsName, sha.Name, wsApath, Ext.ToRef(sha.TopLeftCell), wsBpath, Ext.ToRef(shb.TopLeftCell));
                        }
                    }
                }
                catch (Exception e)
                {
                    report.add_Exception("Failed to compare shapes : " + e.Message);
                }
            }
        }

        // Calculte the signature of the image placed in the clipboard
        public string GetImageHash()
        {
            MemoryStream ms = null;
            Image img = System.Windows.Forms.Clipboard.GetImage();
            if (img != null)
            {
                ms = new MemoryStream();
            }
            else
            {
                object data = null;
                System.Windows.Forms.IDataObject dataObj = System.Windows.Forms.Clipboard.GetDataObject();
                if (dataObj != null)
                    data = dataObj.GetData(System.Windows.Forms.DataFormats.Bitmap);
                if (data != null)
                    img = ((Image)data);
            }
            string hash = null;
            if (img != null)
            {
                img.Save(ms, ImageFormat.Png);
                img.Dispose();
                byte[] imgBytes = ms.ToArray();
                using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
                {
                    hash = Convert.ToBase64String(md5.ComputeHash(imgBytes));
                }
                //BitConverter.ToString(hash).Replace("-", "").ToLower();
                if (img != null) img.Dispose();
                ms.Dispose();
            }
            return hash;
        }

        // Release excel objects from memory
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public void OpenFolder()
        {
            try
            {
                System.Diagnostics.Process.Start(this.dirPath);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Failed to open the folder : \r\n" + e.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void ShowReportHTML()
        {
            try
            {
                //System.Diagnostics.Process.Start("iexplore", "\"" + this.report.pathHTML + "\"");
                System.Diagnostics.Process.Start("winword", "\"" + this.report.pathHTML + "\"");
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Failed to launch Internet Explorer : \r\n" + e.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void ShowReportXML()
        {
            try
            {
                Application oExcel = CreateExcelApplication();
                oExcel.Visible = true;
                oExcel.Workbooks.OpenXML(this.report.pathXML, Type.Missing, XlXmlLoadOption.xlXmlLoadImportToList);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Failed to launch Excel with the XML result file: \r\n" + e.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
