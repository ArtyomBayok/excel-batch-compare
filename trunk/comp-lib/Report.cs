using System;
using System.IO;
using System.Text;

namespace compare_lib
{

    class Report
    {
        Result result;
        private StringBuilder body;
        private StringBuilder header;
        private StringBuilder bottom;
        internal string pathHTML;
        internal string pathXML;
        private string directory;
        private long nbDiff;
        private long nbException;

        public Report(string pTempFolder)
        {
            result = new Result();
            this.header = new StringBuilder();
            this.body = new StringBuilder();
            this.bottom = new StringBuilder();
            this.directory = pTempFolder;
            string path = pTempFolder.TrimEnd('\\') + "\\Report-" + DateTime.Now.ToString(@"yyyy-MM-dd-HH\hmm");
            this.pathHTML = path + ".html";
            this.pathXML = path + ".xml";

            header.AppendLine("<META http-equiv=\"content-type\" content=\"text/html; charset=utf-8\" />");
            header.AppendLine("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01//EN\" \"http://www.w3.org/TR/html4/strict.dtd\">");
            header.AppendLine("<HTML>");
            header.AppendLine("<HEAD>");
            header.AppendLine("<TITLE>Compare Report</TITLE>");
            header.AppendLine("<STYLE type=\"text/css\">");
            header.AppendLine("BODY{ font-family:Arial;font-size:11px; } a:active, a:link{ color:darkBlue;text-decoration:none} a:visited{text-decoration:none}");
            header.AppendLine("</STYLE>");
            header.AppendLine("</HEAD>");
            header.AppendLine("<BODY>");
            header.AppendLine("<P align=center><FONT size=\"5\"><b>Excel Compare Report</b></FONT></P>");

            bottom.AppendLine("</BODY>");
            bottom.AppendLine("</HTML>");
        }

        internal void compile(DateTime pStartTime, DateTime pEndTime, int pNbWorkBooks, int pNbWorkSheets)
        {
            result.ToXML(pathHTML.Replace(".html", ".xml"));

            DateTime totalTime = new DateTime(pEndTime.Subtract(pStartTime).Ticks);

            header.AppendLine("<b>Directory</b> : <A href=\""+new Uri(this.directory).AbsolutePath+"\">"+this.directory+"</A><br>");
            header.AppendLine("<b>Start time</b> : " + pStartTime + "<br>");
            header.AppendLine("<b>End time</b> : " + pEndTime + "<br>");
            header.AppendLine("<b>Total time</b> : " + totalTime.ToString("HH:mm:ss") + "<br>");
            header.AppendLine("<b>Result</b> : " + this.nbDiff + " differences found in " + pNbWorkBooks + " workbooks / " + pNbWorkSheets + " worksheets<br>");
            header.AppendLine("<b>Exception</b> : " + this.nbException + "<br>");

            using (TextWriter outfile = new StreamWriter(this.pathHTML))
            {
                outfile.Write(header.ToString());
                outfile.Write(body.ToString());
                outfile.Write(bottom.ToString());
                outfile.Flush();
                outfile.Close();
            }
        }

        private string Link(string pHref, string pText)
        {
            return "<A target=\"_blank\" href=\"" + pHref + "\">" + pText + "</A>";
        }

        internal void add_Workbook_list(ResComp fileComp)
        {
            body.AppendLine("<br><font size=\"2\" ><b>Selection</b></font> <i>( " + fileComp.lenAandB + " workbooks )</i><br>");
            body.AppendLine("<hr NOSHADE size=\"1px\" width=\"75%\" ALIGN=\"left\">");

            foreach (PairFiles wbs in fileComp.AandB)
            {
                body.AppendLine("&nbsp;<b>• Workbook</b> '"+wbs.A.Name+"' : <br>");
                body.AppendLine("&nbsp;&nbsp;&nbsp; Selection A : <a target=\"_blank\" href=\"" + wbs.A.ToUri() + "\">" + wbs.A.Object + " (" + wbs.A.ToSize() + " Ko)</a><br>");
                body.AppendLine("&nbsp;&nbsp;&nbsp; Selection B : <a target=\"_blank\" href=\"" + wbs.B.ToUri() + "\">" + wbs.B.Object + " (" + wbs.B.ToSize() + " Ko)</a><br>");
            }
            foreach (ExcelFile wbs in fileComp.Aonly)
            {
                result.AddDifference("MissingWorkook", wbs.Name, "", " + ", " - ", null, null);
                body.AppendLine("&nbsp;<font color=red><b>• Workbook</b> '" + wbs.Name + "' is missing in selection B : </font><br>");
                body.AppendLine("&nbsp;&nbsp;&nbsp; <font color=red>Selection A : <a target=\"_blank\" href=\""+wbs.ToUri()+"\">"+wbs.Object+" ("+wbs.ToSize()+" Ko)</a></font><br>");
            }
            foreach (ExcelFile wbs in fileComp.Bonly)
            {
                result.AddDifference("MissingWorkook", wbs.Name, "", " - "," + ", null, null);
                body.AppendLine("&nbsp;<font color=red><b>• Workbook</b> '"+wbs.Name+"' is missing in selection A : </font><br>");
                body.AppendLine("&nbsp;&nbsp;&nbsp; <font color=red>Selection B : <a target=\"_blank\" href=\""+wbs.ToUri()+"\">"+wbs.Object+" ("+wbs.ToSize()+" Ko)</a></font><br>");
            }
        }

        internal void add_workbook_top(string wbName, string tmpWbNameA, string tmpWbNameB, ResComp wscomp)
        {
            body.AppendLine(string.Format("<br><font size=\"2\" ><b>Workbook </b> \"{0}\" :</font> <a target=\"_blank\" href=\"{1}\" >Selection A</a> ({2} Sheets) versus <a target=\"_blank\" href=\"{3}\">Selection B</a> ({4} Sheets)<br>", wbName, tmpWbNameA, wscomp.lenA, tmpWbNameB, wscomp.lenB));
            body.AppendLine("<hr NOSHADE size=\"1px\" width=\"75%\" ALIGN=\"left\">");
            foreach (ExcelFile wbs in wscomp.Aonly)
            {
                result.AddDifference("MissingWorsheet", wbName, wbs.Name, " + ", " - ", null, null);
                body.AppendLine("&nbsp;<font color=red><b>• Worksheet </b>'" + wbs.Name + "' is missing in workbook B</font><br>");
            }
            foreach (ExcelFile wbs in wscomp.Bonly)
            {
                result.AddDifference("MissingWorsheet", wbName, wbs.Name, " - ", " + ", null, null);
                body.AppendLine("&nbsp;<font color=red><b>• Worksheet </b>'" + wbs.Name + "' is missing in workbook A</font><br>");
            }
        }

        internal void add_worksheet_title(string wbName, string wsAname, string wsApath, string refA, string wsBpath, string refB)
        {
            result.AddDifference("", wbName, wsAname, refA, refB, null, null);
            body.AppendLine(string.Format("&nbsp;<b>• Worksheet</b> '{0}', Range <a target=\"_blank\" href=\"{1}!{2}\">A[{2}]</a> versus <a target=\"_blank\" href=\"{3}!{4}\">B[{4}]</a><br>", wsAname, wsApath, refA, wsBpath, refB));
        }

        internal void add_diff_range(string pWbName, string pWsName, string wsApath, string refA, string wsBpath, string refB)
        {
            this.nbDiff++;
            result.AddDifference("DiffRange", pWbName, pWsName, refA, refB, null, null);
            body.AppendLine(string.Format("&nbsp;&nbsp;&nbsp;<font color=red><B> - Range</b> <a target=\"_blank\" href=\"{0}!{1}\">A[{1}]</a> != <a target=\"_blank\" href=\"{2}!{3}\">B[{3}]</a></font><br>", wsApath, refA, wsBpath, refB));
        }

        internal void add_diff_value(string pWbName, string pWsName, string wsApath, string wsBpath, string cellref, object valA, object valB)
        {
            this.nbDiff++;
            result.AddDifference("DiffValue", pWbName, pWsName, cellref, cellref, (valA != null ? valA.ToString() : null), (valB != null ? valB.ToString() : null));
            body.AppendLine(string.Format("&nbsp;&nbsp;&nbsp;<font color=red><B> - Value</b> <a target=\"_blank\" href=\"{0}!{2}\">A[{2}]</a> != <a target=\"_blank\" href=\"{1}!{2}\" >B[{2}]</a> : [{3}] != [{4}]</font><br>", wsApath, wsBpath, cellref, valA, valB));
        }

        internal void add_diff_style(string pWbName, string pWsName, string wsApath, string wsBpath, string cellref)
        {
            this.nbDiff++;
            result.AddDifference("DiffStyle", pWbName, pWsName, cellref, cellref, null, null);
            body.AppendLine(string.Format("&nbsp;&nbsp;&nbsp;<font color=red><B> - Style</b> <a target=\"_blank\" href=\"{0}!{2}\">A[{2}]</a> != <a target=\"_blank\" href=\"{1}!{2}\" >B[{2}]</a></font><br>", wsApath, wsBpath, cellref));
        }

        internal void add_missing_Shapes(string pWbName, string pWsName, string shName, string ws1path, string cell1Ref, string ws2path, bool isA)
        {
            this.nbDiff++;
            result.AddDifference("MissingShape", pWbName, pWsName, (isA ? cell1Ref : null), (isA ? null : cell1Ref), null, null);
            body.AppendLine("&nbsp;&nbsp;&nbsp;<font color=red><B> - Shape</b> <a target=\"_blank\" href=\""+ws1path+"!"+cell1Ref+"\">"+(isA?"A":"B")+"["+shName+"]</a> is missing in worksheet <a target=\"_blank\" href=\""+ws2path+"!A1\" >"+(isA?"B":"A")+"</a></font><br>");
        }

        internal void add_diff_shape(string pWbName, string pWsName, string shName, string wsApath, string refA, string wsBpath, string refB)
        {
            this.nbDiff++;
            result.AddDifference("DiffShape", pWbName, pWsName, refA, refB, null, null);
            body.AppendLine(string.Format("&nbsp;&nbsp;&nbsp;<font color=red><B> - Shape </b> <a target=\"_blank\" href=\"{1}!{2}\" >A[{0}!{2}]</a> != <a target=\"_blank\" href=\"{3}!{4}\" >B[{0}!{4}]</a></font><br>", shName, wsApath, refA, wsBpath, refB));
        }

        internal void add_Exception(string pMessage)
        {
            this.nbException++;
            body.AppendLine("&nbsp;&nbsp;&nbsp;<font color=red> Exception : </b>" + pMessage + "</font><br>");
        }
        internal void add_workbook_Exception(string wbName, string pMessage)
        {
            body.AppendLine(string.Format("<br><font size=\"2\" ><b>Workbook </b> \"{0}\"</font>", wbName));
            body.AppendLine(@"<hr NOSHADE size=""1px"" width=""75%"" ALIGN=""left"">");
            body.AppendLine("&nbsp;<font color=red>" + pMessage + "</font><br>");
        }


    }
}
