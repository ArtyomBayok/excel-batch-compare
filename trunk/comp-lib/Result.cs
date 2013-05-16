using System.Data;

namespace compare_lib
{
    internal class Result
    {
        private DataTable table;

        public Result(){
            table = new DataTable();
            table.TableName = "Differences";
            table.Columns.Add("Difference", typeof(string));
            table.Columns.Add("Workbook", typeof(string));
            table.Columns.Add("WorkSheet", typeof(string));
            table.Columns.Add("RefA", typeof(string));
            table.Columns.Add("RefB", typeof(string));
            table.Columns.Add("DataA", typeof(string));
            table.Columns.Add("DataB", typeof(string));
        }
        
        public void AddDifference(string pDiffType, string pWorkbook, string pWorkSheet, string pRefA, string pRefB, string pDataA, string pDataB  ){
            table.Rows.Add(pDiffType, pWorkbook, pWorkSheet, pRefA, pRefB, pDataA, pDataB);
        }

        public void ToXML(string pDestination ){
            table.WriteXml(pDestination);
        }

    }
}
