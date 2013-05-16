using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace compare_lib
{
    public static class Ext
    {
        public static ResComp Compare(ExcelFile[] a, ExcelFile[] b)
        {
            List<PairFiles> AandB = new List<PairFiles>();
            List<ExcelFile> Aonly = new List<ExcelFile>();
            List<ExcelFile> Bonly = new List<ExcelFile>();
            int lenA = a.GetLength(0);
            int lenB = b.GetLength(0);
            bool[] foundB = new bool[lenB];
            int j;
            for (int i = 0; i < lenA; i++){
                for (j = 0; j < lenB; j++){
                    if (foundB[j] == false){
                        if (a[i].Name == b[j].Name){
                            AandB.Add(new PairFiles { A = a[i], B = b[j] });
                            foundB[j] = true;
                            break;
                        }
                    }
                }
                if (j == lenB) Aonly.Add(a[i]);
            }
            for (j = 0; j < lenB; j++){
                if (foundB[j] == false) Bonly.Add(b[j]);
            }
            ResComp rescomp = new ResComp();
            rescomp.lenA = lenA;
            rescomp.lenB = lenB;
            rescomp.lenAandB = AandB.Count;
            rescomp.lenAonly = Aonly.Count;
            rescomp.lenBonly = Bonly.Count;
            rescomp.AandB = AandB.ToArray();
            rescomp.Aonly = Aonly.ToArray();
            rescomp.Bonly = Bonly.ToArray();
            return rescomp;
        }

        // Compare 2 Shapes collections
        public static ResComp Compare(Shapes pShapesA, Shapes pShapesB)
        {
            List<ExcelFile> lstA = new List<ExcelFile>();
            List<ExcelFile> lstB = new List<ExcelFile>();
            foreach (Shape obj in pShapesA) lstA.Add(new ExcelFile { Name = obj.Name, Object = obj });
            foreach (Shape obj in pShapesB) lstB.Add(new ExcelFile { Name = obj.Name, Object = obj });
            return Ext.Compare(lstA.ToArray(), lstB.ToArray());
        }

        // Compare 2 Sheets collections
        public static ResComp Compare(Sheets pSheetsA, Sheets pSheetsB)
        {
            List<ExcelFile> lstA = new List<ExcelFile>();
            List<ExcelFile> lstB = new List<ExcelFile>();
            foreach (Worksheet obj in pSheetsA) if (obj.Visible == XlSheetVisibility.xlSheetVisible) lstA.Add(new ExcelFile { Name = obj.Name, Object = obj });
            foreach (Worksheet obj in pSheetsB) if (obj.Visible == XlSheetVisibility.xlSheetVisible) lstB.Add(new ExcelFile { Name = obj.Name, Object = obj });
            return Ext.Compare(lstA.ToArray(), lstB.ToArray());
        }

        // Compare 2 file paths arrays
        public static ResComp CompareFileName(string[] pPathA, string[] pPathB)
        {
            List<ExcelFile> lstA = new List<ExcelFile>();
            List<ExcelFile> lstB = new List<ExcelFile>();
            foreach (string obj in pPathA) lstA.Add(new ExcelFile { Name = Path.GetFileNameWithoutExtension(obj), Object = obj });
            foreach (string obj in pPathB) lstB.Add(new ExcelFile { Name = Path.GetFileNameWithoutExtension(obj), Object = obj }); 
            return Ext.Compare(lstA.ToArray(), lstB.ToArray());
        }

        //Convert a number to letters (ex : 1=A, 28=AB)
        public static string ToLetter(int iCol){
            string result = "";
            int iAlpha;
            int iRemainder;
            iAlpha = (iCol / 27);
            iRemainder = iCol - (iAlpha * 26);
            if( iAlpha > 0 )
                result = ((char)(iAlpha + 64)).ToString();
            if( iRemainder > 0 )
                result += ((char)(iRemainder + 64)).ToString();
            return result;
        }

        // Get the range address cleaned of "$" caratere
        public static string ToRef(Range pRange){
            return pRange.AddressLocal.Replace("$","");
        }

    }
}
