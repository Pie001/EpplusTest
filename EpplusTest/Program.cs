using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;

namespace EpplusTest
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //ModifiedExcel();
            //SelectExcel();
            ViewExcel();
        }

        public static void ModifiedExcel()
        {
            var path = string.Format(@"{0}\file.xlsx", TEMP_FOLDER);

            FileInfo file = new FileInfo(path);

            using (var package = new ExcelPackage(file))
            {
                ExcelWorkbook workBook = package.Workbook;
                ExcelWorksheet currentWorksheet = workBook.Worksheets.SingleOrDefault(w => w.Name == "Sheet1");

                int totalRows = currentWorksheet.Dimension.End.Row;
                int totalCols = currentWorksheet.Dimension.End.Column;

                // Formula
                currentWorksheet.Cells["W5"].Formula = "=IFERROR($S$5*$V5/$U$5,0)";
                currentWorksheet.Cells["W6"].Formula = "=IFERROR($S$6*$V6/$U$6,0)";

                package.Save();
            }
        }

        public static void ViewExcel()
        {
            var path = string.Format(@"{0}\file.xlsx", TEMP_FOLDER);
            
            FileInfo file = new FileInfo(path);

            using (var package = new ExcelPackage(file))
            {
                ExcelWorkbook workBook = package.Workbook;
                ExcelWorksheet currentWorksheet = workBook.Worksheets.FirstOrDefault();

                int totalRows = currentWorksheet.Dimension.End.Row;
                int totalCols = currentWorksheet.Dimension.End.Column;

                // get date value from cell
                var ahCell = currentWorksheet.Cells["AH5"];
                // Date Format: value -> datetime
                long dateNum = long.Parse(ahCell.Value.ToString());
                DateTime result = DateTime.FromOADate(dateNum);
                string strResult = result.ToString("yyyy/MM/dd");

                var aiCell = currentWorksheet.Cells["AI5"];
                var ajCell = currentWorksheet.Cells["AJ5"];
                var akCell = currentWorksheet.Cells["AK5"];
                var alCell = currentWorksheet.Cells["AL5"];


                currentWorksheet.Cells["AM7"].Formula = "INT(AK7)";
                currentWorksheet.Cells["AM8"].Formula = "INT(AK8)";
                currentWorksheet.Cells["AM9"].Formula = "INT(AK9)";
                currentWorksheet.Cells["AM10"].Formula = "INT(AK10)";
                currentWorksheet.Cells["AM11"].Formula = "INT(AK11)";
                currentWorksheet.Cells["AM12"].Formula = "INT(AK12)";
                currentWorksheet.Cells["AM13"].Formula = "INT(AK13)";

                // Calculate() invoke level > FILE / Sheet / Cell
                //currentWorksheet.Cells["AM7"].Calculate(); // sample
                currentWorksheet.Calculate();
                var amVal = currentWorksheet.Cells["AM8"].Value;
                var amCell = currentWorksheet.Cells["AM8"];
                var aqCell = currentWorksheet.Cells["AQ3"];


                package.Save();
            }
        }

        public static void SelectExcel()
        {
            var path = string.Format(@"{0}\file.xlsx", TEMP_FOLDER);

            FileInfo file = new FileInfo(path);

            using (var package = new ExcelPackage(file))
            {
                ExcelWorkbook workBook = package.Workbook;
                ExcelWorksheet currentWorksheet = workBook.Worksheets.SingleOrDefault(w => w.Name == "Sheet1");

                int totalRows = currentWorksheet.Dimension.End.Row;
                int totalCols = currentWorksheet.Dimension.End.Column;

                // select 1
                var query =
                    from cell in currentWorksheet.Cells["H4:H558"]
                    where cell.Value?.ToString() == "Exception"
                    select cell;

                foreach(var cell in query)
                {
                    string adr = cell.Address;
                    currentWorksheet.Cells[adr.Replace("H", "W")].Formula = "=iferror($s$1*$v5/$u$1,0)";
                }


                currentWorksheet.Cells["H500"].Value = "TEST";

                // select 2
                var query2 =
                    from cell in currentWorksheet.Cells["H4:H558"]
                    where cell.Value?.ToString() == "TEST"
                    select cell;

                foreach (var cell in query2)
                {
                    string adr = cell.Address;
                    currentWorksheet.Cells[adr.Replace("H", "W")].Formula = "=iferror($s$1111*$v1111/$u$1111,0)";
                }

                package.Save();
            }
        }

        public static string TEMP_FOLDER
        {
            get
            {
                string logFolder = string.Format(@"{0}\{1}", Environment.CurrentDirectory, "Temp");
                if (!Directory.Exists(logFolder)) { Directory.CreateDirectory(logFolder); }
                return logFolder;
            }
        }
    }
}
