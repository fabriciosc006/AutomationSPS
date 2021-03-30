using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiggaPS.tests.dataBaseSAP.MachineDowntimeCalendar
{
    class MachineDowntimeCalendar_Mach_CalExcel
    {
        public static ExcelWorksheet XlsxInput { get; set; }
        public static ExcelPackage ExcelIn { get; set; }
        public void CreateWorkbook(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string pathInput = Path.GetFullPath(path);
            var fileInput = new FileInfo(pathInput);
            if (File.Exists(fileInput.FullName)) { File.Delete(fileInput.FullName); }
            ExcelIn = new ExcelPackage(fileInput);
            XlsxInput = ExcelIn.Workbook.Worksheets.Add("Dados");
            CreateHeader();
        }
        public void CreateHeader()
        {
            XlsxInput.Cells[1, 1].Value = "FUNC_LOC";
            XlsxInput.Cells[1, 2].Value = "PLTXT";
            XlsxInput.Cells[1, 3].Value = "EQUNR";
            XlsxInput.Cells[1, 4].Value = "EQKTX";
            XlsxInput.Cells[1, 5].Value = "START_DATE";
            XlsxInput.Cells[1, 6].Value = "FINISH_DATE";
            XlsxInput.Cells[1, 7].Value = "START_TIME";
            XlsxInput.Cells[1, 8].Value = "FINISH_TIME";
            XlsxInput.Cells[1, 9].Value = "STOP_TYPE";
            XlsxInput.Cells[1, 10].Value = "TIME_CONF";
            XlsxInput.Cells[1, 11].Value = "NOTE";
        }
        public void AddCell(int position, MachineDowntimeCalendar_SapTable row)
        {
            position += 2;
            XlsxInput.Cells[position, 1].Value = row.FUNC_LOC;
            XlsxInput.Cells[position, 2].Value = row.PLTXT;
            XlsxInput.Cells[position, 3].Value = row.EQUNR;
            XlsxInput.Cells[position, 4].Value = row.EQKTX;
            XlsxInput.Cells[position, 5].Value = convertData(row.START_DATE);
            XlsxInput.Cells[position, 6].Value = convertData(row.FINISH_DATE);
            XlsxInput.Cells[position, 7].Value = convertTime(row.START_TIME);
            XlsxInput.Cells[position, 8].Value = convertTime(row.FINISH_TIME);
            XlsxInput.Cells[position, 9].Value = row.STOP_TYPE;
            XlsxInput.Cells[position, 10].Value = row.TIME_CONF;
            XlsxInput.Cells[position, 11].Value = row.NOTE;
        }
        public void Save()
        {
            ExcelIn.Save();
        }
        public static string convertData(string data)
        {
            string excelData = data;
            string cutYear = excelData.Insert(4, "-");
            string cutDay = cutYear.Insert(7, "-");
            string formatingDate = cutDay;
            DateTime myDate = DateTime.ParseExact(formatingDate, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
            string convert = String.Format("{0:dd/MM/yyyy}", myDate);
            return convert;
        }
        public static string convertTime(string time)
        {
            string excelTime = time;
            string cutHour = excelTime.Insert(2, ":");
            string cutMinuts = cutHour.Insert(5, ":");
            string cutSegunds = cutMinuts.Substring(0, 5);
            string formatingTime = cutSegunds;
            return formatingTime;
        }
    }
}
