using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiggaPS.tests.dataBaseSAP.IndividualCapacity
{
    class IndividualCapacity_IndivCapExcel
    {
        public static string calculatorResult { get; set; }
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
            XlsxInput.Cells[1, 1].Value = "PERNR";
            XlsxInput.Cells[1, 2].Value = "ENAME";
            XlsxInput.Cells[1, 3].Value = "START_DATE";
            XlsxInput.Cells[1, 4].Value = "FINISH_DATE";
            XlsxInput.Cells[1, 5].Value = "START_TIME";
            XlsxInput.Cells[1, 6].Value = "FINISH_TIME";
            XlsxInput.Cells[1, 7].Value = "CAPACITY";
            XlsxInput.Cells[1, 8].Value = "NOTE";
        }
        public void AddCell(int position, IndividualCapacity_SapTable row)
        {
            position += 2;
            XlsxInput.Cells[position, 1].Value = row.PERNR;
            XlsxInput.Cells[position, 2].Value = row.ENAME;
            XlsxInput.Cells[position, 3].Value = convertData(row.START_DATE);
            XlsxInput.Cells[position, 4].Value = convertData(row.FINISH_DATE);
            XlsxInput.Cells[position, 5].Value = convertTime(row.START_TIME);
            XlsxInput.Cells[position, 6].Value = convertTime(row.FINISH_TIME);
            calculaTime(convertTime(row.START_TIME), convertTime(row.FINISH_TIME));
            XlsxInput.Cells[position, 7].Value = row.CAPACITY+calculatorResult;
            XlsxInput.Cells[position, 8].Value = row.NOTE;
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
            string formatingTime = cutMinuts;
            return formatingTime;
        }
        public static string calculaTime(string startTime, string finishTime)
        {
            DateTime dtStart = DateTime.Parse(startTime);
            DateTime dtEnd = DateTime.Parse(finishTime);
            string formatingDate = (dtEnd - dtStart).TotalHours.ToString("0.00");
            formatingDate = formatingDate.Remove(formatingDate.Length - 1, 1) + "0";
            calculatorResult = formatingDate;
            return formatingDate;
        }
    }
}
