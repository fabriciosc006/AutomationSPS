using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechTalk.SpecFlow;

namespace SiggaPS.tests.dataBaseSAP.Schedule
{
    class Schedule_BLOperExcel
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
            XlsxInput.Cells[1, 1].Value = "ORDERID";
            XlsxInput.Cells[1, 2].Value = "ACTIVITY";
            XlsxInput.Cells[1, 3].Value = "SUB_ACTIVITY";
            XlsxInput.Cells[1, 4].Value = "CONTROL_KEY";
            XlsxInput.Cells[1, 5].Value = "WORK_CNTR";
            XlsxInput.Cells[1, 6].Value = "PLANT";
            XlsxInput.Cells[1, 7].Value = "DESCRIPTION";
            XlsxInput.Cells[1, 8].Value = "SYSTCOND";
            XlsxInput.Cells[1, 9].Value = "FUNCLOC";
            XlsxInput.Cells[1, 10].Value = "EQUIPMENT";
            XlsxInput.Cells[1, 11].Value = "EARL_SCH_START_D ";
            XlsxInput.Cells[1, 12].Value = "EARL_SCH_START_T ";
        }
        public void AddCell(int position, Schedule_SapTable row)
        {
            position += 2;
            XlsxInput.Cells[position, 1].Value = row.ORDERID;
            XlsxInput.Cells[position, 2].Value = row.ACTIVITY;
            XlsxInput.Cells[position, 3].Value = row.SUB_ACTIVITY;
            XlsxInput.Cells[position, 4].Value = row.CONTROL_KEY;
            XlsxInput.Cells[position, 5].Value = row.WORK_CNTR;
            XlsxInput.Cells[position, 6].Value = row.PLANT;
            XlsxInput.Cells[position, 7].Value = row.DESCRIPTION;
            XlsxInput.Cells[position, 8].Value = row.SYSTCOND;
            XlsxInput.Cells[position, 9].Value = row.FUNCLOC;
            XlsxInput.Cells[position, 10].Value = row.EQUIPMENT;
            if (FeatureContext.Current.FeatureInfo.Title.Equals("Schedule - Move_Operations"))
            {
                XlsxInput.Cells[position, 11].Value = convertData(row.EARL_SCH_START_D);
                XlsxInput.Cells[position, 12].Value = convertTime(row.EARL_SCH_START_T);
            }
            else
            {
                XlsxInput.Cells[position, 11].Value = row.EARL_SCH_START_D;
                XlsxInput.Cells[position, 12].Value = row.EARL_SCH_START_T;
            }
            
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
            string convert = String.Format("{0:dd/MM/yy}", myDate);
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
