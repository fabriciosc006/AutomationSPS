using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiggaPS.tests.dataBaseSAP.Schedule
{
    class Schedule_LoperExcel
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
            XlsxInput.Cells[position, 11].Value = row.EARL_SCH_START_D;
            XlsxInput.Cells[position, 12].Value = row.EARL_SCH_START_T;
        }
        public void Save()
        {
            ExcelIn.Save();
        }
    }
}
    
