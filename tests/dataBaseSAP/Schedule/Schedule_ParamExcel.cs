using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiggaPS.tests.dataBaseSAP.Schedule
{
    class Schedule_ParamExcel
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
            XlsxInput.Cells[1, 1].Value = "FIRST_DAY_WEEK";
            XlsxInput.Cells[1, 2].Value = "SCHED_VALID_TIME";
        }
        public void AddCell(int position, Schedule_SapTable row)
        {
       
            position += 2;
            XlsxInput.Cells[position, 1].Value = row.FIRST_DAY_WEEK;
            XlsxInput.Cells[position, 2].Value = row.SCHED_VALID_TIME;
        }
        public void Save()
        {
            ExcelIn.Save();
        }
    }
}
