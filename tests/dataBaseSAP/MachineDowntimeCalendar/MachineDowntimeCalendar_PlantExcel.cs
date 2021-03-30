using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiggaPS.tests.dataBaseSAP.MachineDowntimeCalendar
{
    class MachineDowntimeCalendar_PlantExcel
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
            XlsxInput.Cells[1, 1].Value = "USERID";
            XlsxInput.Cells[1, 2].Value = "PLANT";
        }
        public void AddCell(int position, MachineDowntimeCalendar_SapTable row)
        {
            position += 2;
            XlsxInput.Cells[position, 1].Value = row.USERID;
            XlsxInput.Cells[position, 2].Value = row.PLANT;          
        }
        public void Save()
        {
            ExcelIn.Save();
        }
    }
}

