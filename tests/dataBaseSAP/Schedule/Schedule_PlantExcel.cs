using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiggaPS.tests.dataBaseSAP.Schedule
{
    class Schedule_PlantExcel
    {
        public static ExcelWorksheet XlsxInput { get; set; }
        public static ExcelPackage ExcelIn { get; set; }

        public List<string> plants = new List<string>();

        public string plantsToString()
        {
            return String.Join(",", plants);
        }
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
        public void AddCell(int position, Schedule_SapTable row)
        {
            position += 2;
            XlsxInput.Cells[position, 1].Value = row.USERID;
            XlsxInput.Cells[position, 2].Value = row.PLANT;
            var rowPlanilha = new List<string>();
            for (int rw = 1; rw <= XlsxInput.Dimension.End.Row; rw++)
            {
                if (XlsxInput.Cells[position, 2].Value != null)
                    rowPlanilha.Add(XlsxInput.Cells[position, 2].Value.ToString());
            }
            List<string> distinct = rowPlanilha.Distinct().ToList();
            foreach (string linha in distinct)
            {
                plants.Add("'" + linha + "'");
            }
        }
        public void Save()
        {
            ExcelIn.Save();
        }
    }
}
