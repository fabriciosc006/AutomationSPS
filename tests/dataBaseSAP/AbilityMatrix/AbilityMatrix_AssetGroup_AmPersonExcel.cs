using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiggaPS.tests.dataBaseSAP.AbilityMatrix
{
    class AbilityMatrix_AssetGroup_AmPersonExcel
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

            XlsxInput.Cells[1, 1].Value = "AM_KEY_TYPE";
            XlsxInput.Cells[1, 2].Value = "AM_KEY";
            XlsxInput.Cells[1, 3].Value = "EQKTX";
            XlsxInput.Cells[1, 4].Value = "PLTXT";
            XlsxInput.Cells[1, 4].Value = "AM_KN";

        }
        public void AddCell(int position, AbilityMatrix_SapTable row)
        {

            position += 2;
            XlsxInput.Cells[position, 1].Value = row.AM_KEY_TYPE;
            XlsxInput.Cells[position, 2].Value = row.AM_KEY;
            XlsxInput.Cells[position, 3].Value = row.EQKTX;
            XlsxInput.Cells[position, 4].Value = row.PLTXT;
            XlsxInput.Cells[position, 5].Value = row.AM_KN;
            
        }
        public void Save()
        {
            ExcelIn.Save();
        }


    }
}
