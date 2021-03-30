using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiggaPS.tests.dataBaseSAP.Schedule
{
    class Schedule_BLPlanExcel
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
            XlsxInput.Cells[1, 1].Value = "PLANNING_ID";
            XlsxInput.Cells[1, 2].Value = "DESCRICAO";
            XlsxInput.Cells[1, 3].Value = "DATA_INICIAL";
            XlsxInput.Cells[1, 4].Value = "DATA_FINAL";
            XlsxInput.Cells[1, 5].Value = "VALIDADE_DE";
            XlsxInput.Cells[1, 6].Value = "VALIDADE_ATE";
        }
        public void AddCell(int position, Schedule_SapTable row)
        {
            string dataInicial = row.DATA_INICIAL;
            string dataFinal = row.DATA_FINAL;
            string validadeDe = row.VALIDADE_DE;
            string validadeAte = row.VALIDADE_ATE;
            string converteDataInicial = dataInicial.Substring(0, 4).ToString() + '/' + dataInicial.Substring(4, 2).ToString() + '/' + dataInicial.Substring(6, 2);
            string converteDataFinal = dataFinal.Substring(0, 4).ToString() + '/' + dataFinal.Substring(4, 2).ToString() + '/' + dataFinal.Substring(6, 2);
            string converteValidadeDe = validadeDe.Substring(0, 4).ToString() + '/' + validadeDe.Substring(4, 2).ToString() + '/' + validadeDe.Substring(6, 2);
            string converteValidadeAte = validadeAte.Substring(0, 4).ToString() + '/' + validadeAte.Substring(4, 2).ToString() + '/' + validadeAte.Substring(6, 2);
            string convertidoDataInicial = DateTime.Parse(converteDataInicial).ToString("dd/MM/yyyy");
            string convertidoDataFinal = DateTime.Parse(converteDataFinal).ToString("dd/MM/yyyy");
            string convertidoValidadeDe = DateTime.Parse(converteValidadeDe).ToString("dd/MM/yyyy");
            string convertidoValidadeAte = DateTime.Parse(converteValidadeAte).ToString("dd/MM/yyyy");

            position += 2;
            XlsxInput.Cells[position, 1].Value = row.PLANNING_ID;
            XlsxInput.Cells[position, 2].Value = row.DESCRICAO;
            XlsxInput.Cells[position, 3].Value = convertidoDataInicial;
            XlsxInput.Cells[position, 4].Value = convertidoDataFinal;
            XlsxInput.Cells[position, 5].Value = convertidoValidadeDe;
            XlsxInput.Cells[position, 6].Value = convertidoValidadeAte;
        }
        public void Save()
        {
            ExcelIn.Save();
        }

    }
}

