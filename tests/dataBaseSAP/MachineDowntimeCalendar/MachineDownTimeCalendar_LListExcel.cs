using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiggaPS.tests.dataBaseSAP.MachineDowntimeCalendar
{
    class MachineDownTimeCalendar_LListExcel
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
            XlsxInput.Cells[1, 1].Value = "MANDT";
            XlsxInput.Cells[1, 2].Value = "CENTRO";
            XlsxInput.Cells[1, 3].Value = "COD_ITEM";
            XlsxInput.Cells[1, 4].Value = "NOME_LISTA";
            XlsxInput.Cells[1, 5].Value = "IDIOMA";
            XlsxInput.Cells[1, 6].Value = "TIPONOTAORDEM";
            XlsxInput.Cells[1, 7].Value = "LIST_LEVEL";
            XlsxInput.Cells[1, 8].Value = "ITEM_PARENT";
            XlsxInput.Cells[1, 9].Value = "EXTRA_FLD1";
            XlsxInput.Cells[1, 10].Value = "EXTRA_FLD2";
            XlsxInput.Cells[1, 11].Value = "EXTRA_FLD3";
            XlsxInput.Cells[1, 12].Value = "EXTRA_FLD4";
            XlsxInput.Cells[1, 13].Value = "DESCRICAO";
            XlsxInput.Cells[1, 14].Value = "ITEM_HASH";
            XlsxInput.Cells[1, 15].Value = "DATA_UPD_CDS";
            XlsxInput.Cells[1, 16].Value = "HORA_UPD_CDS";
            XlsxInput.Cells[1, 17].Value = "USUARIO";
            XlsxInput.Cells[1, 18].Value = "CDS_CHANGE_DATE";
            XlsxInput.Cells[1, 19].Value = "CDS_KEY";
        }
        public void AddCell(int position, MachineDowntimeCalendar_SapTable row)
        {
            position += 2;
            XlsxInput.Cells[position, 1].Value = row.MANDT;
            XlsxInput.Cells[position, 2].Value = row.CENTRO;
            XlsxInput.Cells[position, 3].Value = row.COD_ITEM;
            XlsxInput.Cells[position, 4].Value = row.NOME_LISTA;
            XlsxInput.Cells[position, 5].Value = row.IDIOMA;
            XlsxInput.Cells[position, 6].Value = row.TIPONOTAORDEM;
            XlsxInput.Cells[position, 7].Value = row.LIST_LEVEL;
            XlsxInput.Cells[position, 8].Value = row.ITEM_PARENT;
            XlsxInput.Cells[position, 9].Value = row.EXTRA_FLD1;
            XlsxInput.Cells[position, 10].Value = row.EXTRA_FLD2;
            XlsxInput.Cells[position, 11].Value = row.EXTRA_FLD3;
            XlsxInput.Cells[position, 12].Value = row.EXTRA_FLD4;
            XlsxInput.Cells[position, 13].Value = row.DESCRICAO;
            XlsxInput.Cells[position, 14].Value = row.ITEM_HASH;
            XlsxInput.Cells[position, 15].Value = row.DATA_UPD_CDS;
            XlsxInput.Cells[position, 16].Value = row.HORA_UPD_CDS;
            XlsxInput.Cells[position, 17].Value = row.USUARIO;
            XlsxInput.Cells[position, 18].Value = row.CDS_CHANGE_DATE;
            XlsxInput.Cells[position, 19].Value = row.CDS_KEY;
        }
        public void Save()
        {
            ExcelIn.Save();
        }
    }
}

