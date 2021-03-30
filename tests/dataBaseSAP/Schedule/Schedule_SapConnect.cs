using Newtonsoft.Json;
using OfficeOpenXml;
using SAP.Middleware.Connector;
using SiggaPS.tests.steps;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechTalk.SpecFlow;

namespace SiggaPS.tests.dataBaseSAP.Schedule
{
    class Schedule_SapConnect
    {
        public Schedule_PlantExcel CreateExcel { get; set; }
        public Schedule_LListExcel CreateExcel2 { get; set; }
        public Schedule_BLPlanExcel CreateExcel3 { get; set; }
        public Schedule_BLEorgExcel CreateExcel4 { get; set; }
        public Schedule_BLOperExcel CreateExcel5 { get; set; }
        public Schedule_LoperExcel CreateExcel6 { get; set; }
        public Schedule_ParamExcel CreateExcel7 { get; set; }
        public Schedule_LperWkcExcel CreateExcel8 { get; set; }

        public RfcConfigParameters parms = new RfcConfigParameters();

        public void connect()
        {
            parms.Add(RfcConfigParameters.Name, "QW1");
            parms.Add(RfcConfigParameters.AppServerHost, "10.10.10.177");
            parms.Add(RfcConfigParameters.SystemNumber, "30");
            parms.Add(RfcConfigParameters.User, "SIGGA127");
            parms.Add(RfcConfigParameters.Password, "123690");
            parms.Add(RfcConfigParameters.Client, "100");
        }
        public void tablePlantsInvokeTableLList()
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "USERID PLANT");
            IReader.SetValue("IV_FROM", "/SSCN/PLANT");
            IReader.SetValue("IV_WHERE", "USERID = 'SIGGA127'");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            string path = Path.GetFullPath("Relatorio/pathPlant.txt");
            System.IO.File.WriteAllText(path, optionData);
            List<Schedule_SapTable> jsonList = JsonConvert.DeserializeObject<List<Schedule_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel = new Schedule_PlantExcel();
                CreateExcel.CreateWorkbook("Relatorio/DadosPlant.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel.Save();
            }
            string list = CreateExcel.plantsToString();
            tableLList(list);
        }
        public void tableLList(string lista)
        {
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "*"); // pegar todas as tabelas
            IReader.SetValue("IV_FROM", "/SSCN/LLIST");
            IReader.SetValue("IV_WHERE", "NOME_LISTA = 'CENTRO_TRAB' AND CENTRO IN (" + lista + ") AND EXTRA_FLD3 = 'X' AND IDIOMA = 'EN'");
            IReader.SetValue("IV_ORDER", "COD_ITEM DESCRICAO");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            string path = Path.GetFullPath("Relatorio/pathLList.txt");
            System.IO.File.WriteAllText(path, optionData);
            List<Schedule_SapTable> jsonList = JsonConvert.DeserializeObject<List<Schedule_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel2 = new Schedule_LListExcel();
                CreateExcel2.CreateWorkbook("Relatorio/DadosLList.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel2.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel2.Save();
            }
        }
        public void tableBLPlan(string name)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "PLANNING_ID DESCRICAO DATA_INICIAL DATA_FINAL VALIDADE_DE VALIDADE_ATE");
            IReader.SetValue("IV_FROM", "/SSCN/BL_PLAN");
            IReader.SetValue("IV_WHERE", "DESCRICAO = '" + name + "'");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            string path = Path.GetFullPath("Relatorio/pathBLPlan.txt");
            System.IO.File.WriteAllText(path, optionData);
            List<Schedule_SapTable> jsonList = JsonConvert.DeserializeObject<List<Schedule_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel3 = new Schedule_BLPlanExcel();
                CreateExcel3.CreateWorkbook("Relatorio/DadosBLPlan.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel3.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel3.Save();

            }
        }
        public void tableBLEorg(string numero)
        {
            connect();
            List<Schedule_SapTable> tableBLEorg = new List<Schedule_SapTable>();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("RFC_READ_TABLE");
            IReader.SetValue("QUERY_TABLE", "/SSCN/BL_EORG");
            IReader.SetValue("ROWCOUNT", "0");
            //IReader.Invoke(rfcDest);
            IRfcTable optionsTable = IReader.GetTable("OPTIONS");
            optionsTable.Append();
            optionsTable.SetValue("TEXT", "PLANNING_ID = '" + numero + "'");
            IReader.Invoke(rfcDest);
            IRfcTable optionData = IReader.GetTable("DATA");
            Directory.CreateDirectory("Relatorio");
            tableBLEorg = optionData.AsQueryable().Select(x => new Schedule_SapTable
            {
                PLANNING_ID = x.GetString("WA").Substring(3, 32).TrimEnd() ?? "",
                CENTRO_TRAB = x.GetString("WA").Substring(39, 8).TrimEnd() ?? ""
            }
             ).ToList();
            string s = JsonConvert.SerializeObject(tableBLEorg);
            string path = Path.GetFullPath("Relatorio/pathBLEorg.txt");
            System.IO.File.WriteAllText(path, s);
            List<Schedule_SapTable> jsonList = JsonConvert.DeserializeObject<List<Schedule_SapTable>>(s);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel4 = new Schedule_BLEorgExcel();
                CreateExcel4.CreateWorkbook("Relatorio/DadosBLEorg.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel4.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel4.Save();
            }
        }
        public void tableBLOper(string numero, string orderId)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "ORDERID ACTIVITY SUB_ACTIVITY CONTROL_KEY WORK_CNTR PLANT DESCRIPTION SYSTCOND FUNCLOC EQUIPMENT EARL_SCH_START_D EARL_SCH_START_T");
            IReader.SetValue("IV_FROM", "/SSCN/BL_OPER");
            if (FeatureContext.Current.FeatureInfo.Title.Equals("Schedule - Move_Operations"))
            {
                IReader.SetValue("IV_WHERE", "ORDERID = '" + orderId + "'");
            }
            else
            {
                IReader.SetValue("IV_WHERE", "PLANNING_ID = '" + numero + "'");
            }
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            string path = Path.GetFullPath("Relatorio/pathBLOper.txt");
            System.IO.File.WriteAllText(path, optionData);
            List<Schedule_SapTable> jsonList = JsonConvert.DeserializeObject<List<Schedule_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel5 = new Schedule_BLOperExcel();
                CreateExcel5.CreateWorkbook("Relatorio/DadosBLOper.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel5.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel5.Save();
            }
        }
        public void tableLoper(string nome)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "ORDERID ACTIVITY SUB_ACTIVITY CONTROL_KEY WORK_CNTR PLANT DESCRIPTION SYSTCOND FUNCLOC EQUIPMENT EARL_SCH_START_D EARL_SCH_START_T");
            IReader.SetValue("IV_FROM", "/SSCN/LOPER");
            IReader.SetValue("IV_WHERE", "WORK_CNTR = '" + nome + "'");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            string path = Path.GetFullPath("Relatorio/pathLoper.txt");
            System.IO.File.WriteAllText(path, optionData);
            List<Schedule_SapTable> jsonList = JsonConvert.DeserializeObject<List<Schedule_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel6 = new Schedule_LoperExcel();
                CreateExcel6.CreateWorkbook("Relatorio/DadosLoper.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel6.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel6.Save();
            }
        }
        public void tableParam()
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "FIRST_DAY_WEEK SCHED_VALID_TIME");
            IReader.SetValue("IV_FROM", "/SSCN/PARAM");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            string path = Path.GetFullPath("Relatorio/pathParam.txt");
            System.IO.File.WriteAllText(path, optionData);
            List<Schedule_SapTable> jsonList = JsonConvert.DeserializeObject<List<Schedule_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel7 = new Schedule_ParamExcel();
                CreateExcel7.CreateWorkbook("Relatorio/DadosParam.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel7.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel7.Save();
            }
        }
        public void tableLperWkc()
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "PERNR ENDDA");
            IReader.SetValue("IV_FROM", "/SSCN/LPER_WKC");
            IReader.SetValue("IV_WHERE", "ARBPL = 'PSQA-AUT' AND WERKS = '1000'");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            string path = Path.GetFullPath("Relatorio/pathLperWkc.txt");
            System.IO.File.WriteAllText(path, optionData);
            List<Schedule_SapTable> jsonList = JsonConvert.DeserializeObject<List<Schedule_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel8 = new Schedule_LperWkcExcel();
                CreateExcel8.CreateWorkbook("Relatorio/DadosLperWkc.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel8.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel8.Save();
            }
        }
    }
}

