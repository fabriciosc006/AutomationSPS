using Newtonsoft.Json;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechTalk.SpecFlow;

namespace SiggaPS.tests.dataBaseSAP.MachineDowntimeCalendar
{
    class MachineDowntimeCalendar_SapConnect
    {

        public MachineDowntimeCalendar_PlantExcel CreateExcel { get; set; }
        public MachineDownTuneCalendar_LListExcel CreateExcel2 { get; set; }
        public MachineDowntimeCalendar_FuncLocExcel CreateExcel3 { get; set; }
        public MachineDowntimeCalendar_LeqpExcel CreateExcel4 { get; set; }
        public MachineDowntimeCalendar_Mach_CalExcel CreateExcel5 { get; set; }
        public string filePathTxt = System.IO.Directory.GetParent(System.IO.Directory.GetParent(System.IO.Directory.GetParent(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)).FullName).FullName).FullName + @"\Relatorio";
        public string filePathXlsx = System.IO.Directory.GetParent(System.IO.Directory.GetParent(System.IO.Directory.GetParent(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)).FullName).FullName).FullName + @"\Relatorio";

        public RfcConfigParameters parms = new RfcConfigParameters();
        public void connect()
        {
            parms.Add(RfcConfigParameters.Name, "QW1");
            parms.Add(RfcConfigParameters.AppServerHost, "10.10.10.177");
            parms.Add(RfcConfigParameters.SystemNumber, "30");
            parms.Add(RfcConfigParameters.User, "SIGGA127");
            parms.Add(RfcConfigParameters.Password, "123690");
            if (FeatureContext.Current.FeatureInfo.Title.Equals("MDC - Manage MDC"))
            {
                parms.Add(RfcConfigParameters.Client, "500");
            }
            else
            {
                parms.Add(RfcConfigParameters.Client, "100");
            }
        }
        public void tablePlants()
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
            filePathTxt = filePathTxt + @"\pathPlant.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<MachineDowntimeCalendar_SapTable> jsonList = JsonConvert.DeserializeObject<List<MachineDowntimeCalendar_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel = new MachineDowntimeCalendar_PlantExcel();
                CreateExcel.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosPlant.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel.Save();
            }
        }
        public void tableLList(string lista)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "COD_ITEM DESCRICAO");
            IReader.SetValue("IV_FROM", "/SSCN/LLIST");
            IReader.SetValue("IV_WHERE", "NOME_LISTA = 'CENTRO_TRAB' AND CENTRO = '" + lista + "' AND EXTRA_FLD3 = 'X' AND IDIOMA = 'EN'");
            IReader.SetValue("IV_ORDER", "COD_ITEM DESCRICAO");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathLList.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<MachineDowntimeCalendar_SapTable> jsonList = JsonConvert.DeserializeObject<List<MachineDowntimeCalendar_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel2 = new MachineDownTuneCalendar_LListExcel();
                CreateExcel2.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosLList.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel2.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel2.Save();
            }
        }
        public void tableLListScenario3(string lista, string code)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "COD_ITEM DESCRICAO");
            IReader.SetValue("IV_FROM", "/SSCN/LLIST");
            IReader.SetValue("IV_WHERE", "NOME_LISTA = 'CENTRO_TRAB' AND CENTRO = '" + lista + "' AND COD_ITEM = '" + code + "' AND EXTRA_FLD3 = 'X' AND IDIOMA = 'EN'");
            IReader.SetValue("IV_ORDER", "COD_ITEM DESCRICAO");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathLList.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<MachineDowntimeCalendar_SapTable> jsonList = JsonConvert.DeserializeObject<List<MachineDowntimeCalendar_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel2 = new MachineDownTuneCalendar_LListExcel();
                CreateExcel2.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosLList.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel2.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel2.Save();
            }
        }
        public void tableFuncLoc(string plant, string data, string workCenter, string functionalLoc)
        {
            connect();
            int datalimit = Convert.ToInt32(data);
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "A~FUNC_LOC B~PLTXT");
            IReader.SetValue("IV_FROM", "/SSCN/LFUNC_LOC AS A LEFT JOIN /SSCN/LFUN_LOC_T AS B ON ( A~FUNC_LOC = B~TPLNR AND B~SPRAS = 'EN' )");
            string scenario = ScenarioContext.Current.ScenarioInfo.Title;
            switch (scenario)
            {
                case "Filtering only by plant":
                    IReader.SetValue("IV_WHERE", "A~MAINTPLANT = '" + plant + "' AND A~FLAG_DELETED = '' AND A~FLAG_INACTIVE = ''");
                    break;
                case "Filtering by plant and work center":
                    IReader.SetValue("IV_WHERE", "A~MAINTPLANT = '" + plant + "' AND A~WORK_CTR = '" + workCenter + "' AND A~FLAG_DELETED = '' AND A~FLAG_INACTIVE = ''");
                    break;
                case "Filtering by plant and Functional location":
                    IReader.SetValue("IV_WHERE", "A~MAINTPLANT = '" + plant + "' AND A~FUNC_LOC LIKE '%" + functionalLoc + "%' AND A~FLAG_DELETED = '' AND A~FLAG_INACTIVE = ''");
                    break;
            }
            IReader.SetValue("IV_ORDER", "A~FUNC_LOC");
            IReader.SetValue("IV_ROWS", datalimit);
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathFuncLoc.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<MachineDowntimeCalendar_SapTable> jsonList = JsonConvert.DeserializeObject<List<MachineDowntimeCalendar_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel3 = new MachineDowntimeCalendar_FuncLocExcel();
                CreateExcel3.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosFuncLoc.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel3.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel3.Save();
            }
        }

        public void tableLEQP(string plant, string data, string workCenter, string equipment)
        {
            connect();
            int datalimit = Convert.ToInt32(data);
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "A~EQUNR B~EQKTX");
            IReader.SetValue("IV_FROM", "/SSCN/LEQP AS A LEFT JOIN /SSCN/LEQP_T AS B ON ( A~EQUNR = B~EQUNR AND B~SPRAS = 'EN' )");
            string scenario = ScenarioContext.Current.ScenarioInfo.Title;
            switch (scenario)
            {
                case "Filtering only by plant":
                    IReader.SetValue("IV_WHERE", "A~MAINTPLANT = '" + plant + "'");
                    break;
                case "Filtering by plant and work center":
                    IReader.SetValue("IV_WHERE", "A~MAINTPLANT = '" + plant + "' AND A~WORK_CTR = '" + workCenter + "'");
                    break;
                case "Filtering by plant and Equipment":
                    IReader.SetValue("IV_WHERE", "A~MAINTPLANT = '" + plant + "' AND A~EQUNR LIKE '%" + equipment + "%'");
                    break;
            }
            IReader.SetValue("IV_ORDER", "A~EQUNR");
            IReader.SetValue("IV_ROWS", datalimit);
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathLeqp.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<MachineDowntimeCalendar_SapTable> jsonList = JsonConvert.DeserializeObject<List<MachineDowntimeCalendar_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel4 = new MachineDowntimeCalendar_LeqpExcel();
                CreateExcel4.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosLeqp.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel4.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel4.Save();
            }
        }
        public void tableMachCal(string plant, string workCenter, string functionalLoc, string equipment, string startDate, string finishDate)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "A~FUNC_LOC B~PLTXT A~EQUNR C~EQKTX A~START_DATE A~FINISH_DATE A~START_TIME A~FINISH_TIME A~STOP_TYPE A~TIME_CONF A~NOTE");
            IReader.SetValue("IV_FROM", "/SSCN/MACH_CAL AS A LEFT JOIN /SSCN/LFUN_LOC_T AS B ON ( A~FUNC_LOC = B~TPLNR AND B~SPRAS = 'EN') LEFT JOIN /SSCN/LEQP_T AS C ON ( A~EQUNR = C~EQUNR AND C~SPRAS = 'EN' )");
            string scenario = ScenarioContext.Current.ScenarioInfo.Title;
            switch (scenario)
            {
                case "Filtering only by plant":
                    IReader.SetValue("IV_WHERE", "PLANT = '" + plant + "'");
                    break;
                case "Filtering by plant and work center":
                    IReader.SetValue("IV_WHERE", "PLANT = '" + plant + "' AND WORK_CENTER = '" + workCenter + "'");
                    break;
                case "Filtering by plant and Functional location":
                    IReader.SetValue("IV_WHERE", "PLANT = '" + plant + "' AND A~FUNC_LOC LIKE '%" + functionalLoc + "%'");
                    break;
                case "Filtering by plant and Equipment":
                    IReader.SetValue("IV_WHERE", "PLANT = '" + plant + "' AND A~EQUNR LIKE '%" + equipment + "%'");
                    break;
                case "Filtering by plant and date":
                    IReader.SetValue("IV_WHERE", "PLANT = '" + plant + "' AND A~START_DATE GE '" + startDate + "' AND A~START_DATE LE '" + finishDate + "'");
                    break;
            }
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathMachCal.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<MachineDowntimeCalendar_SapTable> jsonList = JsonConvert.DeserializeObject<List<MachineDowntimeCalendar_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel5 = new MachineDowntimeCalendar_Mach_CalExcel();
                CreateExcel5.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosMachCal.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel5.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel5.Save();
            }
        }

        public void tableMachCalFeatureManage(string plant, string functionalLoc)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "A~FUNC_LOC B~PLTXT A~EQUNR C~EQKTX A~START_DATE A~FINISH_DATE A~START_TIME A~FINISH_TIME A~STOP_TYPE A~TIME_CONF A~NOTE");
            IReader.SetValue("IV_FROM", "/SSCN/MACH_CAL AS A LEFT JOIN /SSCN/LFUN_LOC_T AS B ON ( A~FUNC_LOC = B~TPLNR AND B~SPRAS = 'EN') LEFT JOIN /SSCN/LEQP_T AS C ON ( A~EQUNR = C~EQUNR AND C~SPRAS = 'EN' )");
            string scenario = ScenarioContext.Current.ScenarioInfo.Title;
            switch (scenario)
            {
                case "Deleting two (or more) downtimes":
                    IReader.SetValue("IV_WHERE", "PLANT = '" + plant + "'");
                    break;
                case "Editing one downtime":
                    IReader.SetValue("IV_WHERE", "PLANT = '" + plant + "'");
                    break;
                case "Registering a downtime for two (or more) equipments":
                    IReader.SetValue("IV_WHERE", "PLANT = '" + plant + "' AND A~FUNC_LOC LIKE '%" + functionalLoc + "%'");
                    break;
                case "Deleting one downtime":
                    IReader.SetValue("IV_WHERE", "PLANT = '" + plant + "'");
                    break;
                case "Registering a downtime for one functional location":
                    IReader.SetValue("IV_WHERE", "PLANT = '" + plant + "' AND A~FUNC_LOC LIKE '%" + functionalLoc + "%'");
                    break;
                case "Registering a downtime for two (or more) functional locations":
                    IReader.SetValue("IV_WHERE", "PLANT = '" + plant + "' AND A~FUNC_LOC LIKE '%" + functionalLoc + "%'");
                    break;
                case "Registering a downtime for one equipment":
                    IReader.SetValue("IV_WHERE", "PLANT = '" + plant + "'");
                    break;
            }
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathMachCal.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<MachineDowntimeCalendar_SapTable> jsonList = JsonConvert.DeserializeObject<List<MachineDowntimeCalendar_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel5 = new MachineDowntimeCalendar_Mach_CalExcel();
                CreateExcel5.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosMachCal.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel5.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel5.Save();
            }
        }
        public void cleanTableMachCal(string datetine)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_DELETE");
            IReader.SetValue("IV_FROM", "/SSCN/MACH_CAL");
            IReader.SetValue("IV_WHERE", "FINISH_DATE >= '" + datetine + "'");
            IReader.Invoke(rfcDest);
        }
    }
}
