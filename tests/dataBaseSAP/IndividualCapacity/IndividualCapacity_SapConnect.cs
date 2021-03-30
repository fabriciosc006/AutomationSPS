using Newtonsoft.Json;
using NUnit.Framework;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechTalk.SpecFlow;

namespace SiggaPS.tests.dataBaseSAP.IndividualCapacity
{
    class IndividualCapacity_SapConnect
    {
        public IndividualCapacity_PlantExcel CreateExcel { get; set; }
        public IndividualCapacity_LListExcel CreateExcel1 { get; set; }
        public IndividualCapacity_LPersonExcel CreateExcel2 { get; set; }
        public IndividualCapacity_LperWkcExcel CreateExcel3 { get; set; }
        public IndividualCapacity_IndivCapExcel CreateExcel4 { get; set; }
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
            if (FeatureContext.Current.FeatureInfo.Title.Equals("IndividualCapacity - Manage"))
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
            List<IndividualCapacity_SapTable> jsonList = JsonConvert.DeserializeObject<List<IndividualCapacity_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel = new IndividualCapacity_PlantExcel();
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

        public void tableLList(string plant)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "COD_ITEM DESCRICAO");
            IReader.SetValue("IV_FROM", "/SSCN/LLIST");
            IReader.SetValue("IV_WHERE", "NOME_LISTA = 'CENTRO_TRAB' AND CENTRO = '" + plant + "' AND EXTRA_FLD3 = 'X' AND IDIOMA = 'EN'");
            IReader.SetValue("IV_ORDER", "COD_ITEM DESCRICAO");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathLList.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<IndividualCapacity_SapTable> jsonList = JsonConvert.DeserializeObject<List<IndividualCapacity_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel1 = new IndividualCapacity_LListExcel();
                CreateExcel1.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosLList.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel1.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel1.Save();
            }
        }

        public void tableLPerson(string plant, string workcenter)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "A~PERNR A~ENAME");
            IReader.SetValue("IV_FROM", "/SSCN/LPERSON AS A INNER JOIN /SSCN/LPER_WKC AS B ON ( A~PERNR = B~PERNR )");
            IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "'");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathLPerson.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<IndividualCapacity_SapTable> jsonList = JsonConvert.DeserializeObject<List<IndividualCapacity_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel2 = new IndividualCapacity_LPersonExcel();
                CreateExcel2.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosLPerson.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel2.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel2.Save();
            }
        }

        public void tableLPer_Wkc(string plant, string workcenter)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "A~PERNR A~ENAME");
            IReader.SetValue("IV_FROM", "/SSCN/LPERSON AS A INNER JOIN /SSCN/LPER_WKC AS B ON ( A~PERNR = B~PERNR )");
            IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "'");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathLPerWkc.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<IndividualCapacity_SapTable> jsonList = JsonConvert.DeserializeObject<List<IndividualCapacity_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel3 = new IndividualCapacity_LperWkcExcel();
                CreateExcel3.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosLPerWkc.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel3.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel3.Save();
            }
        }

        public void tableIndivCapa(string plant, string workcenter, string onePerson, string morPerson, string startDate, string finishDate)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "A~PERNR A~ENAME C~START_DATE C~FINISH_DATE C~START_TIME C~FINISH_TIME C~NOTE");
            IReader.SetValue("IV_FROM", "/SSCN/LPERSON AS A INNER JOIN /SSCN/LPER_WKC AS B ON ( A~PERNR = B~PERNR ) INNER JOIN /SSCN/INDIV_CAPA AS C ON ( C~PERNR = B~PERNR )");
            string scenario = ScenarioContext.Current.ScenarioInfo.Title;
            switch (scenario)
            {
                case "Search by plant":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "'");
                    break;
                case "Search by work center":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "'");
                    break;
                case "Search by an specific person":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR = '" + onePerson + "'");
                    break;
                case "Search by a group of people":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR IN (" + morPerson + ")");
                    break;
                case "Search by plant in an specific period":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND C~START_DATE GE '" + startDate + "' AND C~FINISH_DATE LE '" + finishDate + "'");
                    break;
                case "Search by work center in an specific period":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND C~START_DATE GE '" + startDate + "' AND C~FINISH_DATE LE '" + finishDate + "'");
                    break;
                case "Search by an specific person in a specific period":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR = '" + onePerson + "' AND C~START_DATE GE '" + startDate + "' AND C~FINISH_DATE LE '" + finishDate + "'");
                    break;
                case "Search by a group of people in a specific period":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR IN (" + morPerson + ") AND C~START_DATE GE '" + startDate + "' AND C~FINISH_DATE LE '" + finishDate + "'");
                    break;
                case "Registering a capacity for one person":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR = '" + onePerson + "'");
                    break;
                case "Registering a capacity for two or more people":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "'");
                    break;
                case "Registering a capacity for one person that already has a capacity for that period, but without overlaping of time":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR = '" + onePerson + "'");
                    break;
                case "Deleting a capacity":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR = '" + onePerson + "'");
                    break;
                case "Deleting two or more capacities":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR IN (" + morPerson + ")");
                    break;
                case "Editing a capacity":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR = '" + onePerson + "'");
                    break;
                case "Editing a capacity with overlap":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR = '" + onePerson + "'");
                    break;
                case "Copy a capacity":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR = '" + onePerson + "'");
                    break;
                case "Copy a capacity to more than one person":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR IN (" + morPerson + ")");
                    break;
                case "Copy a capacity with overlap":
                    IReader.SetValue("IV_WHERE", "B~WERKS = '" + plant + "' AND B~ARBPL = '" + workcenter + "' AND A~PERNR = '" + onePerson + "'");
                    break;
            }

            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathIndivCapa.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<IndividualCapacity_SapTable> jsonList = JsonConvert.DeserializeObject<List<IndividualCapacity_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel4 = new IndividualCapacity_IndivCapExcel();
                CreateExcel4.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosIndivCapa.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel4.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel4.Save();
            }
        }

        public void cleanTableIndividualCapacityOnePerson(string person)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_DELETE");
            IReader.SetValue("IV_FROM", "/SSCN/INDIV_CAPA");
            IReader.SetValue("IV_WHERE", "PERNR = '" + person + "'");
            IReader.Invoke(rfcDest);
        }

        public void cleanTableIndividualCapacityMorePerson(string persons)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_DELETE");
            IReader.SetValue("IV_FROM", "/SSCN/INDIV_CAPA");
            IReader.SetValue("IV_WHERE", "PERNR IN (" + persons + ")");
            IReader.Invoke(rfcDest);
        }
    }
}

