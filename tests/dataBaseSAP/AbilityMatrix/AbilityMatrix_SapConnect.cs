using Newtonsoft.Json;
using SAP.Middleware.Connector;
using SiggaPS.tests.dataBaseSAP.IndividualCapacity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechTalk.SpecFlow;

namespace SiggaPS.tests.dataBaseSAP.AbilityMatrix
{
    class AbilityMatrix_SapConnect
    {
        //Atributos das Planilhas
     public AbilityMatrix_LListExcel CreateExcel { get; set; }
     public AbilityMatrix_AmPersonExcel CreateExcel1 { get; set; }
     public AbilityMatrix_AssetGroup_AmPersonExcel CreateExcel2 { get; set; }



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
            parms.Add(RfcConfigParameters.Client, "100");
            
        }

        public void tableLList()
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "COD_ITEM DESCRICAO");
            IReader.SetValue("IV_FROM", "/SSCN/LLIST");
            
            if (FeatureContext.Current.FeatureInfo.Title.Equals("AbilityMatrix_AssetClass"))
            {
                IReader.SetValue("IV_WHERE", "NOME_LISTA = 'CLASSES' AND IDIOMA = 'EN'");
            
            } 
            else
            {
                IReader.SetValue("IV_WHERE", "NOME_LISTA = 'TIPO_OBJ' AND IDIOMA = 'EN'");
            } 

            
            IReader.SetValue("IV_ORDER", "COD_ITEM DESCRICAO");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathLList.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<AbilityMatrix_SapTable> jsonList = JsonConvert.DeserializeObject<List<AbilityMatrix_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel = new AbilityMatrix_LListExcel();
                CreateExcel.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosLList.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel.Save();
            }
        }

        public void tableAmPerson(string pernr)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "A~COD_ITEM A~DESCRICAO B~AM_KN");

            if (FeatureContext.Current.FeatureInfo.Title.Equals("AbilityMatrix_AssetClass"))
            {
                IReader.SetValue("IV_FROM", "/SSCN/LLIST AS A LEFT JOIN /SSCN/AM_PERSON AS B ON ( A~COD_ITEM = B~AM_KEY AND B~AM_TYPE = '1' AND B~PERNR = '" + pernr + "')");
                IReader.SetValue("IV_WHERE", "A~NOME_LISTA = 'CLASSES' AND A~IDIOMA = 'EN'");
                IReader.SetValue("IV_ORDER", "A~COD_ITEM");
            }
            else 
            {
                IReader.SetValue("IV_FROM", "/SSCN/LLIST AS A LEFT JOIN /SSCN/AM_PERSON AS B ON ( A~COD_ITEM = B~AM_KEY AND B~AM_TYPE = '2' AND B~PERNR = '" + pernr + "')");
                IReader.SetValue("IV_WHERE", "A~NOME_LISTA = 'TIPO_OBJ' AND A~IDIOMA = 'EN'");
                IReader.SetValue("IV_ORDER", "A~COD_ITEM");
            }
            
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathLPerson.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<AbilityMatrix_SapTable> jsonList = JsonConvert.DeserializeObject<List<AbilityMatrix_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel1 = new AbilityMatrix_AmPersonExcel();
                CreateExcel1.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosAmPerson.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel1.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel1.Save();
            }
        }

        public void tableAssetGroupAmPerson(string pernr)
        {
            connect();
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DYNAMIC_SELECT");
            IReader.SetValue("IV_SELECT", "A~PERNR A~AM_KEY_TYPE A~AM_KEY B~EQKTX C~PLTXT A~AM_KN A~EXP_EQUI");
            IReader.SetValue("IV_FROM", "/SSCN/AM_PERSON AS A LEFT JOIN /SSCN/LEQP_T AS B ON ( A~AM_KEY = B~EQUNR AND B~SPRAS = 'EN' ) LEFT JOIN /SSCN/LFUN_LOC_T AS C ON(A~AM_KEY = C~TPLNR AND C~SPRAS = 'EN')");
            IReader.SetValue("IV_WHERE", "A~AM_TYPE = '3' AND A~PERNR = '100029'");
            IReader.SetValue("IV_ORDER", "A~AM_KEY_TYPE A~AM_KEY");
            IReader.Invoke(rfcDest);
            string optionData = (string)IReader.GetValue("EV_RESULT_SET");
            Directory.CreateDirectory("Relatorio");
            filePathTxt = filePathTxt + @"\pathAGLPerson.txt";
            System.IO.File.WriteAllText(filePathTxt, optionData);
            List<AbilityMatrix_SapTable> jsonList = JsonConvert.DeserializeObject<List<AbilityMatrix_SapTable>>(optionData);
            if (jsonList.Count() > 0)
            {
                //file
                CreateExcel2 = new AbilityMatrix_AssetGroup_AmPersonExcel();
                CreateExcel2.CreateWorkbook(filePathXlsx = filePathXlsx + @"\DadosAgAmPerson.xlsx");
                //Adicionar Celulas
                jsonList
                    .Select((mandante, x) => new { mandante, position = x })
                    .ToList()
                    .ForEach(item => CreateExcel1.AddCell(item.position, item.mandante));
                //Salvando dados no excel
                CreateExcel2.Save();
            }

        }
    }
}
