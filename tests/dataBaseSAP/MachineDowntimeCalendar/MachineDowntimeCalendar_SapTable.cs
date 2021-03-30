using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiggaPS.tests.dataBaseSAP.MachineDowntimeCalendar
{
    class MachineDowntimeCalendar_SapTable
    {
        public string USERID { get; set; }
        public string PLANT { get; set; }
        public string COD_ITEM { get; set; }
        public string DESCRICAO { get; set; }
        public string FUNC_LOC { get; set; }
        public string EQUNR { get; set; }
        public string EQKTX { get; set; }
        public string START_DATE { get; set; }
        public string FINISH_DATE { get; set; }
        public string PLTXT { get; set; }
        public string START_TIME { get; set; }
        public string FINISH_TIME { get; set; }
        public string STOP_TYPE { get; set; }
        public string TIME_CONF { get; set; }
        public string NOTE { get; set; }
    }
}
