using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiggaPS.tests.dataBaseSAP.MachineDowntimeCalendar
{
   public class MachineDowntimeCalendar_CellsDTO
    {
        public string Number { get; set; }
        public string Text { get; set; }
        public string Instalation { get; set; }
        public string DescriptionPltxt { get; set; }
        public string Equipments { get; set; }
        public string DescriptionEqktx { get; set; }
        public string InitialData { get; set; }
        public string finishData { get; set; }
        public string InitialTime { get; set; }
        public string finishTime { get; set; }
        public string paradeType { get; set; }
        public string paradeAlocation { get; set; }
        public string observation { get; set; }
        public int index { get; set; }
    }
}
