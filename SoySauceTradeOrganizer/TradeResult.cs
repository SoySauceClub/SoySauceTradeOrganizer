using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoySauceTradeOrganizer
{
    public class TradeResult
    {
        public DateTime EntryDate { get; set; }
        public DateTime ExitDate { get; set; }
        public string EnterDirection { get; set; }
        public decimal EnterPrice { get; set; }
        public decimal ExitPrice { get; set; }
        public string Ticker { get; set; }
        public decimal InitialStop { get; set; }
        public int isMissed { get; set; }
        public decimal TargetPrice { get; set; }
    }
}
