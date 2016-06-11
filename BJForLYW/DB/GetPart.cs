using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BJForLYW.DB
{
    public class GetPart
    {
        public Int64 GetPartid { get; set; }
        public string PartNum { get; set; }
        public String PartName { get; set; }
        public string PartType { get; set; }
        public string Unit { get; set; }
        public decimal Price { get; set; }
        public Int64 GetNum { get; set; }
        public string Remark { get; set; }
        public string GetTime { get; set; }
    }
}
