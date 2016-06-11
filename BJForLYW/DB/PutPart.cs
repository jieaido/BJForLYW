using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BJForLYW.DB
{
    public class PutPart
    {
        public Int64 PutPartId { get; set; }
        public string PartNum { get; set; }
        public String PartName { get; set; }
        public string PartType { get; set; }
        public string Unit { get; set; }
        public decimal Price { get; set; }
        public Int64 PutNum { get; set; }
        public string Remark { get; set; }
        public string PutTime { get; set; }
        public string PutPeopleName { get; set; }
        public Int64 PartId { get; set; }
        public virtual Part Part { get; set; }

        
    }
}
