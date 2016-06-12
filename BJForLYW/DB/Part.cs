using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BJForLYW.DB
{/// <summary>
/// 设备表
/// </summary>
    public class Part
    {
        /// <summary>
        /// 设备ID
        /// </summary>
        public Int64 Partid { get; set; }
        /// <summary>
        /// 物料编码
        /// </summary>
        public string PartNum { get; set; }
        /// <summary>
        /// 备件名称
        /// </summary>
        public String PartName { get; set; }
        /// <summary>
        /// 备件型号
        /// </summary>
        public string PartType { get; set; }
        /// <summary>
        /// 单位
        /// </summary>
        public string Unit { get; set; }
        /// <summary>
        /// 单价
        /// </summary>
        public decimal? Price { get; set; }
        /// <summary>
        /// 数量
        /// </summary>
    public Int64 Num { get; set; }
        /// <summary>
        /// 备注
        /// </summary>
        public string Remark { get; set; }
 
     
       
    }
}
