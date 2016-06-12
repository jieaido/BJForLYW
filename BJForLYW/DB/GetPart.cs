using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BJForLYW.DB
{/// <summary>
/// 入库表
/// </summary>
    public class GetPart
    {/// <summary>
    /// 入库单ID
    /// </summary>
        public Int64 GetpartId { get; set; }
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
        /// 入库数量
        /// </summary>
        public Int64 GetNum { get; set; }
        /// <summary>
        /// 备注
        /// </summary>
        public string Remark { get; set; }
        /// <summary>
        /// 时间
        /// </summary>
        public string GetTime { get; set; }
    }
}
