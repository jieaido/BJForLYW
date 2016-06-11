using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;

namespace BJForLYW.DB
{
    public static  class ExcelHelper
    {
        public static void GetPartFromExcel(string filePath)
        {
            HSSFWorkbook hssfWorkbook;
            using (FileStream fileStream=new FileStream(filePath,FileMode.Open,FileAccess.Read))
            {
                hssfWorkbook=new HSSFWorkbook(fileStream);
            }
            var sheet= hssfWorkbook.GetSheetAt(0);
            var rows = sheet.GetRowEnumerator();
            
            return ;
        }
              
    }
}
