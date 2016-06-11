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
        public static List<Part> GetPartFromExcel(string filePath)
        {
            List<Part> parts = new List<Part>();
            HSSFWorkbook hssfWorkbook;
            using (FileStream fileStream=new FileStream(filePath,FileMode.Open,FileAccess.Read))
            {
                hssfWorkbook=new HSSFWorkbook(fileStream);
            }
            var sheet= hssfWorkbook.GetSheetAt(0);
            var rows = sheet.GetRowEnumerator();
            rows.MoveNext();
            while (rows.MoveNext())
            {
                HSSFRow row= (HSSFRow) rows.Current;
                Part part = new Part();
                if (row.FirstCellNum == 0)
                {
                    part.PartNum = row.Cells[0].ToString();
                    part.PartName = row.Cells[1].ToString();
                    part.PartType = row.Cells[2].ToString();
                    part.Unit = row.Cells[3].ToString();
                    part.Price = (decimal?) row.Cells[4].NumericCellValue;
                    part.Num = (long) row.Cells[5].NumericCellValue;
                }
                else
                {
                    part.PartNum = "";
                    part.PartName = row.Cells[0].ToString();
                    part.PartType = row.Cells[1].ToString();
                    part.Unit = row.Cells[2].ToString();
                    part.Price = (decimal?)row.Cells[3].NumericCellValue;
                    part.Num = (long)row.Cells[4].NumericCellValue;
                }
             
              
                parts.Add(part);




            }
            
            return parts;
        }
              
    }
}
