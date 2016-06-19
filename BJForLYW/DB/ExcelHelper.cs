using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Entity.Migrations;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BJForLYW.Properties;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace BJForLYW.DB
{
    public static class ExcelHelper
    {
        /// <summary>
        /// 从excel文件导入到入库表
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static List<GetPart> GetPartFromExcel(string filePath)
        {
            List<GetPart> parts = new List<GetPart>();
            HSSFWorkbook hssfWorkbook;
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                hssfWorkbook = new HSSFWorkbook(fileStream);
            }
            var sheet = hssfWorkbook.GetSheetAt(0);
            var rows = sheet.GetRowEnumerator();
            rows.MoveNext();
            using (PartContext pc = new PartContext())
            {
                while (rows.MoveNext())
                {
                    HSSFRow row = (HSSFRow)rows.Current;
                    GetPart part = new GetPart();
                    if (row.FirstCellNum<0)
                    {
                        continue;
                    }
                    if (row.FirstCellNum == 0)
                    {
                        
                        part.PartNum = row.Cells[0].ToString().Trim();
                        part.PartName = row.Cells[1].ToString().Trim();
                        part.PartType = row.Cells[2].ToString().Trim();
                        part.Unit = row.Cells[3].ToString().Trim();
                        part.Price = (decimal?) row.Cells[4].NumericCellValue;
                        part.GetNum = long.Parse(GetStringCellValue(row.Cells[5]));
                    }
                    else
                    {
                        part.PartNum = "";
                        part.PartName = row.Cells[0].ToString().Trim();
                        part.PartType = row.Cells[1].ToString().Trim();
                        part.Unit = row.Cells[2].ToString().Trim();

                        part.Price = (decimal?) row.Cells[3].NumericCellValue;
                        part.GetNum = long.Parse(GetStringCellValue(row.Cells[4]));
                    }
                    part.GetTime = DateTime.Now.ToShortDateString();
                    parts.Add(part);
                    // pc.GetParts.Add(part);
                }
                //pc.SaveChanges();
            }
            return parts;
        }

        public static void ConfimGetPart(IEnumerable<GetPart> getParts,PartContext pc)
        {
           
                foreach (var getPart in getParts)
                {
                    Part findPart;  
                    if (getPart.PartNum != "")
                    {
                        findPart = pc.Parts.FirstOrDefault(gp => gp.PartNum == getPart.PartNum);

                    }
                    else
                    {
                        findPart =
                            pc.Parts.FirstOrDefault(
                                gp => gp.PartName == getPart.PartName && gp.PartType == getPart.PartType);

                    }
                    if (findPart != null)
                    {
                        findPart.Num += getPart.GetNum; 
                    }
                    else
                    {
                        findPart = new Part()
                        {
                            PartName = getPart.PartName,
                            PartType = getPart.PartType,
                            PartNum = getPart.PartNum,
                            Price = getPart.Price,
                            Num = getPart.GetNum,
                            Unit = getPart.Unit,
                            Remark = getPart.GetTime
                        };
                        
                    }
                    pc.GetParts.AddOrUpdate(getPart);
                    pc.Parts.AddOrUpdate(findPart);

                }
                pc.SaveChanges();
            


        }

        public static PutPart GenerationPutPartFromPart(Part part, long putnum, string puttime, string putPeopleName)
        {
            PutPart  putPart=new PutPart();
            putPart.PartNum = part.PartNum;
            putPart.PartName = part.PartName;
            putPart.PartType = part.PartType;
            putPart.Unit = part.Unit;
            putPart.Price = part.Price;
            putPart.PartId = part.Partid;
            putPart.Part = part;
            putPart.PutNum = putnum;
            putPart.PutTime = puttime;
            putPart.PutPeopleName = putPeopleName;
            return putPart;
        }
       static HSSFWorkbook InitializeWorkbook()
        {
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "安钢动力厂计控车间";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "备件表";
            hssfworkbook.SummaryInformation = si;
            return hssfworkbook;
        }
        static void WriteToFile(HSSFWorkbook hssfWorkbook, string filename)
       {
            string pathCurr = System.Environment.CurrentDirectory;
            string pathstr= Path.Combine(pathCurr, "导出Excel文件", filename);
           if (!Directory.Exists(pathstr))
           {
               Directory.CreateDirectory(pathstr);
           }
           string filePath = Path.Combine(pathstr, DateTime.Now.ToString("yyyy年MM月dd天HH时mm分ss秒")+".xls");
           // string ss=  System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //Write the stream data of workbook to the root directory
           //string sss = ss + "\\test.xls";
            FileStream file = new FileStream(filePath, FileMode.Create);
            hssfWorkbook.Write(file);
            file.Close();
         
           if (MessageBox.Show(Resources.ExcelHelper_WriteToFile_导出成功是否打开,Resources.ExcelHelper_WriteToFile_提示,MessageBoxButtons.YesNo)==DialogResult.Yes)
           {
               System.Diagnostics.Process.Start(filePath);
           }
        }

        public static void DataGridViewToExcel(DataGridView dataGridView, string filename)
        {
            HSSFWorkbook hssfWorkbook = InitializeWorkbook();
            var sheet1 = hssfWorkbook.CreateSheet("Sheet1");
            var row1 = sheet1.CreateRow(0);
            for (int i = 0; i < dataGridView.ColumnCount-1; i++)
            {
                row1.CreateCell(i).SetCellValue(dataGridView.Columns[i+1].HeaderText);
            }
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                var row2 = sheet1.CreateRow(i + 1);
                for (int j = 0; j <dataGridView.Rows[i].Cells.Count-1; j++)
                {
                    var value = dataGridView.Rows[i].Cells[j+1].Value;
                    if (value != null)
                        row2.CreateCell(j).SetCellValue(value.ToString());
                }
            }

            ExcelHelper.WriteToFile(hssfWorkbook, filename);
        }

        private static string GetStringCellValue(ICell cell)
        {
            if (cell==null)
            {
                return "";
            }
            else
            {
                return cell.ToString();
            }

        }

       
    }
}
