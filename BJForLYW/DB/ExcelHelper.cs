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
        public static List<GetPart> GetgetPartTableFromExcel(string filePath)
        {
            var parts = new List<GetPart>();
           // List<GetPart> parts = new List<GetPart>();
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
                        
                        part.PartNum = row.GetCell(0).ToString().Trim();
                        part.PartName = row.GetCell(1).ToString().Trim();
                        part.PartType = row.GetCell(2).ToString().Trim();
                        part.Unit = row.GetCell(3).ToString().Trim();
                        //bool b = row.GetCell(4) == null;
                        part.Price = row.GetCell(4)==null?0:(decimal?) row.GetCell(4).NumericCellValue;
                        part.GetNum = (long) row.GetCell(5).NumericCellValue;
                    }
                    else
                    {
                        part.PartNum = "";
                        part.PartName = row.Cells[0].ToString().Trim();
                        part.PartType = row.Cells[1].ToString().Trim();
                        part.Unit = row.Cells[2].ToString().Trim();

                        part.Price = (decimal?) row.Cells[3].NumericCellValue;
                        part.GetNum = (long)row.Cells[4].NumericCellValue;
                    }
                    part.GetTime = DateTime.Now.ToShortDateString();
                    parts.Add(part);
                    // pc.GetParts.Add(part);
                }
                //pc.SaveChanges();
            }
            MessageBox.Show($"成功导入{parts.Count}条数据");
            return parts;
        }
        /// <summary>
        /// 从excel文件导入到设备表
        /// </summary>
        /// <param name="filePath">导入的文件路径</param>
        /// <returns></returns>
        public static List<Part> GetPartTableFromExcel(string filePath)
        {
            List<GetPart> getParts = GetgetPartTableFromExcel(filePath);
            List<Part> parts=new List<Part>();
            foreach (var getPart in getParts)
            {
                parts.Add(ConvertGetPartToPart(getPart));
            }
            return parts;


        }
        /// <summary>
        /// 从getpart转换到Part
        /// </summary>
        /// <param name="getPart">要转换的实体</param>
        /// <returns></returns>
        static Part ConvertGetPartToPart(GetPart getPart)
        {
            Part part=new Part()
            {
                PartNum = getPart.PartNum,
                PartName = getPart.PartName,
                PartType = getPart.PartType,
                Price = getPart.Price,
                Unit = getPart.Unit,
                Num = getPart.GetNum
            };
            return part;
        }
        /// <summary>
        /// 确认导入的入库表数量并更新设备表
        /// </summary>
        /// <param name="getParts"></param>
        /// <param name="pc"></param>
        public static void ConfimGetPart(IEnumerable<GetPart> getParts,PartContext pc)
        {
           
                foreach (var getPart in getParts)
                {
                    Part findPart;  
                      //首先要确认的是物料编码
                    if (getPart.PartNum != "")
                    {
                        findPart = pc.Parts.FirstOrDefault(gp => gp.PartNum == getPart.PartNum);

                    }
                   
                    else//如果不行就查询名称或者型号
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
        /// <summary>
        /// 根据选择的备件生成出库表
        /// </summary>
        /// <param name="part">要出的备件</param>
        /// <param name="putnum">要出的备件数量</param>
        /// <param name="puttime">出库的时间</param>
        /// <param name="putPeopleName">出库人</param>
        /// <param name="remarks">备注</param>
        /// <returns></returns>
        public static PutPart GenerationPutPartFromPart(Part part, long putnum, string puttime, string putPeopleName,string remarks)
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
            putPart.Remark = remarks;
            return putPart;
        }

        /// <summary>
        /// 初始化excel表
        /// </summary>
        /// <returns></returns>
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
        /// <summary>
        /// 写入Excel文件
        /// </summary>
        /// <param name="hssfWorkbook"></param>
        /// <param name="filename"></param>
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
        /// <summary>
        /// 从datagridview导出到excel
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="filename"></param>
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
                    {
                        decimal tempnum;
                        if (decimal.TryParse(value.ToString(), out tempnum))
                        {
                            row2.CreateCell(j).SetCellValue((double) tempnum);
                        }
                        else
                        {
                            row2.CreateCell(j).SetCellValue(value.ToString());
                        }
                       
                    }
                    
                }
            }
            for (int i = 0; i < 10; i++)
            {
                sheet1.AutoSizeColumn(i);
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
