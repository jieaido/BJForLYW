﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Entity.Migrations;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;

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
                    if (row.FirstCellNum == 0)
                    {
                        part.PartNum = row.Cells[0].ToString().Trim();
                        part.PartName = row.Cells[1].ToString().Trim();
                        part.PartType = row.Cells[2].ToString().Trim();
                        part.Unit = row.Cells[3].ToString().Trim();
                        part.Price = (decimal?)row.Cells[4].NumericCellValue;
                        part.GetNum = (long)row.Cells[5].NumericCellValue;
                    }
                    else
                    {
                        part.PartNum = "";
                        part.PartName = row.Cells[0].ToString().Trim();
                        part.PartType = row.Cells[1].ToString().Trim();
                        part.Unit = row.Cells[2].ToString().Trim();
                        part.Price = (decimal?)row.Cells[3].NumericCellValue;
                        part.GetNum = (long)row.Cells[4].NumericCellValue;
                    }
                    part.GetTime = DateTime.Now.ToShortDateString();
                    parts.Add(part);
                    // pc.GetParts.Add(part);
                }
                //pc.SaveChanges();
            }
            return parts;
        }

        public static void ConfimGetPart(BindingList<GetPart> getParts)
        {
            using (PartContext pc = new PartContext())
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
                    pc.Parts.AddOrUpdate(findPart);

                }
                pc.SaveChanges();
            }


        }
    }
}
