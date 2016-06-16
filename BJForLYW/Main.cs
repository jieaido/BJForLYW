﻿using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BJForLYW.DB;
using NPOI.SS.Formula.Functions;

namespace BJForLYW
{
    public partial class Main : Form
    {
        private IEnumerable<Part> allpartlist;
        private readonly PartContext pc = new PartContext();
        private Part selectPart=new Part();
       

        public Main()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            pc.Parts.Load();
            partbindingSource1.DataSource = pc.Parts.Local.ToBindingList();


            bindingNavigator1.BindingSource = partbindingSource1;
        }

        private void 保存SToolStripButton_Click(object sender, EventArgs e)
        {
            pc.SaveChanges();
        }

       

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var filename = openFileDialog1.FileName;
                var getPartlist = ExcelHelper.GetPartFromExcel(filename);
                pc.GetParts.AddRange(getPartlist);
                GetPartBindingSource.DataSource = pc.GetParts.Local.ToBindingList();
                dataGridView2.AutoGenerateColumns = true;

                //MessageBox.Show(filename);
            }
        }


        private void bindingNavigator2_RefreshItems(object sender, EventArgs e)
        {
        }

        private void splitContainer3_Panel2_Paint(object sender, PaintEventArgs e)
        {
        }

        private void 保存SToolStripButton1_Click_1(object sender, EventArgs e)
        {
            var ss = pc.GetParts.Local.ToBindingList();
            ExcelHelper.ConfimGetPart(ss);
            pc.SaveChanges();
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            //dataGridView1.DataSource = null;
           


        }

    

        private void PartbindingNavigator1_RefreshItems(object sender, EventArgs e)
        {
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
         
            
        }

        private void dataGridView1_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            long partid =long.Parse(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString()) ;
            dataGridView4.AutoGenerateColumns = false;
            dataGridView4.DataSource =
                pc.Parts.Where(s => s.Partid == partid ).ToList();
            selectPart= pc.Parts.First(s => s.Partid == partid);
            PutNumNup_shebei.Maximum = selectPart.Num;


        }

        private void FindPartCom_Shebei_Validated(object sender, EventArgs e)
        {
          
        }

        private void PutPartBtn_shebei_Click(object sender, EventArgs e)
        {
            string sss = "%电磁%";

            //  var ss = pc.Database.SqlQuery<Part>("SELECT Parts.PartName FROM Parts WHERE Parts.PartName LIKE '%电磁%' ").ToList();//这样写会报错,前面的类型必须和后面查找的属性相配,所以改成了下面的
           
            string trtr = "MTL隔离器";
            bool ssd = trtr.Contains("阿");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            var searchTxt = textBox1.Text.Trim();
            if (string.IsNullOrEmpty(searchTxt))
            {
                dataGridView1.DataSource = partbindingSource1;
            }
            else
            {
           
                var searchResult = pc.Database.SqlQuery<Part>("SELECT * FROM Parts WHERE Parts.PartName LIKE @name1 or Parts.PartType Like @name1 or Parts.PartNum like @name1"
               , new SQLiteParameter("@name1", "%" + searchTxt + "%")).ToList();
                var tt = pc.Parts.ToList().Intersect(searchResult, new PartComparer());
               // var searchResult = pc.Parts.Where(
               //s => s.PartName.Contains(searchTxt) || s.PartNum.Contains(searchTxt) || s.PartType.Contains(searchTxt))
               //.Distinct();
                //partbindingSource1.DataSource = searchResult.ToList();
                // bindingNavigator1.BindingSource = partbindingSource1;

                dataGridView1.DataSource = tt.ToArray();//
                //todo 直接更改dataview的绑定值,数据量不变,更改bingsource的,数据量都变了
                // dataGridView1.ResetBindings();
            }
            dataGridView1.ResetBindings();
            //comboBox1.BeginUpdate();

            //foreach (var part in searchResult)
            //{
            //    comboBox1.Items.Add(part);

            //}
            //comboBox1.DisplayMember = "PartName";
            //comboBox1.ValueMember = "partid";
            //comboBox1.EndUpdate();
        }
    }
}