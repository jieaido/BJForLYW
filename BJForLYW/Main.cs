using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Migrations;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BJForLYW.DB;
using BJForLYW.Properties;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;

namespace BJForLYW
{
    public partial class Main : Form
    {
        private IEnumerable<Part> allpartlist;
        private  PartContext pc = new PartContext();
        private Part selectPart = null;
        public List<GetPart> GetPartlistFromExcel;


        public Main()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadPart();
            LoadPutPart();
        }
        /// <summary>
        ///加载Part表到datatableview
        /// </summary>
        private void LoadPart()
        {
            pc.Parts.Load();
            partbindingSource1.DataSource = pc.Parts.Local.ToBindingList();


            bindingNavigator1.BindingSource = partbindingSource1;
            PartDtv.AutoGenerateColumns = false;
        }

        private void 保存SToolStripButton_Click(object sender, EventArgs e)
        {
            pc.SaveChanges();
        }

       

        private void button2_Click(object sender, EventArgs e)
        {
            
        }


        private void bindingNavigator2_RefreshItems(object sender, EventArgs e)
        {
        }

        private void splitContainer3_Panel2_Paint(object sender, PaintEventArgs e)
        {
        }

        private void 保存SToolStripButton1_Click_1(object sender, EventArgs e)
        {
           
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
            long partid =long.Parse(PartDtv.Rows[e.RowIndex].Cells[0].Value.ToString()) ;
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
            if (selectPart==null)
            {
                MessageBox.Show(Resources.Main_PutPartBtn_shebei_Click_请选择要出库的备件);
                return;
            }
            if (string.IsNullOrEmpty(PutPeopleNameTxt_shebei.Text))
            {
                MessageBox.Show(Resources.Main_PutPartBtn_shebei_Click_请填写出库人);
                return;
            }
            if (PutNumNup_shebei.Value==0)
            {
                MessageBox.Show(Resources.Main_PutPartBtn_shebei_Click_);
                return;
                
            }
            string partinfo = "物料编码：" + selectPart.PartNum + "\n" + "备件名称:" + selectPart.PartName + "\n" + "备件型号:" +
                              selectPart.PartType;
            if (MessageBox.Show(partinfo, "确认要出此备件吗？",MessageBoxButtons.OKCancel)==DialogResult.OK)
            {
                pc.PutParts.Add(ExcelHelper.GenerationPutPartFromPart(selectPart, (int)PutNumNup_shebei.Value,
                PutTImeDtp_shebei.Value.ToShortDateString(), PutPeopleNameTxt_shebei.Text));
                selectPart.Num = selectPart.Num - (int)PutNumNup_shebei.Value;
                pc.Parts.AddOrUpdate(selectPart);
                pc.SaveChanges();
                MessageBox.Show("成功出库!");
                selectPart = null;
            }
            

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            var searchTxt = textBox1.Text.Trim();
            if (string.IsNullOrEmpty(searchTxt))
            {
                PartDtv.DataSource = partbindingSource1;
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

                PartDtv.DataSource = tt.ToArray();//
                //todo 直接更改dataview的绑定值,数据量不变,更改bingsource的,数据量都变了
                // dataGridView1.ResetBindings();
            }
            PartDtv.ResetBindings();
            //comboBox1.BeginUpdate();

            //foreach (var part in searchResult)
            //{
            //    comboBox1.Items.Add(part);

            //}
            //comboBox1.DisplayMember = "PartName";
            //comboBox1.ValueMember = "partid";
            //comboBox1.EndUpdate();
        }

        private void tabControl1_Enter(object sender, EventArgs e)
        {
            LoadGetPart();
        }

        private void LoadGetPart()
        {
            #region 时间初始化器
            for (int i = -3; i < 3; i++)
            {
                GetStripCbb_year.Items.Add(DateTime.Now.Year + i);
            }
            GetStripCbb_year.SelectedItem = DateTime.Now.Year;

            for (int i = 1; i < 13; i++)
            {
                GetStripCbb_month.Items.Add(i);
            }
            GetStripCbb_month.SelectedItem = DateTime.Now.Month;    
            # endregion
            pc.GetParts.Load();
            GetPartBindingSource.DataSource = pc.GetParts.Local.ToBindingList();
            GetbindingNavigator2.BindingSource = GetPartBindingSource;
        }

        /// <summary>
        /// 加载出库表
        /// </summary>
        private void LoadPutPart()
        {
            pc.PutParts.Load();
           
            putPartbindingSource1.DataSource = pc.PutParts.Local.ToBindingList();
            putPartbindingNavigator2.BindingSource = putPartbindingSource1;


          
            #region 初始化时间选择器
            for (int i = -3; i < 3; i++)
            {
                PutPartTimeStart_txt.Items.Add(DateTime.Now.Year + i);
            }
            PutPartTimeStart_txt.SelectedItem = DateTime.Now.Year;

            for (int i = 1; i < 13; i++)
            {
                PutPartTimeStop_txt.Items.Add(i);
            }
            PutPartTimeStop_txt.SelectedItem = DateTime.Now.Month; 
            #endregion
      
            // PutPartDtv.AutoGenerateColumns = true;
            // PutPartDtv.DataSource = putPartbindingSource1;
        }

        private void PutPartToExcel_btn_Click(object sender, EventArgs e)
        {
           ExcelHelper.DataGridViewToExcel(PutPartDtv, "出库导出表");
        }

        private void PutPatNameSerach_btn_Click(object sender, EventArgs e)
        {
            string serachtxt = PutPartNameSerach_txt.Text.Trim();
            if (!string.IsNullOrEmpty(serachtxt))
            {
                var serachSource= pc.PutParts.Where(
                    p =>
                        p.PartName.Contains(serachtxt) || p.PartType.Contains(serachtxt) ||
                        p.PartNum.Contains(serachtxt));
                putPartbindingSource1.DataSource = serachSource.ToList();
                PutPartDtv.ResetBindings();
            }
            else
            {
                putPartbindingSource1.DataSource= pc.PutParts.Local.ToBindingList();
            }

        }

        private void PutPartPeopleName_btn_Click(object sender, EventArgs e)
        {
            string serachtxt = PutPartPeopleName_txt.Text.Trim();
            if (!string.IsNullOrEmpty(serachtxt))
            {
                var serachSource = pc.PutParts.Where(
                    p =>
                        p.PutPeopleName.Contains(serachtxt));
                putPartbindingSource1.DataSource = serachSource.ToList();
                PutPartDtv.ResetBindings();
            }
            else
            {
                putPartbindingSource1.DataSource = pc.PutParts.Local.ToBindingList();
            }
        }

        private void PutPartTime_btn_Click(object sender, EventArgs e)
        {
            int year = int.Parse(PutPartTimeStart_txt.Text);
            int month = int.Parse(PutPartTimeStop_txt.Text);
            
            DateTime dt=new DateTime(year,month,1);
            string ss = dt.ToString("yyyy/M");
             var serachsource= pc.PutParts.Where(p => p.PutTime.StartsWith(ss)).ToList();
            putPartbindingSource1.DataSource = serachsource.ToList();
            PutPartDtv.ResetBindings();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
           
        }

        private void bindingNavigator2_RefreshItems_1(object sender, EventArgs e)
        {

        }

        private void 打开OToolStripButton1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var filename = openFileDialog1.FileName;
                GetPartlistFromExcel = ExcelHelper.GetgetPartTableFromExcel(filename);
                //pc.GetParts.AddRange(getPartlist);
                GetPartBindingSource.DataSource = GetPartlistFromExcel;
                //  GetPartDtv.AutoGenerateColumns = true;

                //MessageBox.Show(filename);
            }
        }

        

        private void toolStripDropDownButton1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            int year = int.Parse(GetStripCbb_year.Text);
            int month = int.Parse(GetStripCbb_month.Text);

            DateTime dt = new DateTime(year, month, 1);
            string ss = dt.ToString("yyyy/M");
            var serachsource = pc.GetParts.Where(p => p.GetTime.StartsWith(ss)).ToList();
            GetPartBindingSource.DataSource = serachsource.ToList();
            GetPartDtv.ResetBindings();

        }

        private void 保存SToolStripButton2_Click(object sender, EventArgs e)
        {
            pc.SaveChanges();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

            foreach (var pp  in pc.GetParts.Local.ToBindingList())
            {
                pc.Entry(pp).State=EntityState.Unchanged;
            }
            
            // GetPartBindingSource.DataSource = null;
            LoadGetPart();
            GetPartDtv.ResetBindings();

            
        }

        private void GetCofirmToDbToolStripButton1_Click(object sender, EventArgs e)
        {
            if (GetPartlistFromExcel==null)
            {
                MessageBox.Show(Resources.Main_GetCofirmToDbToolStripButton1_Click_请选择要导入的入库文件);
                return;
            }
            ExcelHelper.ConfimGetPart(GetPartlistFromExcel, pc);
            MessageBox.Show(Resources.Main_GetCofirmToDbToolStripButton1_Click_导入成功);
            // pc.SaveChanges();
            GetPartlistFromExcel = null;
            LoadGetPart();
        }

        private void PartDtv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            ExcelHelper.DataGridViewToExcel(PartDtv, "库存导出表");
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(Resources.Main_toolStripButton5_Click_,"警告",MessageBoxButtons.OKCancel)==DialogResult.OK)
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    var filename = openFileDialog1.FileName;
                    var parts = ExcelHelper.GetPartTableFromExcel(filename);
                    pc.Parts.RemoveRange(pc.Parts.ToList());
                  
                  
                    
                    partbindingSource1.DataSource = parts;
                    pc.Parts.AddRange(parts);
                    pc.SaveChanges();



                }
            }
        }
    }
}