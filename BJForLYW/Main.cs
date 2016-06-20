using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Migrations;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using BJForLYW.DB;
using BJForLYW.Properties;

namespace BJForLYW
{
    public partial class Main : Form
    {
       
        public List<GetPart> GetPartlistFromExcel;
        private PartContext pc = new PartContext();
        /// <summary>
        /// 设备表中选择的要出库的设备
        /// </summary>
        private Part _selectPart;


        public Main()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadPart();
            LoadPutPart();
            LoadGetPart();
        }

        /// <summary>
        ///     加载Part表到datatableview
        /// </summary>
        private void LoadPart()
        {
            pc.Parts.Load();
            partbindingSource1.DataSource = pc.Parts.Local.ToBindingList();
            bindingNavigator1.BindingSource = partbindingSource1;
            PartDtv.AutoGenerateColumns = false;
            _selectPart = null;
        }

        private void 保存SToolStripButton_Click(object sender, EventArgs e)
        {
            DtvSaveAndMBox();
        }
        /// <summary>
        /// 双击表头添加设备到出库表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            long partid = long.Parse(PartDtv.Rows[e.RowIndex].Cells[0].Value.ToString());
            //dataGridView4.AutoGenerateColumns = false;
            selectPutPartDtv.DataSource =
                pc.Parts.Where(s => s.Partid == partid).ToList();
            _selectPart = pc.Parts.First(s => s.Partid == partid);
            PutNumNup_shebei.Maximum = _selectPart.Num;
        }
        /// <summary>
        /// 点击确定出库按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void PutPartBtn_shebei_Click(object sender, EventArgs e)
        {
            if (_selectPart == null)
            {
                MessageBox.Show(Resources.Main_PutPartBtn_shebei_Click_请选择要出库的备件);
                return;
            }
            if (string.IsNullOrEmpty(PutPeopleNameTxt_shebei.Text))
            {
                MessageBox.Show(Resources.Main_PutPartBtn_shebei_Click_请填写出库人);
                return;
            }
            if (PutNumNup_shebei.Value == 0)
            {
                MessageBox.Show(Resources.Main_PutPartBtn_shebei_Click_);
                return;
            }
            string partinfo = "物料编码：" + _selectPart.PartNum + "\n" + "备件名称:" + _selectPart.PartName + "\n" + "备件型号:" +
                              _selectPart.PartType;

            if (MessageBox.Show(partinfo, "确认要出此备件吗？", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                pc.PutParts.Add(ExcelHelper.GenerationPutPartFromPart(_selectPart, (int) PutNumNup_shebei.Value,
                    PutTImeDtp_shebei.Value.ToShortDateString(), PutPeopleNameTxt_shebei.Text, PartRemarks_txt.Text));

                _selectPart.Num = _selectPart.Num - (int) PutNumNup_shebei.Value;
                pc.Parts.AddOrUpdate(_selectPart);
                pc.SaveChanges();
                MessageBox.Show(Resources.Main_PutPartBtn_shebei_Click_成功出库_);
                _selectPart = null;
            }
        }
      

       /// <summary>
       /// 初始化入库表
       /// </summary>

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
        ///     初始化出库表
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
        /// <summary>
        /// 出库表导出excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PutPartToExcel_btn_Click(object sender, EventArgs e)
        {
            ExcelHelper.DataGridViewToExcel(PutPartDtv, "出库导出表");
        }
        /// <summary>
        /// 出库表备件名称查询
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PutPatNameSerach_btn_Click(object sender, EventArgs e)
        {
            string serachtxt = PutPartNameSerach_txt.Text.Trim();
            if (!string.IsNullOrEmpty(serachtxt))
            {
                var serachSource = pc.PutParts.Where(
                    p =>
                        p.PartName.Contains(serachtxt) || p.PartType.Contains(serachtxt) ||
                        p.PartNum.Contains(serachtxt));
                putPartbindingSource1.DataSource = serachSource.ToList();
                PutPartDtv.ResetBindings();
            }
            else
            {
                putPartbindingSource1.DataSource = pc.PutParts.Local.ToBindingList();
            }
        }
        /// <summary>
        /// 出库人查询
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// 出库时间查询
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PutPartTime_btn_Click(object sender, EventArgs e)
        {
            int year = int.Parse(PutPartTimeStart_txt.Text);
            int month = int.Parse(PutPartTimeStop_txt.Text);

            DateTime dt = new DateTime(year, month, 1);
            string ss = dt.ToString("yyyy/M");
            var serachsource = pc.PutParts.Where(p => p.PutTime.StartsWith(ss)).ToList();
            putPartbindingSource1.DataSource = serachsource.ToList();
            PutPartDtv.ResetBindings();
        }

 /// <summary>
 /// 从excel导入入库表,但是没有确认
 /// </summary>
 /// <param name="sender"></param>
 /// <param name="e"></param>
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

        /// <summary>
        /// 入库表的时间查询
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

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
            DtvSaveAndMBox();
        }
        /// <summary>
        /// 保存成功病提示
        /// </summary>
        private void DtvSaveAndMBox()
        {
            pc.SaveChanges();
            MessageBox.Show(Resources.Main_保存SToolStripButton_Click_保存成功);
        }
        /// <summary>
        /// 重新刷新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            //foreach (var pp  in pc.GetParts.Local.ToBindingList())
            //{
            //    pc.Entry(pp).State=EntityState.Unchanged;
            //}
            refresh();
        }

        private void refresh()
        {
            pc.Dispose();
            pc = new PartContext();
            // GetPartBindingSource.DataSource = null;
            LoadGetPart();
            LoadPart();
            LoadPutPart();
            GetPartDtv.ResetBindings();
        }

        /// <summary>
        /// 确认导入的入库表入库,这是才会在设备表进行数量相加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetCofirmToDbToolStripButton1_Click(object sender, EventArgs e)
        {
            if (GetPartlistFromExcel == null)
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

        /// <summary>
        /// 导出出库表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            ExcelHelper.DataGridViewToExcel(PartDtv, "库存导出表");
        }
        /// <summary>
        ///重新全部导入设备表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(Resources.Main_toolStripButton5_Click_, "警告", MessageBoxButtons.OKCancel) ==
                DialogResult.OK)
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

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            ExcelHelper.DataGridViewToExcel(GetPartDtv, "入库导出表");
        }
        /// <summary>
        /// 清理入库表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton6_Click_1(object sender, EventArgs e)
        {
            if (MessageBox.Show("清理功能会清理掉所有的入库记录!\n除非你知道你在干什么,否则不要继续", "警告", MessageBoxButtons.YesNo) ==
                DialogResult.Yes)
            {
                int count = pc.Database.ExecuteSqlCommand("delete  from GetParts");
                MessageBox.Show($"清理掉{count}条记录!");
                // GetPartDtv.ResetBindings();
            }
        }
        /// <summary>
        /// 设备表实时查询
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void partSearchTxt_TextChanged(object sender, EventArgs e)
        {
            var searchTxt = partSearchTxt.Text.Trim();
            if (string.IsNullOrEmpty(searchTxt))
            {
                PartDtv.DataSource = partbindingSource1;
            }
            else
            {
                var searchResult =
                    pc.Database.SqlQuery<Part>(
                        "SELECT * FROM Parts WHERE Parts.PartName LIKE @name1 or Parts.PartType Like @name1 or Parts.PartNum like @name1"
                        , new SQLiteParameter("@name1", "%" + searchTxt + "%")).ToList();
                var tt = pc.Parts.ToList().Intersect(searchResult, new PartComparer());
                // var searchResult = pc.Parts.Where(
                //s => s.PartName.Contains(searchTxt) || s.PartNum.Contains(searchTxt) || s.PartType.Contains(searchTxt))
                //.Distinct();
                //partbindingSource1.DataSource = searchResult.ToList();
                // bindingNavigator1.BindingSource = partbindingSource1;

                PartDtv.DataSource = tt.ToArray(); //
                //todo 直接更改dataview的绑定值,数据量不变,更改bingsource的,数据量都变了
                // dataGridView1.ResetBindings();
            }
            PartDtv.ResetBindings();
        }

        private void 清理数据库ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(Resources.Main_清理数据库ToolStripMenuItem_Click_,"警告",MessageBoxButtons.YesNo)==DialogResult.Yes)
            {
                pc.Database.ExecuteSqlCommand("delete from GetParts");
                pc.Database.ExecuteSqlCommand("delete from Parts");
                pc.Database.ExecuteSqlCommand("delete from PutParts");
                MessageBox.Show("清理成功!");
                refresh();
            }
           
        }

        private void 备份数据库ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string pathCurr = System.Environment.CurrentDirectory;
            string pathstr = Path.Combine(pathCurr, "数据库备份");
            if (!Directory.Exists(pathstr))
            {
                Directory.CreateDirectory(pathstr);
            }
            File.Copy("Part.db",Path.Combine(pathstr,$"part{DateTime.Now.ToString("yyyy_mmmm_dd_hh_mm_ss")}.db"));
            MessageBox.Show("保存成功");
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void 关于ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(Resources.Main_关于ToolStripMenuItem_Click_版权_崔健,"关于",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }
    }
}