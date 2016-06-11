using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BJForLYW.DB;

namespace BJForLYW
{
    public partial class Main : Form
    {
        PartContext pc=new PartContext();
        private IEnumerable<Part> allpartlist;
        public Main()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            pc.Parts.Load();
            bindingSource1.DataSource = pc.Parts.Local.ToBindingList();
           
           
           bindingNavigator1.BindingSource = bindingSource1;
           
            

        }

        private void 保存SToolStripButton_Click(object sender, EventArgs e)
        {
            pc.SaveChanges();

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void splitContainer2_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var filename = openFileDialog1.FileName;
                ExcelHelper.GetPartFromExcel(filename);
                //MessageBox.Show(filename);
            }
        }
    }
}
