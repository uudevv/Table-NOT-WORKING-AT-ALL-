using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace SpreadSheet_Beta_
{
    public partial class CloneTable : Form
    {
        string Name = "";
        string FileName = "";
        DataSet dataSet = new DataSet();
        string Select1;
        string Choose1;
        OleDbConnection Strcn;
        OleDbCommand command1;
        OleDbDataAdapter adapter1;

        public CloneTable()
        {
            InitializeComponent();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog f1 = new OpenFileDialog();
            if (f1.ShowDialog() != DialogResult.Cancel)
            {
                try
                {
                    if (f1.FileName != null)
                    {
                        Name = f1.FileName.Remove(f1.FileName.LastIndexOf("\\"));
                        FileName = f1.FileName.Remove(0, f1.FileName.LastIndexOf("\\") + 1);
                        Select1 = "SELECT * FROM[" + FileName + "]";
                        Strcn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;DataSource=" + Name + ";Extended Properties=text");
                        command1 = new OleDbCommand(Select1, Strcn);
                        adapter1 = new OleDbDataAdapter(command1);

                        Strcn.Open();

                        adapter1.Fill(dataSet);
                        dataGridView1.DataSource = dataSet.Tables[0];

                        Strcn.Close();
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Error:" + ex.Message);
                }
            }
        }
    }
}
