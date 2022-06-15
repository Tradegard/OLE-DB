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
using System.Data.SqlClient;

namespace OLE_DB
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            string ConnectionString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=C:\\USERS\\KUPTSOV_AE\\SOURCE\\REPOS\\OLE DB\\OLE DB\\DATABASE1.MDF;Integrated Security=True;Connect Timeout=5";
            SqlConnection con = new SqlConnection(ConnectionString);

            con.Open();

            SqlCommand cmd2 = new SqlCommand("select * from dbo.KKC1_DB", con);
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = cmd2;

            DataTable dt2 = new DataTable();
            sda.Fill(dt2);

            BindingSource bSource = new BindingSource();
            bSource.DataSource = dt2;
            dataGridView2.DataSource = bSource;
            sda.Update(dt2);

            con.Close();
        }

        
        public void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                CheckFileExists = true,
                CheckPathExists = true,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }

            string conneсtionString = null;
            string dataSource = textBox1.Text;
            conneсtionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + dataSource + ";" + "Extended Properties='Excel 8.0;HDR=Yes'"; ;
            OleDbConnection cnn = new OleDbConnection(conneсtionString);
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = cnn;
            cnn.Open();
            
            DataTable dtExcelSchema;
            dtExcelSchema = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            comboBox1.DataSource = dtExcelSchema;
            comboBox1.DisplayMember = "TABLE_NAME";

            string sheetName = comboBox1.GetItemText(this.comboBox1.SelectedItem);
            string sql = $"SELECT * FROM [{sheetName}]";
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            ds.Tables.Add(dt);
            OleDbDataAdapter da;
            da = new OleDbDataAdapter(sql, cnn);
            da.Fill(dt);
            dataGridView1.DataSource = dt;

            cnn.Close();

        }
        
        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            string sheetName = comboBox1.GetItemText(this.comboBox1.SelectedItem);
            string sql = $"SELECT * FROM [{sheetName}]";
            
            string conneсtionString = null;
            string dataSource = textBox1.Text;
            conneсtionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + dataSource + ";" + "Extended Properties='Excel 8.0;HDR=Yes'"; ;
            OleDbConnection cnn = new OleDbConnection(conneсtionString);
            cnn.Open();

            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            ds.Tables.Add(dt);
            OleDbDataAdapter da = new OleDbDataAdapter(sql, cnn);
            da.Fill(dt);
            dataGridView1.DataSource = dt;

            cnn.Close();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=|DataDirectory|DATABASE1.MDF;Integrated Security=True;Connect Timeout=5";
            SqlConnection con2 = new SqlConnection(ConnectionString);

            con2.Open();
            SqlCommand cmd3 = new SqlCommand("insert into [KKC1_DB] (NPLV, VES_BEF, VIDERZ, PREDICT) VALUES (@NPLV, @VES_BEF, @VIDERZ, @PREDICT)", con2);
            cmd3.Parameters.AddWithValue("@NPLV", textBox2.Text);
            cmd3.Parameters.AddWithValue("@VES_BEF", textBox3.Text);
            cmd3.Parameters.AddWithValue("@VIDERZ", textBox4.Text);
            cmd3.Parameters.AddWithValue("@PREDICT", 123);
            cmd3.ExecuteNonQuery();
            dataGridView2.Update();
            dataGridView2.Refresh();
            
            con2.Close();
            
            
        }
    }
}
