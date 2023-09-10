using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ReporteExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog guardar = new SaveFileDialog();
            Excel.Application app = new Excel.Application();
            Excel.Workbook work = app.Workbooks.Open(Path.GetFullPath("club nautico.xlsx"), null, true);
            Excel.Worksheet sheet = work.Worksheets[1];

            int j = 2;
            int n = dataGridView1.Rows.Count;

            for (int i = 0; i < n; i++)
            {
                sheet.Range["A" + j.ToString()].Value = j - 1;
                sheet.Range["B" + j.ToString()].Value = dataGridView1.Rows[i].Cells[0].Value;
                sheet.Range["C" + j.ToString()].Value2 = dataGridView1.Rows[i].Cells[1].Value;
                sheet.Range["D" + j.ToString()].Value2 = dataGridView1.Rows[i].Cells[2].Value;
                sheet.Range["E" + j.ToString()].Value2 = dataGridView1.Rows[i].Cells[3].Value;
                sheet.Range["F" + j.ToString()].Value2 = dataGridView1.Rows[i].Cells[4].Value;
                sheet.Range["G" + j.ToString()].Value2 = dataGridView1.Rows[i].Cells[5].Value;
                sheet.Range["H" + j.ToString()].Value2 = dataGridView1.Rows[i].Cells[6].Value;

                if (i < n-1)
                {
                    sheet.Range["A"+(j + 1).ToString()].EntireRow.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                }
                j++;
            }
            app.Visible = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SqlConnection conexion = new SqlConnection("Server='localhost\\SQLEXPRESS'; Database='club_nautico';Trusted_Connection=True;");
            conexion.Open();
            SqlCommand query = new SqlCommand("select * from barco", conexion);
            SqlDataReader datos = query.ExecuteReader();
            DataSet ds = new DataSet();
            ds.Load(datos,LoadOption.OverwriteChanges,"datos");
            dataGridView1.DataSource = ds.Tables["datos"];
        }
    }
}
