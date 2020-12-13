using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication8
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        public String connectionString = ConfigurationManager.ConnectionStrings["connstr"].ConnectionString;


        private void button1_Click(object sender, EventArgs e)
        {

            Form3 f3 = new Form3();
            f3.ShowDialog();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            // string t = this.dataGridView1.CurrentRow.Cells[0].Value.ToString();

            string t = this.dataGridView1.CurrentRow.Cells[1].Value.ToString();


            try
            {

                using (SqlConnection sqlCon = new SqlConnection(connectionString))
                {
                    sqlCon.Open();
                    SqlCommand sqlCmd = new SqlCommand("DeleteByID1", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;

                    sqlCmd.Parameters.AddWithValue("@t", t);

                    sqlCmd.ExecuteNonQuery();

                    MessageBox.Show("Submitted Successfully");

                    GridFill();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                //     MessageBox.Show(ex.ToString());
            }

        }


        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Form3 f3 = new Form3();
            f3.textBox1.Text = this.dataGridView1.CurrentRow.Cells[0].Value.ToString();
            f3.textBox2.Text = this.dataGridView1.CurrentRow.Cells[1].Value.ToString();
            f3.textBox3.Text = this.dataGridView1.CurrentRow.Cells[2].Value.ToString();
            f3.ShowDialog();
        }

        public void GridFill()
        {
            using (SqlConnection sqlCon = new SqlConnection(connectionString))
            {
                sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("ViewAll1", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);
                dataGridView1.DataSource = dtbl;
            }

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.textBox1.Text = "";

            GridFill();

            if (this.dataGridView1.Rows.Count == 0)
            {
                this.button2.Enabled = false;
                //this.button3.Enabled = false;
                //this.textBox1.Enabled = false;
            }
            else
            {
                this.button2.Enabled = true;
                //this.button3.Enabled = true;
                //this.textBox1.Enabled = true;
            }

        }

        private void Form2_Activated(object sender, EventArgs e)
        {
            this.textBox1.Text = "";

            GridFill();

            if (this.dataGridView1.Rows.Count == 0)
            {
                this.button2.Enabled = false;
                //this.button3.Enabled = false;
                //this.textBox1.Enabled = false;
            }
            else
            {
                this.button2.Enabled = true;
                //this.button3.Enabled = true;
                //this.textBox1.Enabled = true;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            using (SqlConnection sqlCon = new SqlConnection(connectionString))
            {
                sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("SearchByValue1", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                sqlDa.SelectCommand.Parameters.AddWithValue("@SearchValue", this.textBox1.Text);
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);
                dataGridView1.DataSource = dtbl;

                if (this.dataGridView1.Rows.Count == 0)
                {
                    this.button2.Enabled = false;
                }
                else
                {
                    this.button2.Enabled = true;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }


        private void ExportToExcel()
        {

            // Creating an Excel object.
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {
                //worksheet = workbook.Sheets["Sheet1"];

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "ExportedFromDatGrid";

                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                //Loop through each row and read value from each column.
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }

                //Getting the location and file name of the excel to save from user.
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel 97-2003 (*.xls)|*.xls|Excel (*.xlsx)|*.xlsx";
                saveDialog.FilterIndex = 2;

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Export Successful");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                workbook.Close(false, Type.Missing, Type.Missing);

                excel.Quit();
                GC.Collect();

                Marshal.FinalReleaseComObject(worksheet);

                Marshal.FinalReleaseComObject(workbook);

                Marshal.FinalReleaseComObject(excel);

                //excel.Quit();
                //workbook = null;
                //excel = null;
            }
        }
    }
}
