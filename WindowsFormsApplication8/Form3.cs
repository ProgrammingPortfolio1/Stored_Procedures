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
using System.Configuration;

namespace WindowsFormsApplication8
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        public String connectionString = ConfigurationManager.ConnectionStrings["connstr"].ConnectionString;


        Form2 f2 = new Form2();

        public void GridFill()
        {
            using (SqlConnection sqlCon = new SqlConnection(connectionString))
            {
                sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("ViewAll1", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);
                f2.dataGridView1.DataSource = dtbl;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String uni = this.textBox1.Text;
            String uname = this.textBox2.Text;
            String pword = this.textBox3.Text;

            Guid gu = Guid.NewGuid();

            String uni2 = gu.ToString("D").ToUpper();

            if (uname != "" && pword != "")
            {

                try
                {

                    using (SqlConnection sqlCon = new SqlConnection(connectionString))
                    {
                        sqlCon.Open();
                        SqlCommand sqlCmd = new SqlCommand("AddOrEdit1", sqlCon);

                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        sqlCmd.Parameters.AddWithValue("@txt", uni);
                        sqlCmd.Parameters.AddWithValue("@userid", uni2);
                        sqlCmd.Parameters.AddWithValue("@username", uname);
                        sqlCmd.Parameters.AddWithValue("@userpassword", pword);

                        sqlCmd.ExecuteNonQuery();

                        MessageBox.Show("Submitted Successfully");

                        GridFill();

                        this.Close();
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("username alleady exists");

                    MessageBox.Show(ex.ToString());
                }

            }
            else
            {

                MessageBox.Show("All fields are required");

            }

        }

    }
}
