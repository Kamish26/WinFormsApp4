using Microsoft.SqlServer.Server;
using System.Data.SqlClient;
using System.Data;
using Microsoft.Data.SqlClient; // Nugit Package Sql Client
using static System.ComponentModel.Design.ObjectSelectorEditor;
using System.Text;
using System.IO;

namespace WinFormsApp4
{
    public partial class Form1 : Form
    {

        public static string beginD;
        public static string endD;


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string ConnectionString = @"Data Source=HQ-W10-L111\SQLEXPRESS; Initial Catalog=AdventureWorks2022; User ID = Kamish; Password = Welcome123!; TrustServerCertificate=True";
            SqlConnection connection = new SqlConnection(ConnectionString);


        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ConnectionString1 = @"Data Source=HQ-W10-L111\SQLEXPRESS; Initial Catalog=AdventureWorks2022; User ID = Kamish; Password = Welcome123!; TrustServerCertificate=True";
            SqlConnection connection = new SqlConnection(ConnectionString1);
            connection.Open();

            beginD = textBox1.Text;
            endD = textBox2.Text;


            SqlCommand command1 = new SqlCommand("SELECT TOP 10 Production.Product.ProductID, Name, SUM(OrderQty) AS 'Order-Quantity' FROM Production.Product INNER JOIN Sales.SalesOrderDetail ON Production.Product.ProductID = Sales.SalesOrderDetail.ProductID INNER JOIN" +
                " Sales.SalesOrderHeader ON Sales.SalesOrderDetail.SalesOrderID = Sales.SalesOrderHeader.SalesOrderID WHERE Sales.SalesOrderHeader.OrderDate >'" + beginD + "'AND Sales.SalesOrderHeader.OrderDate " +
                " <'" + endD + "'GROUP BY Production.Product.ProductID, Production.Product.Name ORDER BY SUM(OrderQty) DESC", connection);


            DataTable dt = new DataTable();

            SqlDataAdapter adapter = new SqlDataAdapter(command1);
            adapter.Fill(dt);

            dataGridView1.DataSource = dt;
            dataGridView1.Visible = true;

            button2.Visible = true;

            connection.Close();




        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "CSV (*.csv)|*.csv";
                saveFile.FileName = "Datagrid.csv";
                bool fileError = false;
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(saveFile.FileName))
                    {
                        try
                        {
                            File.Delete(saveFile.FileName);
                        }
                        catch (IOException ex)
                        {
                            fileError = true;
                            MessageBox.Show("It wasn't possible to write the data to the disk." + ex.Message);
                        }
                    }
                    if (!fileError)
                    {
                        try
                        {
                            int countCol = dataGridView1.ColumnCount;
                            string colName = "";
                            string[] outCsv = new string[dataGridView1.Rows.Count + 1];
                            for (int i = 0; i < countCol; i++)
                            {
                                colName = colName + dataGridView1.Columns[i].HeaderText.ToString() + ",";
                            }
                            outCsv[0] += colName;

                            for (int i = 1; i < dataGridView1.Rows.Count; i++)
                            {
                                for (int k = 0; k < countCol; k++)
                                {
                                    outCsv[i] += dataGridView1.Rows[i - 1].Cells[k].Value.ToString() + ",";

                                }
                            }

                            File.WriteAllLines(saveFile.FileName, outCsv, Encoding.UTF8);
                            MessageBox.Show("Data has been exported");
                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show("There was an error " + ex.Message);
                        }
                    }
                }
            }
        }
    }

}