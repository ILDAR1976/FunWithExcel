using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FunWithExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void load_sheet_excel_Click(object sender, EventArgs e)
        {
            String filePath = @"V:\projects\cs\example\10_result.xls";
            String connectionString = String.Empty;
            
            info.Clear();

            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("ru-RU", true);
                Thread.CurrentThread.CurrentCulture.ClearCachedData();

                if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xlsx")
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + "No" + ";IMEX=0\""; 
                else
                    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=" + "No" + ";IMEX=0\"";

                OleDbConnection con = new System.Data.OleDb.OleDbConnection(connectionString);
                con.Open(); 
                DataTable schemaTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                DataRow schemaRow = schemaTable.Rows[0]; 
                string sheet = schemaRow["TABLE_NAME"].ToString(); 
                
                if (!sheet.EndsWith("_")) { string query = "SELECT * FROM [" + sheet + "]"; 
                    OleDbDataAdapter daexcel = new OleDbDataAdapter(query, con);
                    schemaTable.Locale = CultureInfo.CurrentCulture;
                    daexcel.Fill(schemaTable);
                       
                    DataSet ds = new DataSet();
                    ds.Tables.Add(schemaTable);
                    int c = 0;
                    foreach (DataTable dt in ds.Tables)
                    {
                        
                        // перебор всех строк таблицы
                        foreach (DataRow row in dt.Rows)
                        {
                            // получаем все ячейки строки
                            var cells = row.ItemArray;
                            if ((++c) != 1) { 
                                foreach (object cell in cells) info.AppendText(cell.ToString().Trim() + " ");
                                info.AppendText("\n");
                            }

                        }

                    }

                }

                con.Close();
                con.Dispose();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
