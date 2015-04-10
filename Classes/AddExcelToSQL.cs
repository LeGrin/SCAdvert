using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Windows;

namespace SCAdvert.Classes
{
    public class AddExcelToSQL
    {
        private string GetExcelSheetNames(string ExcelFileName)
        {
            OleDbConnection objConn = null;
            DataTable dt = null;

            try
            {
                string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFileName +
                                    ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;\"";

                objConn = new OleDbConnection(connString);
                objConn.Open();

                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                string excelSheets = null;

                foreach (DataRow row in dt.Rows)
                {
                    if (row["TABLE_NAME"].ToString().Contains("$"))
                    {
                        excelSheets = row["TABLE_NAME"].ToString();
                    }
                }

                return excelSheets;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                // Clean up.
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        public DataTable oleDBTable(string ExcelPath)
        {
            var dt = new DataTable();

            try
            {
                var connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelPath +
                                    ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;\"";
                var conn = new OleDbConnection(connectString);

                conn.Open();

                string sheetname = GetExcelSheetNames(ExcelPath);

                var sheetNameCute = sheetname.Substring(0, sheetname.IndexOf('$'));

                string sheet = "[" + sheetNameCute + "$" + "]";
                //string sheet = "[" + ListSheetInExcel(ExcelPath) + "]";



                var da = new OleDbDataAdapter(@"Select * From" + sheet + @"WHERE [MediaType] IS NOT NULL AND 
                                                [Year] IS NOT NULL AND [Month] IS NOT NULL AND [Week] IS NOT NULL AND [Date] IS NOT NULL AND 
                                                [Start Time] IS NOT NULL AND [End Time] IS NOT NULL AND [Sector] IS NOT NULL AND [Category] IS NOT NULL AND 
                                                [Class] IS NOT NULL AND [Producer] IS NOT NULL AND [Brand] IS NOT NULL AND [Product] IS NOT NULL AND 
                                                [Copy] IS NOT NULL AND [Market] IS NOT NULL AND [Publishing house] IS NOT NULL AND [Distributor] IS NOT NULL AND 
                                                [Periodical type] IS NOT NULL AND [Site type] IS NOT NULL AND [Ad Type] IS NOT NULL AND [Ad Format] IS NOT NULL AND 
                                                [Ad Size] IS NOT NULL AND [Audio code Outdoor] IS NOT NULL AND [Audio code Press] IS NOT NULL AND 
                                                [Audio code Radio] IS NOT NULL AND [Audio code Internet] IS NOT NULL AND [Ad Section Type] IS NOT NULL AND 
                                                [Ad Section] IS NOT NULL AND [Ad Position] IS NOT NULL AND [Ad Page] IS NOT NULL AND 
                                                [Issue No] IS NOT NULL AND [Ad Color] IS NOT NULL AND [Circulation] IS NOT NULL AND 
                                                [Display Perc] IS NOT NULL AND [Extension] IS NOT NULL AND [Agency Internet] IS NOT NULL AND 
                                                [Buyer Internet] IS NOT NULL AND [Damage] IS NOT NULL AND [Direction] IS NOT NULL AND
                                                [Programme/Location] IS NOT NULL AND [Prog/Location Typology\Variables] IS NOT NULL AND 
                                                [Insertions] IS NOT NULL AND [Investment] IS NOT NULL", conn);

                da.Fill(dt);
                da.Dispose();

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return dt;
        }

        public int TotalRows(string ExcelPath)
        {

            string sheetname = GetExcelSheetNames(ExcelPath);

            var sheetNameCute = sheetname.Substring(0, sheetname.IndexOf('$'));

            string sheet = "[" + sheetNameCute + "$" + "]";

            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelPath +
                                    ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;\"";
            var conn = new OleDbConnection(ConnectionString);

            conn.Open();

            var cmd = new OleDbCommand();

            cmd.CommandText = @"Select COUNT(*) From" + sheet + @"WHERE [MediaType] IS NOT NULL";

            cmd.Connection = conn;

            int counter = (int)cmd.ExecuteScalar();

            conn.Close();

            return counter;
        }
    }
}