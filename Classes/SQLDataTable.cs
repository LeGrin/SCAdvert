using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace SCAdvert.Classes
{
    public class SQLDataTable
    {
        private string Connect = ConfigurationManager.ConnectionStrings["SqlConnect"].ConnectionString;

        public DataTable DGVonPage()
        {
            var con = new SqlConnection(Connect);
            con.Open();

            var ds = new DataSet();
            var dt = new DataTable();
            ds.Tables.Add(dt);

            string query = @"SELECT MediaType, Year, Month ,Week ,Date ,[Start Time] ,[End Time]
                                                ,Sector
                                                ,Category
                                                ,Class
                                                ,Producer
                                                ,Brand
                                                ,Product
                                                ,Copy
                                                ,Market
                                                ,Publishinghouse
                                                ,Distributor
                                                ,PeriodicalType
                                                ,SiteType
                                                ,AdType
                                                ,AdFormat
                                                ,AdSize
                                                ,AudioCodeOutdoor
                                                ,AudioCodePress
                                                ,AudioCodeRadio
                                                ,AudioCodeInternet
                                                ,AdSectionType
                                                ,AdSection
                                                ,AdPosition
                                                ,AdPage
                                                ,Issue
                                                ,AdColor
                                                ,Circulation
                                                ,DisplayPerc
                                                ,Extension
                                                ,AgencyInternet
                                                ,BuyerInternet
                                                ,Damage
                                                ,Direction
                                                ,ProgrammeLocation
                                                ,ProgLocationTypologyVariables
                                                ,Insertions
                                                ,Investment FROM dbo.Import";
            var cmd = new SqlCommand(query, con);

            SqlDataReader dr = cmd.ExecuteReader();

            dt.Load(dr);

            con.Close();

            return dt;
        }
    }
}