using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;
using SCAdvert.Classes;
using System.Configuration;
using Path = System.IO.Path;

namespace SCAdvert
{
    /// <summary>
    /// Interaction logic for AdminForm.xaml
    /// </summary>
    public partial class AdminForm : Window
    {
        public AdminForm()
        {
            InitializeComponent();
        }

        private string Connect = ConfigurationManager.ConnectionStrings["SqlConnect"].ConnectionString;

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BtnAddToSQL_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                AddToSql(TxtPathName.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static string UPLOADFOLDER = "Uploads";
        AddExcelToSQL ExcToSQL = new AddExcelToSQL();

        // Convert Bytes to Megabytes
        static double ConvertBytesToMegabytes(long bytes)
        {
            return (bytes / 1024f) / 1024f;
        }

        private void BtnAddToGroupImport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new OpenFileDialog() { Filter = @"Excel Files|*.xlsx|All Files|*.*" };
                if (openFileDialog.ShowDialog() == true) ;

                BtnAddToSQL.IsEnabled = true;
                BtnAddToSQL.Foreground = Brushes.White;

                TxtPathName.Text = openFileDialog.FileName;
                string ExcelPath = openFileDialog.FileName;

                FileInfo fi = new FileInfo(ExcelPath);

                FileNotAccept.Visibility = Visibility.Hidden;
                FileSize1.Visibility = Visibility.Visible;

                // Now convert to a string in megabytes.
                string s = ConvertBytesToMegabytes(fi.Length).ToString("0.00");
                FileSize1.Content = s + " MB";

                LabelNameFile1.Visibility = Visibility.Visible;
                LabelNameFile1.Content = fi.Name;

                ProgressBar1.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
   
        //Create a Delegate that matches 
        //the Signature of the ProgressBar's SetValue method
        private delegate void UpdateProgressBarDelegate(DependencyProperty dp, Object value);

        private void AdministratorForm_Loaded(object sender, RoutedEventArgs e)
        {
            LabelNameFile1.Visibility = Visibility.Hidden;
            LabelNameFile2.Visibility = Visibility.Hidden;
            LabelNameFile3.Visibility = Visibility.Hidden;
            LabelNameFile4.Visibility = Visibility.Hidden;
            LabelCurrentRow.Visibility = Visibility.Hidden;
            LabelTotalRows.Visibility = Visibility.Hidden;
            LableSlash.Visibility = Visibility.Hidden;

            FileSize1.Visibility = Visibility.Hidden;
            FileSize2.Visibility = Visibility.Hidden;
            FileSize3.Visibility = Visibility.Hidden;
            FileSize4.Visibility = Visibility.Hidden;

            ProgressBar1.Visibility = Visibility.Hidden;
            ProgressBar2.Visibility = Visibility.Hidden;
            ProgressBar3.Visibility = Visibility.Hidden;
            ProgressBar4.Visibility = Visibility.Hidden;

            if (ProgressBar1.Visibility == Visibility.Hidden)
            {
                BtnAddToSQL.IsEnabled = false;
                BtnAddToSQL.Foreground = Brushes.Gray;
            }

  
        }

        // Пусть width – это ширина всей полосы загрузки. 
        // numOfStrings – количество строк. 
        // currentString – это та строка, которая загрузилась последней либо грузится сейчас. 
        // Найдём x - то есть ширину заполненной полосы загрузки

        // x = currentString * width / numOfStrings

        public void AddToSql(string ExcelPath)
        {
            try
            {
                var con = new SqlConnection(Connect);
                con.Open();

                var dataTable = ExcToSQL.oleDBTable(ExcelPath);

                //Configure the ProgressBar
                ProgressBar1.Minimum = 0;

                int totalnumbersrow = ExcToSQL.TotalRows(ExcelPath);

                LabelTotalRows.Content = totalnumbersrow;

                ProgressBar1.Maximum = totalnumbersrow;

                //MessageBox.Show(totalnumbersrow.ToString());

                ProgressBar1.Value = 0;

                //Stores the value of the ProgressBar
                double value = 0;

                // to the ProgressBar's SetValue method.
                UpdateProgressBarDelegate updatePbDelegate = ProgressBar1.SetValue;

                //Tight Loop: Loop until the ProgressBar.Value reaches the max

                LabelCurrentRow.Visibility = Visibility.Visible;
                LabelTotalRows.Visibility = Visibility.Visible;
                LableSlash.Visibility = Visibility.Visible;

                foreach (DataRow dataRow in dataTable.Rows)
                {

                    int currentString = dataTable.Rows.IndexOf(dataRow);

                    //var x = currentString * 100 / totalnumbersrow;

                    LabelCurrentRow.Content = currentString;

                    var MediaType = dataRow["MediaType"].ToString();

                    var Year = dataRow["Year"].ToString();
                    DateTime YearDate = DateTime.ParseExact(Year, "yyyy", CultureInfo.InvariantCulture);
                    var YearDateString = YearDate.ToString("yyyy");
             
                    var Month = dataRow["Month"].ToString();
                    var Week = dataRow["Week"].ToString();
                    
                    var Date = dataRow["Date"].ToString();
                    DateTime DateDate = DateTime.ParseExact(Date, "dd.MM.yyyy H:mm:ss", CultureInfo.InvariantCulture);
                    var DateDateString = DateDate.ToString("dd.MM.yyyy");

                    var timeFormat = "dd.MM.yyyy H:mm:ss";
              
                        var StartTime = dataRow["Start Time"].ToString();
                        DateTime StartTimeDate = DateTime.ParseExact(StartTime, timeFormat, CultureInfo.InvariantCulture);
                        var StartTimeDateString = StartTimeDate.ToString("H:mm:ss");
                     
                        var EndTime = dataRow["End Time"].ToString();
                        DateTime EndTimeDate = DateTime.ParseExact(EndTime, timeFormat, CultureInfo.InvariantCulture);
                        var EndTimeDateString = EndTimeDate.ToString("H:mm:ss");

                    var Sector = dataRow["Sector"].ToString();
                    var Category = dataRow["Category"].ToString();
                    var Class = dataRow["Class"].ToString();
                    var Producer = dataRow["Producer"].ToString(); var producer = Producer.Replace("'", "\""); // Замена ' на "
                    var Brand = dataRow["Brand"].ToString(); var brand = Brand.Replace("'", "\"");
                    var Product = dataRow["Product"].ToString(); var product = Product.Replace("'", "\""); // Замена ' на "
                    var Copy = dataRow["Copy"].ToString(); var copy = Copy.Replace("'", "\"");
                    var Market = dataRow["Market"].ToString();
                    var PublishingHouse = dataRow["Publishing house"].ToString();
                    var Distributor = dataRow["Distributor"].ToString();
                    var PeriodicalType = dataRow["Periodical type"].ToString();
                    var SiteType = dataRow["Site type"].ToString();
                    var AdType = dataRow["Ad Type"].ToString();
                    var AdFormat = dataRow["Ad Format"].ToString();
                    var AdSize = dataRow["Ad Size"].ToString();
                    var AudioCodeOutdoor = dataRow["Audio code Outdoor"].ToString();
                    var AudioCodePress = dataRow["Audio code Press"].ToString();
                    var AudioCodeRadio = dataRow["Audio code Radio"].ToString();
                    var AudioCodeInternet = dataRow["Audio code Internet"].ToString();
                    var AdSectionType = dataRow["Ad Section Type"].ToString();
                    var AdSection = dataRow["Ad Section"].ToString();
                    var AdPosition = dataRow["Ad Position"].ToString();
                    var AdPage = dataRow["Ad Page"].ToString();
                    var IssueNo = dataRow["Issue No"].ToString();
                    var AdColor = dataRow["Ad Color"].ToString();
                    var Circulation = dataRow["Circulation"].ToString();
                    var DisplayPerc = dataRow[@"Display Perc"].ToString();
                    var Extension = dataRow["Extension"].ToString();
                    var AgencyInternet = dataRow["Agency Internet"].ToString();
                    var BuyerInternet = dataRow["Buyer Internet"].ToString();
                    var Damage = dataRow["Damage"].ToString();
                    var Direction = dataRow["Direction"].ToString();
                    var ProgrammeLocation = dataRow[@"Programme/Location"].ToString();
                    var progLocationTypologyVariables = dataRow[@"Prog/Location Typology\Variables"].ToString();
                    var Insertions = dataRow["Insertions"].ToString();
                    var Investment = dataRow["Investment"].ToString(); var investment = Investment.Replace(",", ".");

                    value += 1;

                    var cmdprocedure = new SqlCommand("UpbateFillterReference", con);
                    cmdprocedure.CommandType = CommandType.StoredProcedure;

                    cmdprocedure.Parameters.Add("@MediaType", SqlDbType.NVarChar).Value = MediaType;
                    cmdprocedure.Parameters.Add("@Sector", SqlDbType.NVarChar).Value = Sector;
                    cmdprocedure.Parameters.Add("@Category", SqlDbType.NVarChar).Value = Category;
                    cmdprocedure.Parameters.Add("@Class", SqlDbType.NVarChar).Value = Class;

                    //cmdprocedure.Parameters.Add("@Year", SqlDbType.Date).Value = YearDate.Date;
                    cmdprocedure.Parameters.Add("@Year", SqlDbType.NVarChar).Value = YearDateString;

                    cmdprocedure.Parameters.Add("@Month", SqlDbType.NVarChar).Value = Month;
                    cmdprocedure.Parameters.Add("@Week", SqlDbType.NVarChar).Value = Week;

                    //cmdprocedure.Parameters.Add("@Date", SqlDbType.Date).Value = DateDate.Date;
                    cmdprocedure.Parameters.Add("@Date", SqlDbType.NVarChar).Value = DateDateString;

                    //cmdprocedure.Parameters.Add("@StartTime", SqlDbType.Time).Value = StartTimeDate.TimeOfDay;
                    //cmdprocedure.Parameters.Add("@EndTime", SqlDbType.Time).Value = EndTimeDate.TimeOfDay;
                    cmdprocedure.Parameters.Add("@StartTime", SqlDbType.NVarChar).Value = StartTimeDateString;
                    cmdprocedure.Parameters.Add("@EndTime", SqlDbType.NVarChar).Value = EndTimeDateString;

                    cmdprocedure.Parameters.Add("@Producer", SqlDbType.NVarChar).Value = Producer;
                    cmdprocedure.Parameters.Add("@Brand", SqlDbType.NVarChar).Value = Brand;
                    cmdprocedure.Parameters.Add("@Product", SqlDbType.NVarChar).Value = Product;
                    cmdprocedure.Parameters.Add("@Copy", SqlDbType.NVarChar).Value = Copy;
                    cmdprocedure.Parameters.Add("@Market", SqlDbType.NVarChar).Value = Market;
                    cmdprocedure.Parameters.Add("@PublishingHouse", SqlDbType.NVarChar).Value = PublishingHouse;
                    cmdprocedure.Parameters.Add("@Distributor", SqlDbType.NVarChar).Value = Distributor;
                    cmdprocedure.Parameters.Add("@PeriodicalType", SqlDbType.NVarChar).Value = PeriodicalType;
                    cmdprocedure.Parameters.Add("@SiteType", SqlDbType.NVarChar).Value = SiteType;
                    cmdprocedure.Parameters.Add("@AdType", SqlDbType.NVarChar).Value = AdType;
                    cmdprocedure.Parameters.Add("@AdFormat", SqlDbType.NVarChar).Value = AdFormat;
                    cmdprocedure.Parameters.Add("@AdSize", SqlDbType.NVarChar).Value = AdSize;
                    cmdprocedure.Parameters.Add("@AudiocodeOutdoor", SqlDbType.NVarChar).Value = AudioCodeOutdoor;
                    cmdprocedure.Parameters.Add("@AudiocodePress", SqlDbType.NVarChar).Value = AudioCodePress;
                    cmdprocedure.Parameters.Add("@AudiocodeRadio", SqlDbType.NVarChar).Value = AudioCodeRadio;
                    cmdprocedure.Parameters.Add("@AudiocodeInternet", SqlDbType.NVarChar).Value = AudioCodeInternet;
                    cmdprocedure.Parameters.Add("@AdSectionType", SqlDbType.NVarChar).Value = AdSectionType;
                    cmdprocedure.Parameters.Add("@AdSection", SqlDbType.NVarChar).Value = AdSection;
                    cmdprocedure.Parameters.Add("@AdPosition", SqlDbType.NVarChar).Value = AdPosition;
                    cmdprocedure.Parameters.Add("@AdPage", SqlDbType.NVarChar).Value = AdPage;
                    cmdprocedure.Parameters.Add("@Issue", SqlDbType.NVarChar).Value = IssueNo;
                    cmdprocedure.Parameters.Add("@AdColor", SqlDbType.NVarChar).Value = AdColor;
                    cmdprocedure.Parameters.Add("@Circulation", SqlDbType.NVarChar).Value = Circulation;
                    cmdprocedure.Parameters.Add("@DisplayPerc", SqlDbType.NVarChar).Value = DisplayPerc;
                    cmdprocedure.Parameters.Add("@Extension", SqlDbType.NVarChar).Value = Extension;
                    cmdprocedure.Parameters.Add("@AgencyInternet", SqlDbType.NVarChar).Value = AgencyInternet;
                    cmdprocedure.Parameters.Add("@BuyerInternet", SqlDbType.NVarChar).Value = BuyerInternet;
                    cmdprocedure.Parameters.Add("@Damage", SqlDbType.NVarChar).Value = Damage;
                    cmdprocedure.Parameters.Add("@Direction", SqlDbType.NVarChar).Value = Direction;
                    cmdprocedure.Parameters.Add("@ProgLoc", SqlDbType.NVarChar).Value = ProgrammeLocation;
                    cmdprocedure.Parameters.Add("@ProgLocTV", SqlDbType.NVarChar).Value = progLocationTypologyVariables;
                    cmdprocedure.Parameters.Add("@Insertions", SqlDbType.NVarChar).Value = Insertions;
                    cmdprocedure.Parameters.Add("@Investment", SqlDbType.NVarChar).Value = Investment;

                cmdprocedure.ExecuteNonQuery();

                Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, value });
  
            }

            con.Close();
                
            MessageBox.Show("Загрузка выполнена!");

            ProgressBar1.Value = 0;
            LabelCurrentRow.Visibility = Visibility.Hidden;
            LableSlash.Visibility = Visibility.Hidden;
            LabelTotalRows.Visibility = Visibility.Hidden;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

    }
}