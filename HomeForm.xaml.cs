using System;
using System.Collections.Generic;
using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using SCAdvert.Classes;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace SCAdvert
{
    /// <summary>
    /// Interaction logic for HomeForm.xaml
    /// </summary>
    public partial class HomeForm : System.Windows.Window
    {
        /// <summary>
        /// Command to be binded with the gridpaging command.
        /// </summary>
        private readonly RoutedUICommand changedIndex;

        public HomeForm()
        {
            InitializeComponent();
        }

        System.Data.DataTable dt;
        SqlDataAdapter da;

        //int minValue = 0;
        //int maxValue = 100;

        private System.Data.DataTable _dataSource;
        public System.Data.DataTable DataSource
        {
            get
            {
                return _dataSource;
            }
            set
            {
                _dataSource = value;
            }
        }

        private readonly string _connect = ConfigurationManager.ConnectionStrings["SqlConnect"].ConnectionString;

        private LoginForm _logForm;

        private void BtnEnter_Click(object sender, RoutedEventArgs e)
        {
            _logForm = new LoginForm();
            _logForm.Show();
        }

        private void FormHome_Loaded(object sender, RoutedEventArgs e)
        {
            btnPrevious.IsEnabled = false;
            btnFirst.IsEnabled = false;





            ComboBoxMediaType();
            ComboBoxYear();
            ComboBoxMonth();
            ComboBoxSector();
            ComboBoxCategory();
            ComboBoxClass();
            ComboBoxProducer();
            ComboBoxBrand();
            ComboBoxProduct();
            ComboBoxMarket();
            ComboBoxDistributor();
            ComboBoxAdType();
            ComboBoxAdFormat();

            DataGridSql.ItemsSource = DataBind(BindTable()).DefaultView;
        }

        #region " COMBOBOX "
        
        public void ComboBoxMediaType()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT DISTINCT MediaTypeID, MediaTypeName FROM dbo.MediaType"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboMediaType.SelectedValuePath = "MediaTypeID";
            CboMediaType.DisplayMemberPath = "MediaTypeName";
            CboMediaType.ItemsSource = dt.DefaultView;

            con.Close();

        }

        public void ComboBoxYear()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT DISTINCT Year FROM dbo.OrderDetail"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboYear.SelectedValuePath = "FillterReferenceID";
            CboYear.DisplayMemberPath = "Year";
            CboYear.ItemsSource = dt.DefaultView;

            con.Close();
        }

        public void ComboBoxMonth()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT DISTINCT Month FROM dbo.OrderDetail"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboMonth.SelectedValuePath = "FillterReferenceID";
            CboMonth.DisplayMemberPath = "Month";
            CboMonth.ItemsSource = dt.DefaultView;

            con.Close();
        }

        public void ComboBoxSector()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT DISTINCT SectorID, SectorName FROM dbo.Sector"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboSector.SelectedValuePath = "SectorID";
            CboSector.DisplayMemberPath = "SectorName";
            CboSector.ItemsSource = dt.DefaultView;

            con.Close();
        }

        public void ComboBoxCategory()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT DISTINCT CategoryID, CategoryName FROM dbo.Category"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboCategory.SelectedValuePath = "CategoryID";
            CboCategory.DisplayMemberPath = "CategoryName";
            CboCategory.ItemsSource = dt.DefaultView;

            con.Close();
        }

        public void ComboBoxClass()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT DISTINCT ClassID, ClassName FROM dbo.Class"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboClass.SelectedValuePath = "ClassID";
            CboClass.DisplayMemberPath = "ClassName";
            CboClass.ItemsSource = dt.DefaultView;

            con.Close();
        }

        public void ComboBoxProducer()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT  DISTINCT FillterReferenceID, Producer FROM dbo.OrderDetail"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboProducer.SelectedValuePath = "FillterReferenceID";
            CboProducer.DisplayMemberPath = "Producer";
            CboProducer.ItemsSource = dt.DefaultView;

            con.Close();
        }

        public void ComboBoxBrand()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT DISTINCT FillterReferenceID, Brand FROM dbo.OrderDetail"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboBrand.SelectedValuePath = "FillterReferenceID";
            CboBrand.DisplayMemberPath = "Brand";
            CboBrand.ItemsSource = dt.DefaultView;

            con.Close();
        }

        public void ComboBoxProduct()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT DISTINCT FillterReferenceID, Product FROM dbo.OrderDetail"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboProduct.SelectedValuePath = "FillterReferenceID";
            CboProduct.DisplayMemberPath = "Product";
            CboProduct.ItemsSource = dt.DefaultView;

            con.Close();
        }

        public void ComboBoxMarket()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT DISTINCT FillterReferenceID, Market FROM dbo.OrderDetail"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboMarket.SelectedValuePath = "FillterReferenceID";
            CboMarket.DisplayMemberPath = "Market";
            CboMarket.ItemsSource = dt.DefaultView;

            con.Close();
        }

        public void ComboBoxDistributor()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT DISTINCT FillterReferenceID, Distributor FROM dbo.OrderDetail"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboDistributor.SelectedValuePath = "FillterReferenceID";
            CboDistributor.DisplayMemberPath = "Distributor";
            CboDistributor.ItemsSource = dt.DefaultView;

            con.Close();
        }

        public void ComboBoxAdType()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT DISTINCT FillterReferenceID, AdType FROM dbo.OrderDetail"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboAdType.SelectedValuePath = "FillterReferenceID";
            CboAdType.DisplayMemberPath = "AdType";
            CboAdType.ItemsSource = dt.DefaultView;

            con.Close();
        }

        public void ComboBoxAdFormat()
        {
            dt = new System.Data.DataTable();

            const string queryString = @"SELECT DISTINCT FillterReferenceID, AdFormat FROM dbo.OrderDetail"; ;

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(queryString, con);
            con.Open();

            da.Fill(dt);

            CboAdFormat.SelectedValuePath = "FillterReferenceID";
            CboAdFormat.DisplayMemberPath = "AdFormat";
            CboAdFormat.ItemsSource = dt.DefaultView;

            con.Close();
        }

        public void BindComboBoxFilter()
        {
            try
            {
                if (DataTableFilter().DefaultView != null)
                {
                    DataGridSql.ItemsSource = DataBindFilter(DataTableFilter()).DefaultView;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // MediaTyp
        private void CboMediaType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        // Year
        private void CboYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        // Month
        private void CboMonth_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        // Sector
        private void CboSector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        // Category
        private void CboCategory_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        // Class
        private void CboClass_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        // Producer
        private void CboProducer_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        // Brand
        private void CboBrand_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        // Product
        private void CboProduct_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        // Market
        private void CboMarket_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        // Distributor
        private void CboDistributor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        // AdType
        private void CboAdType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        // AdFormat
        private void CboAdFormat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindComboBoxFilter();
        }

        private System.Data.DataTable DataTableFilter()
        {
            var myDataView = new DataView(BindTable().DefaultView.Table);


            if (CboMediaType.SelectedValue != null)
            {
                var sItemMediaType = ((DataRowView)CboMediaType.SelectedItem).Row[1].ToString();

                if (sItemMediaType != "")
                {
                    myDataView.RowFilter = "[MediaTypeName] IN (" + "'" + sItemMediaType + "'" + ")";
                };
            }

            if (CboYear.SelectedValue != null)
            {
                var sItem = ((DataRowView)CboYear.SelectedItem).Row["Year"].ToString();
                var yearDate = DateTime.ParseExact(sItem, "dd.MM.yyyy H:mm:ss", CultureInfo.InvariantCulture);
                var yearDateString = yearDate.ToString("yyyy");
                if (yearDateString != "")
                {
                    myDataView.RowFilter = "[Column1] IN (" + "'" + yearDateString + "'" + ")";
                }
            }

            if (CboMonth.SelectedValue != null)
            {
                var sItemMonth = ((DataRowView)CboMonth.SelectedItem).Row[1].ToString();
                if (sItemMonth != "")
                {
                    myDataView.RowFilter = "[Month] IN (" + "'" + sItemMonth + "'" + ")";
                }
            }

            if (CboSector.SelectedValue != null)
            {
                var sItemSector = ((DataRowView)CboSector.SelectedItem).Row[1].ToString();
                if (sItemSector != "")
                {
                    myDataView.RowFilter = "[SectorName] IN (" + "'" + sItemSector + "'" + ")";
                }
            }

            if (CboCategory.SelectedValue != null)
            {
                var sItemCategory = ((DataRowView)CboCategory.SelectedItem).Row[1].ToString();
                if (sItemCategory != "")
                {
                    myDataView.RowFilter = "[CategoryName] IN (" + "'" + sItemCategory + "'" + ")";
                }
            }

            if (CboClass.SelectedValue != null)
            {
                var sItemClass = ((DataRowView)CboClass.SelectedItem).Row[1].ToString();
                if (sItemClass != "")
                {
                    myDataView.RowFilter = "[ClassName] IN (" + "'" + sItemClass + "'" + ")";
                }
            }

            if (CboProducer.SelectedValue != null)
            {
                var sItemProducer = ((DataRowView)CboProducer.SelectedItem).Row["Producer"].ToString();
                if (sItemProducer != "")
                {
                    myDataView.RowFilter = "[Producer] IN (" + "'" + sItemProducer + "'" + ")";
                }
            }

            if (CboBrand.SelectedValue != null)
            {
                var sItemBrand = ((DataRowView)CboBrand.SelectedItem).Row["Brand"].ToString();
                if (sItemBrand != "")
                {
                    myDataView.RowFilter = "[Brand] IN (" + "'" + sItemBrand + "'" + ")";
                }
            }

            if (CboProduct.SelectedValue != null)
            {
                var sItemProduct = ((DataRowView)CboProduct.SelectedItem).Row["Product"].ToString();
                if (sItemProduct != "")
                {
                    myDataView.RowFilter = "[Product] IN (" + "'" + sItemProduct + "'" + ")";
                }
            }

            if (CboMarket.SelectedValue != null)
            {
                var sItemMarket = ((DataRowView)CboMarket.SelectedItem).Row["Market"].ToString();
                if (sItemMarket != "")
                {
                    myDataView.RowFilter = "[Market] IN (" + "'" + sItemMarket + "'" + ")";
                }
            }

            if (CboDistributor.SelectedValue != null)
            {
                var sItemDistributor = ((DataRowView)CboDistributor.SelectedItem).Row[1].ToString();
                if (sItemDistributor != "")
                {
                    myDataView.RowFilter = "[Distributor] IN (" + "'" + sItemDistributor + "'" + ")";
                }
            }

            if (CboAdType.SelectedValue != null)
            {
                var sItemAdType = ((DataRowView)CboAdType.SelectedItem).Row["AdType"].ToString();
                if (sItemAdType != "")
                {
                    myDataView.RowFilter = "[AdType] IN (" + "'" + sItemAdType + "'" + ")";
                }
            }

            if (CboAdFormat.SelectedValue != null)
            {

                var sItemAdFormat = ((DataRowView)CboAdFormat.SelectedItem).Row["AdFormat"].ToString();
                if (sItemAdFormat != "")
                {
                    myDataView.RowFilter = "[AdFormat] IN (" + "'" + sItemAdFormat + "'" + ")";
                }
            }

            var dtView = myDataView.ToTable();

            return dtView;

        }

        #endregion

        public class ColumnDataTable
        {
            public string mediaType { get; set; }
            public string year { get; set; }
            public string month { get; set; }
            public string SectorName { get; set; }
        }
        public List<ColumnDataTable> GetDataTableList()
        {
            var nameColumn = new List<ColumnDataTable>();
            try
            {
                System.Data.DataTable dt = BindTable();

                foreach (DataRow datarow in dt.Rows)
                {
                    var ColumnValue = new ColumnDataTable
                    {
                        mediaType = datarow["MediaTypeName"].ToString(),
                        year = datarow["Column1"].ToString(),
                        month = datarow["Month"].ToString(),
                        SectorName = datarow["SectorName"].ToString()
                    };
                    nameColumn.Add(ColumnValue);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return nameColumn;
        }

        private delegate void UpdateProgressBarDelegate(DependencyProperty dp, Object value);
        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {

            var dialog = new SaveFileDialog { FileName = "Excel_Data.xls", Filter = "Excel files|*.xls" };
            bool? showDialog = dialog.ShowDialog();

            if (showDialog == true)
            {
                var misValue = System.Reflection.Missing.Value;
                var xlApp = new Microsoft.Office.Interop.Excel.Application();
                var xlWorkBook = (Workbook)xlApp.Workbooks.Add(misValue);
                var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.Item[1];

                var pbExcel = new pbDataTableToExcel();
                pbExcel.Show();
                //Configure the ProgressBar
                pbExcel.PbDtToExcel.Minimum = 0;

                var totalNumbersRow = DataTableFilter().Rows.Count;
                //int totalNumbersRow = DataGridSql.Items.Count;

                pbExcel.PbDtToExcel.Maximum = totalNumbersRow;

                //MessageBox.Show(totalnumbersrow.ToString());

                pbExcel.PbDtToExcel.Value = 0;

                //Stores the value of the ProgressBar
                double value = 0;

                // to the ProgressBar's SetValue method.
                UpdateProgressBarDelegate updatePbDelegate = pbExcel.PbDtToExcel.SetValue;

                //называем колонки
                xlWorkSheet.Cells[1, 1] = "MediaType";
                xlWorkSheet.Cells[1, 2] = "Year";
                xlWorkSheet.Cells[1, 3] = "Month";
                xlWorkSheet.Cells[1, 4] = "SectorName";
                xlWorkSheet.Cells[1, 5] = "CategoryName";
                xlWorkSheet.Cells[1, 6] = "ClassName";
                xlWorkSheet.Cells[1, 7] = "Producer";
                xlWorkSheet.Cells[1, 8] = "Brand";
                xlWorkSheet.Cells[1, 9] = "Product";
                xlWorkSheet.Cells[1, 10] = "Market";
                xlWorkSheet.Cells[1, 11] = "Distributor";
                xlWorkSheet.Cells[1, 12] = "AdType";
                xlWorkSheet.Cells[1, 13] = "AdFormat";

                #region " Rows from SQL "
                //заполняем строки

                var column1 = 0;
                var column2 = 0;
                var column3 = 0;
                var column4 = 0;
                var column5 = 0;
                var column6 = 0;
                var column7 = 0;
                var column8 = 0;
                var column9 = 0;
                var column10 = 0;
                var column11 = 0;
                var column12 = 0;
                var column13 = 0;

                foreach (DataRow r in DataTableFilter().Rows)
                {
                    var columnMediaType = r["MediaTypeName"].ToString();
 
                    xlWorkSheet.Cells[column1++ + 2, 1] = columnMediaType;
                    value += 1;

                    var columnYear = r["Column1"].ToString();
                    xlWorkSheet.Cells[column2++ + 2, 2] = columnYear;

                    var columnMonth = r["Month"].ToString();
                    xlWorkSheet.Cells[column3++ + 2, 3] = columnMonth;

                    var columnSectorName = r["SectorName"].ToString();
                    xlWorkSheet.Cells[column4++ + 2, 4] = columnSectorName;

                    var columnCategoryName = r["CategoryName"].ToString();
                    xlWorkSheet.Cells[column5++ + 2, 5] = columnCategoryName;

                    var columnClassName = r["ClassName"].ToString();
                    xlWorkSheet.Cells[column6++ + 2, 6] = columnClassName;

                    var columnProducer = r["Producer"].ToString();
                    xlWorkSheet.Cells[column7++ + 2, 7] = columnProducer;

                    var columnBrand = r["ClassName"].ToString();
                    xlWorkSheet.Cells[column8++ + 2, 8] = columnBrand;

                    var columnProduct = r["Product"].ToString();
                    xlWorkSheet.Cells[column9++ + 2, 9] = columnProduct;

                    var columnMarket = r["Market"].ToString();
                    xlWorkSheet.Cells[column10++ + 2, 10] = columnMarket;

                    var columnDistributor = r["Distributor"].ToString();
                    xlWorkSheet.Cells[column11++ + 2, 11] = columnDistributor;

                    var columnAdType = r["AdType"].ToString();
                    xlWorkSheet.Cells[column12++ + 2, 12] = columnAdType;

                    var columnAdFormat = r["AdFormat"].ToString();
                    xlWorkSheet.Cells[column13++ + 2, 13] = columnAdFormat;

                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, value });
                }
                #endregion

                #region " Rows from DataGrid "
                ////заполняем строки
                //for (var rowInd = 0; rowInd < DataGridSql.Items.Count; rowInd++)
                //{
                //    var columnMediaType = (DataGridSql.Items[rowInd] as DataRowView).Row["MediaTypeName"].ToString();

                //    xlWorkSheet.Cells[rowInd + 2, 1] = columnMediaType;

                //    value += 1;

                //    var columnYear = ((DataRowView) DataGridSql.Items[rowInd]).Row["Column1"].ToString();
                //    xlWorkSheet.Cells[rowInd + 2, 2] = columnYear;

                //    var columnMonth = ((DataRowView) DataGridSql.Items[rowInd]).Row["Month"].ToString();
                //    xlWorkSheet.Cells[rowInd + 2, 3] = columnMonth;

                //    var columnSectorName = ((DataRowView) DataGridSql.Items[rowInd]).Row["SectorName"].ToString();
                //    xlWorkSheet.Cells[rowInd + 2, 4] = columnSectorName;

                //    var columnCategoryName = ((DataRowView) DataGridSql.Items[rowInd]).Row["CategoryName"].ToString();
                //    xlWorkSheet.Cells[rowInd + 2, 5] = columnCategoryName;

                //    var columnClassName = ((DataRowView) DataGridSql.Items[rowInd]).Row["ClassName"].ToString();
                //    xlWorkSheet.Cells[rowInd + 2, 6] = columnClassName;

                //    var columnProducer = ((DataRowView) DataGridSql.Items[rowInd]).Row["Producer"].ToString();
                //    xlWorkSheet.Cells[rowInd + 2, 7] = columnProducer;

                //    var columnBrand = ((DataRowView) DataGridSql.Items[rowInd]).Row["ClassName"].ToString();
                //    xlWorkSheet.Cells[rowInd + 2, 8] = columnBrand;

                //    var columnProduct = ((DataRowView)DataGridSql.Items[rowInd]).Row["Product"].ToString();
                //    xlWorkSheet.Cells[rowInd + 2, 9] = columnProduct;

                //    var columnMarket = ((DataRowView)DataGridSql.Items[rowInd]).Row["Market"].ToString();
                //    xlWorkSheet.Cells[rowInd + 2, 10] = columnMarket;

                //    var columnDistributor = ((DataRowView)DataGridSql.Items[rowInd]).Row["Distributor"].ToString();
                //    xlWorkSheet.Cells[rowInd + 2, 11] = columnDistributor;

                //    var columnAdType = ((DataRowView)DataGridSql.Items[rowInd]).Row["AdType"].ToString();
                //    xlWorkSheet.Cells[rowInd + 2, 12] = columnAdType;

                //    var columnAdFormat = ((DataRowView)DataGridSql.Items[rowInd]).Row["AdFormat"].ToString();
                //    xlWorkSheet.Cells[rowInd + 2, 13] = columnAdFormat;

                //    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, value });
                //}
                #endregion

                //выбираем всю область данных
                Microsoft.Office.Interop.Excel.Range xlSheetRange = xlWorkSheet.UsedRange;

                //выравниваем строки и колонки по их содержимому
                xlSheetRange.Columns.AutoFit();
                xlSheetRange.Rows.AutoFit();

                xlWorkBook.SaveAs(dialog.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                misValue, misValue, misValue, misValue,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue,
                misValue, misValue, misValue, misValue);

                xlWorkBook.Close(true, misValue, misValue);

                xlApp.Quit();
                pbExcel.Close();
            }
            else
            {
                return;
            }

            MessageBox.Show("Файл сохранен");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private System.Data.DataTable BindTable()
        {

            dt = new System.Data.DataTable();

            const string querystring = @"SELECT  YEAR(OrderDetail.[Year]), Category.CategoryName, 
                    Class.ClassName, MediaType.MediaTypeName, 
                    Sector.SectorName, OrderDetail.Date, OrderDetail.StartWeek, OrderDetail.EndWeek, 
                    OrderDetail.StartTime, OrderDetail.EndTime, 
                    (DATEDIFF (second, OrderDetail.[StartTime], OrderDetail.[EndTime])) AS 'Duration',  
                    OrderDetail.Producer, OrderDetail.Brand, OrderDetail.Product, OrderDetail.Copy, OrderDetail.Market, 
                    OrderDetail.PublishingHouse, OrderDetail.Distributor, OrderDetail.PeriodicalType, OrderDetail.SiteType, 
                    OrderDetail.AdType, OrderDetail.AdFormat, 
                    OrderDetail.AdSize, OrderDetail.AudiocodeOutdoor, OrderDetail.AudiocodePress, OrderDetail.AudiocodeRadio, 
                    OrderDetail.AudiocodeInternet, 
                    OrderDetail.AdSectionType, OrderDetail.AdSection, OrderDetail.AdPosition, 
                    OrderDetail.AdPage, OrderDetail.Issue, OrderDetail.Circulation, OrderDetail.AdColor, 
                    OrderDetail.DisplayPerc, OrderDetail.Extension, OrderDetail.AgencyInternet, OrderDetail.BuyerInternet, 
                    OrderDetail.Damage, OrderDetail.Direction, 
                    OrderDetail.ProgLoc, OrderDetail.ProgLocTV, OrderDetail.Insertions,  OrderDetail.Investment, OrderDetail.Month, 
                    OrderDetail.Week
                    FROM Category INNER JOIN
                    FillterReference ON Category.CategoryID = FillterReference.CategoryID INNER JOIN
                    Class ON FillterReference.ClassID = Class.ClassID INNER JOIN
                    MediaType ON FillterReference.MediaTypeID = MediaType.MediaTypeID INNER JOIN
                    Sector ON FillterReference.SectorID = Sector.SectorID INNER JOIN
                    OrderDetail ON FillterReference.ReferenceID = OrderDetail.FillterReferenceID";

            var con = new SqlConnection(_connect);

            da = new SqlDataAdapter(querystring, con);
            con.Open();

            da.Fill(dt);
            
            return dt;
        }

        private void MyGrid_LoadingRow( object sender, DataGridRowEventArgs e )
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        #region " paging DataGrid "

        private int _pageSize = 200;
        public int PageSize
        {
            get
            {
                return _pageSize;
            }
            set
            {
                _pageSize = value;
            }
        }

        private System.Data.DataTable ShowDataFilter(int pageNumber)
        {
            var dt = new System.Data.DataTable();
            int startIndex = PageSize * (pageNumber - 1);

            var result = DataTableFilter().AsEnumerable().Where((s, k) => (k >= startIndex && k < (startIndex + PageSize)));

            foreach (DataColumn colunm in DataTableFilter().Columns)
            {
                dt.Columns.Add(colunm.ColumnName);
            }

            foreach (var item in result)
            {
                dt.ImportRow(item);
            }

            txtPaging.Text = string.Format("Страница {0} из {1}", pageNumber, (DataTableFilter().Rows.Count / PageSize) + 1);

            return dt;
        }

        private System.Data.DataTable ShowData(int pageNumber)
        {
            var dt = new System.Data.DataTable();
            int startIndex = PageSize * (pageNumber - 1);

            var result = BindTable().AsEnumerable().Where((s, k) => (k >= startIndex && k < (startIndex + PageSize)));

            foreach (DataColumn colunm in BindTable().Columns)
            {
                dt.Columns.Add(colunm.ColumnName);
            }

            foreach (var item in result)
            {
                dt.ImportRow(item);
            }

            txtPaging.Text = string.Format("Страница {0} из {1}", pageNumber, (BindTable().Rows.Count / PageSize) + 1);

            return dt;
        }

        public System.Data.DataTable DataBindFilter(System.Data.DataTable dataTable)
        {
            dt = new System.Data.DataTable();

            DataSource = dataTable;

            dt = ShowDataFilter(1);

            return dt;
        } 

        public System.Data.DataTable DataBind(System.Data.DataTable dataTable)
        {
            dt = new System.Data.DataTable();

            DataSource = dataTable;

            dt = ShowData(1);
       
            return dt;
        } 

        private int _width;
        public int ControlWidth
        {
            get
            {
                if (_width == 0)
                    return Convert.ToInt32(DataGridSql.Width);
                else
                    return _width;
            }
            set
            {
                _width = value;
                DataGridSql.Width = _width;
            }
        }

        private int _height;
        public int ControlHeight
        {
            get
            {
                if (_height == 0)
                    return Convert.ToInt32(DataGridSql.Height);
                else
                    return _height;
            }
            set
            {
                _height = value;
                DataGridSql.Height = _height;
            }
        }

        private string _firstButtonText = string.Empty;
        public string FirstButtonText
        {
            get
            {
                if (_firstButtonText == string.Empty)
                    return Convert.ToString(btnFirst.Content);
                else
                    return _firstButtonText;
            }
            set
            {
                _firstButtonText = value;
                btnFirst.Content = _firstButtonText;
            }
        }

        private string _lastButtonText = string.Empty;
        public string LastButtonText
        {
            get
            {
                if (_lastButtonText == string.Empty)
                    return Convert.ToString(btnLast.Content);
                else
                    return _lastButtonText;
            }
            set
            {
                _lastButtonText = value;
                btnLast.Content = _lastButtonText;
            }
        }

        private string _previousButtonText = string.Empty;
        public string PreviousButtonText
        {
            get
            {
                if (_previousButtonText == string.Empty)
                    return Convert.ToString(btnPrevious.Content);
                else
                    return _previousButtonText;
            }
            set
            {
                _previousButtonText = value;
                btnPrevious.Content = _previousButtonText;
            }
        }

        private string _nextButtonText = string.Empty;
        public string NextButtonText
        {
            get
            {
                if (_nextButtonText == string.Empty)
                    return Convert.ToString(btnNext.Content);
                else
                    return _nextButtonText;
            }
            set
            {
                _nextButtonText = value;
                btnNext.Content = _nextButtonText;
            }
        }

        #endregion

        private int _currentPage;

        private void btnFirst_Click(object sender, RoutedEventArgs e)
        {
            btnPrevious.IsEnabled = false;
            btnFirst.IsEnabled = false;
            btnNext.IsEnabled = true;
            btnLast.IsEnabled = true;

            if (_currentPage == 1)
            {
            }
            else
            {
                _currentPage = 1;
                //DataGridSql.ItemsSource = ShowData(_currentPage).DefaultView;
                DataGridSql.ItemsSource = ShowDataFilter(_currentPage).DefaultView;
            }
        }

        private void btnPrevious_Click(object sender, RoutedEventArgs e)
        {
            btnNext.IsEnabled = true;
            btnLast.IsEnabled = true;

            if (_currentPage == 1)
            {
                btnPrevious.IsEnabled = false;
                btnFirst.IsEnabled = false;
            }
            else
            {
                _currentPage -= 1;
                //DataGridSql.ItemsSource = ShowData(_currentPage).DefaultView;
                DataGridSql.ItemsSource = ShowDataFilter(_currentPage).DefaultView;
            }
        }

        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            btnPrevious.IsEnabled = true;
            btnFirst.IsEnabled = true;

            var lastPage = (DataSource.Rows.Count / PageSize) + 1;

            if (_currentPage == lastPage)
            {
                btnNext.IsEnabled = false;
                btnLast.IsEnabled = false;
            }
            else
            {
                _currentPage += 1;

                //DataGridSql.ItemsSource = ShowData(_currentPage).DefaultView;

                //if (CboCategory.Text != "")
                //{
                    DataGridSql.ItemsSource = ShowDataFilter(_currentPage).DefaultView;
                //}

             
            }      
        }

        private void btnLast_Click(object sender, RoutedEventArgs e)
        {
            int previousPage = _currentPage;
            _currentPage = (DataSource.Rows.Count / PageSize) + 1;

            btnNext.IsEnabled = false;
            btnLast.IsEnabled = false;
            btnPrevious.IsEnabled = true;
            btnFirst.IsEnabled = true;

            if (previousPage == _currentPage)
            {
            }
            else
            {
                //DataGridSql.ItemsSource = ShowData(_currentPage).DefaultView;
                DataGridSql.ItemsSource = ShowDataFilter(_currentPage).DefaultView;
            }
        }

        private void btnClearFilter_Click(object sender, RoutedEventArgs e)
        {
            DataGridSql.ItemsSource = DataBind(BindTable()).DefaultView;
        }
    }
}