using Common.Logging;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.Caching;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SocialMediaDataHarvester
{
    public enum ImportMode { SocialMediaLeads }
    internal static class DocumentMapHelper
    {
        internal static Dictionary<ImportMode, string> DocumentNameMap = new Dictionary<ImportMode, string>()
        {
            { ImportMode.SocialMediaLeads,"Social Media Leads" }
        };
    }

    public class ListData
    {
        public string Message { get; set; }
        public string DisplayMode { get; set; }
    }
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ILog log = LogManager.GetLogger<MainWindow>();
        private bool isloading;
        private string myFile = "";
        private MessageLine ActiveMessageLine;
        private MessageLine ChangeMessageLine;
        private AlterProgress ActiveProgress;
        private FileName ActiveFileName;
        Hashtable columns = new Hashtable();
        public System.Collections.ObjectModel.ObservableCollection<ListData> MyItems { get; set; }
        private ImportMode ActiveFormat;



        delegate void MessageLine(string line, string mode);
        delegate void FileName(string name);
        delegate void AlterProgress(int percent);
        private void ProgressChanged(int percent)
        {
            Progress.Value = percent;
        }

        private void AddMessageLine(string line, string mode)
        {
            log.Debug(m => m("Adding Text to Message Log {0}", line));
            ListData ld = new ListData();
            ld.DisplayMode = mode;
            ld.Message = line;
            MyItems.Insert(0, ld);
        }

        private void UpdateMessageLine(string line, string mode)
        {
            log.Debug(m => m("Updating Text in Message Log {0}", line));
            ListData ld = new ListData();
            ld.DisplayMode = mode;
            ld.Message = line;
            MyItems[0] = ld;
        }

        private void AddFileName(string name)
        {
            log.Debug(m => m("Setting File Name {0}", name));
            activeFile.Content = name;
        }

        public MainWindow()
        {
            MyItems = new System.Collections.ObjectModel.ObservableCollection<ListData>();
            InitializeComponent();
            VersionLabel.Text = String.Format("Version: {0}", System.Reflection.Assembly.GetEntryAssembly().GetName().Version);

            log.Info(m => m("Exam Data Import is Starting Up"));
            ActiveMessageLine = new MessageLine(AddMessageLine);
            ChangeMessageLine = new MessageLine(UpdateMessageLine);
            ActiveProgress = new AlterProgress(ProgressChanged);
            ActiveFileName = new FileName(AddFileName);
            this.messageList.DataContext = MyItems;
        }

        private void ProcessRecords_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!isloading)
                {
                    log.Debug(m => m("Beginning Import"));
                    ObjectCache cache = MemoryCache.Default;
                    List<ExcelRow> myRows = cache["excelimport"] as List<ExcelRow>;
                    if (myRows == null)
                        throw new Exception("You have not selected an Excel Sheet to Import");
                    log.Debug(m => m("Beginning Import"));
                    BackgroundWorker bw = new BackgroundWorker();

                    bw.WorkerReportsProgress = true;
                    bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
                    bw.DoWork += new DoWorkEventHandler(bw_DoImportWork);
                    bw.ReportProgress(-1);
                    bw.RunWorkerAsync();
                }
            }
            catch (Exception ex)
            {
                log.Error(m => m("Import Error: {0}", ex.Message));
                this.Dispatcher.BeginInvoke(ActiveMessageLine, new object[] { String.Format("Import Error: {0}", ex.Message), "Error" });
                //MessageBox.Show(ex.Message);
            }
        }

        private void dlg_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            log.Debug(m => m("Processing Selected Files"));
            BackgroundWorker bw = new BackgroundWorker();
            bw.WorkerReportsProgress = true;
            bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
            bw.DoWork += new DoWorkEventHandler(bw_LoadFilesDoWork);
            object[] o = new object[] { ((Microsoft.Win32.OpenFileDialog)sender).FileNames, ((Microsoft.Win32.OpenFileDialog)sender).SafeFileNames };
            bw.RunWorkerAsync(o);
        }

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.Dispatcher.BeginInvoke(ActiveProgress, new object[] { e.ProgressPercentage });
        }

        private void FileLoadButton_Click(object sender, RoutedEventArgs e)
        {
            if (!isloading)
            {
                log.Debug(m => m("Showing Open File Dialog"));
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.Filter = "Excel Files|*.xlsx";
                dlg.Title = "Select Excel Files";
                dlg.Multiselect = true;
                dlg.FileOk += new System.ComponentModel.CancelEventHandler(dlg_FileOk);
                dlg.ShowDialog();
            }
        }

        void bw_LoadFilesDoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bw = sender as BackgroundWorker;
            string[] files = (string[])((object[])e.Argument)[0];
            string[] names = (string[])((object[])e.Argument)[1];
            int count = 0;
            foreach (string s in files)
            {
                bw.ReportProgress(-1);
                if (s != string.Empty)
                {
                    string name = names[count];
                    this.Dispatcher.BeginInvoke(ActiveMessageLine, new object[] { String.Format("Loading File {0}", name), "Normal" });
                    this.Dispatcher.BeginInvoke(ActiveFileName, new object[] { name });
                    bw.ReportProgress(0);
                    LoadFile(s, name, bw);
                }
                count++;
            }
        }

        private void bw_DoImportWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bw = sender as BackgroundWorker;
            try
            {
                isloading = true;
                this.Dispatcher.BeginInvoke(ActiveMessageLine, new object[] { String.Format("Import Initiating"), "Success" });

                // do work based on Mode
                switch (ActiveFormat)
                {
                    case ImportMode.SocialMediaLeads:
                        SocialMediaLeads(bw);
                        break;
                }

                this.Dispatcher.BeginInvoke(ActiveMessageLine, new object[] { String.Format("Import Complete"), "Success" });
                isloading = false;
            }
            catch (Exception ex)
            {
                isloading = false;
                log.Error(m => m("Import Error: {0}", ex.Message));
                this.Dispatcher.BeginInvoke(ActiveMessageLine, new object[] { String.Format("Import Error: {0}", ex.Message), "Error" });
                //MessageBox.Show(ex.Message);
            }
        }

        private void SocialMediaLeads(BackgroundWorker bw)
        {
            ObjectCache cache = MemoryCache.Default;
            List<ExcelRow> myRows = cache["excelimport"] as List<ExcelRow>;
            ExcelPackage p = new ExcelPackage(new FileInfo(myFile));
            ExcelWorksheet mySheet = p.Workbook.Worksheets[1];

            string email;
            DateTime date;
            string type;

            int total = myRows.Count;
            int count = 1;
            var rrcString = ConfigurationManager.AppSettings["MESqlConnection"]; //@"server=www.rrc.co.uk;uid=SA;pwd=21brazil;database=RRCME";

            foreach (ExcelRow row in myRows)
            {
                int per = (int)Math.Round((double)(100 * count) / total);
                bw.ReportProgress(per);

                email = mySheet.Cells[row.Row, (int)columns["Email"]].Value.ToString().Trim();
                type = mySheet.Cells[row.Row, (int)columns["Type"]].Value.ToString().Trim();
                bool success = DateTime.TryParse(mySheet.Cells[row.Row, (int)columns["Date"]].Text, out date);

                if (success)
                {
                    //upload record to RRCME.SocialMedia
                    // write data into MoodleData
                    using (var lconn = new SqlConnection(rrcString))
                    {
                        if (lconn.State != ConnectionState.Open)
                        {
                            lconn.Open();
                        }
                        string sql = String.Format("Insert into SocialMedia (email,type,date) values ('{0}','{1}','{2}')", email, type, date.ToString("MMMM dd yyyy"));
                        var lcmd = new SqlCommand(sql, lconn);
                        lcmd.ExecuteNonQuery();

                        this.Dispatcher.BeginInvoke(ActiveMessageLine, new object[] { String.Format("Adding Record {0} {1} ", email, type), "Success" });
                    }

                }
                else
                {
                    this.Dispatcher.BeginInvoke(ActiveMessageLine, new object[] { String.Format("Could Not Read Date {0}: {1}", email, mySheet.Cells[row.Row, (int)columns["Date"]].Text), "Error" });
                }

                count++;
            }
        }

        private void LoadFile(string path, string name, BackgroundWorker bw)
        {
            myFile = path;
            FileInfo fi = new FileInfo(path);
            switch (fi.Extension)
            {
                case ".xlsx":
                    isloading = true;
                    columns = new Hashtable();
                    log.Debug(m => m("Processing File {0}", name));
                    this.Dispatcher.BeginInvoke(ActiveMessageLine, new object[] { String.Format("Processing File {0}", name), "Normal" });
                    try
                    {

                        using (ExcelPackage p = new ExcelPackage(fi))
                        {

                            try
                            {
                                ObjectCache cache = MemoryCache.Default;
                                CacheItemPolicy policy = new CacheItemPolicy();
                                policy.AbsoluteExpiration =
                                    DateTimeOffset.Now.AddSeconds(1000.0);

                                // wipe cache
                                cache.Remove("excelimport");

                                // get the first worksheet in the workbook
                                ExcelWorksheet worksheet = p.Workbook.Worksheets[1];

                                int sRow = -1;
                                var query = (from cell in worksheet.Cells["a:z"] where cell.Value != null && (cell.Value.ToString().Trim().ToLower() == "email") select cell);
                                foreach (var cell in query)
                                {
                                    sRow = Convert.ToInt32(Regex.Match(cell.Address, @"\d+").Value);
                                }

                                if (sRow == -1)
                                {
                                    throw new Exception("No Header Row Found");
                                }

                                //SaveReject(path, worksheet, sRow);

                                var query2 = (from cell in worksheet.Cells[String.Format("a{0}:az{0}", sRow)] select cell);
                                int colnum = 1;
                                foreach (var cell in query2)
                                {
                                    if (cell.Value != null)
                                    {
                                        switch (cell.Value.ToString().ToLower().Trim())
                                        {
                                            case "email":
                                                columns["Email"] = colnum;
                                                break;
                                            case "type":
                                                columns["Type"] = colnum;
                                                break;
                                            case "date":
                                                columns["Date"] = colnum;
                                                break;
                                        }
                                    }
                                    colnum++;
                                }

                                log.Debug(m => m("Creating Data Cache File"));
                                List<ExcelRow> myRows = cache["excelimport"] as List<ExcelRow>;
                                myRows = new List<ExcelRow>();

                                //detect last row
                                int colB = 2;
                                int colD = 4;
                                int lastRow = worksheet.Dimension.End.Row;
                                while (lastRow >= 1)
                                {
                                    var range = worksheet.Cells[lastRow, colB, lastRow, colD];
                                    if (range.Any(c => c.Value != null))
                                    {
                                        break;
                                    }
                                    lastRow--;
                                }


                                // read in data and import find header
                                for (int iRow = sRow + 1; iRow <= lastRow; iRow++)
                                {
                                    int per = (int)Math.Round((double)(100 * iRow) / lastRow);
                                    bw.ReportProgress(per);

                                    // skip hidden rows
                                    if (worksheet.Row(iRow).Hidden)
                                        continue;

                                    ActiveFormat = ImportMode.SocialMediaLeads;                       

                                    log.Debug(m => m("Detected Document as {0}", ActiveFormat));
                                    myRows.Add(worksheet.Row(iRow));

                                }

                                // display detected doc type
                                string type; DocumentMapHelper.DocumentNameMap.TryGetValue(ActiveFormat, out type);
                                this.Dispatcher.BeginInvoke(ActiveMessageLine, new object[] { String.Format("Detected Document Type: {0}", type), "Success" });

                                // cache data
                                log.Debug(m => m("Written Cache"));
                                cache.Set("excelimport", myRows, policy);
                                log.Debug(m => m("Excel File Read Complete - Found {0} Rows", myRows.Count));
                                this.Dispatcher.BeginInvoke(ActiveMessageLine, new object[] { String.Format("Read Complete - Found {0} Rows", myRows.Count), "Success" });

                            }
                            catch (Exception ex)
                            {
                                this.Dispatcher.BeginInvoke(ActiveMessageLine, new object[] { String.Format("Import Error: {0}", ex.Message), "Error" });
                                isloading = false;
                                break;
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        this.Dispatcher.BeginInvoke(ActiveMessageLine, new object[] { String.Format("Excel Error: {0}", ex.Message), "Error" });
                    }
                    isloading = false;
                    break;


            }
        }
    }
}
