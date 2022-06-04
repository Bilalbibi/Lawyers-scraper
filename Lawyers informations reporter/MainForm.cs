using ExcelHelperExe;
using Lawyers_informations_reporter.Models;
using MetroFramework.Controls;
using MetroFramework.Forms;
using Newtonsoft.Json;
using scrapingTemplateV51.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace scrapingTemplateV51
{
    public partial class MainForm : MetroForm
    {
        public bool LogToUi = true;
        public bool LogToFile = true;

        private static string _path = Assembly.GetEntryAssembly().Location;
        private Dictionary<string, string> _config;
        public HttpCaller HttpCaller = new HttpCaller();
        List<LawyerInputData> _inputs;
        List<LawyerInputData> newInputs;
        List<Lawyer> _lawyersFromlsaMemberpro = new List<Lawyer>();
        NewSearch _newSearch;
        List<LawyerReportElements> _lawyersReportElements = new List<LawyerReportElements>();
        bool IsThefirstTime = true;

        public MainForm()
        {
            InitializeComponent();
            Reporter.OnLog += OnLog;
            Reporter.OnError += OnError;
            Reporter.OnProgress += OnProgress;
        }
        private void OnProgress(object sender, (int nbr, int total, string message) e)
        {
            Display($"{e.message} {e.nbr} / {e.total}");
            SetProgress(e.nbr * 100 / e.total);
        }

        private void OnError(object sender, string e)
        {
            ErrorLog(e);
        }

        private void OnLog(object sender, string e)
        {
            Display(e);
            NormalLog(e);
        }

        private async Task MainWork()
        {
            if (inputI.Text == "")
            {
                MessageBox.Show("Please select input excel file");
                return;
            }
            if (!File.Exists(inputI.Text))
            {
                MessageBox.Show("The input lawyer excel file not found, Please make sure that you select an exist input file");
                return;
            }
            _inputs = ExcelHelper.ReadFromExcel<LawyerInputData>(inputI.Text);
            //_newSearch = new NewSearch();
            //await _newSearch.ScrapeYears();
            if (!File.Exists("lawyers.txt"))
            {
                Reporter.Log("We are scraping the Law society Alberta lawyers's database");
                await ScrapeDataBase();
                Reporter.Log("The Law society Alberta lawyers's database scraped successfully");
                _lawyersFromlsaMemberpro = JsonConvert.DeserializeObject<List<Lawyer>>(File.ReadAllText("lawyers.txt"));
            }
            _lawyersFromlsaMemberpro = JsonConvert.DeserializeObject<List<Lawyer>>(File.ReadAllText("lawyers.txt"));

            var filledLawyerInputFiles = Directory.GetFiles("Inputs").ToList();
            if (filledLawyerInputFiles.Count > 0)
            {
                await SelectTheRquiredlawyers(filledLawyerInputFiles);
                var lastInputs = JsonConvert.DeserializeObject<List<List<Lawyer>>>(File.ReadAllText("Last inputs.text"));
                for (int i = 0; i < lastInputs.Count; i++)
                {

                    if (lastInputs[i].Count > 0)
                    {
                        var url = lastInputs[i][0].Url;
                        if (url.Contains("lsa.memberpro.net"))
                        {
                            var IsSelecyedlawSocityAlberta = lastInputs[i].Where(x => x.Url.Contains("lsa.memberpro.net")).Any(y => y.Url != null);
                            var Isrequired = lastInputs[i].Where(x => x.Url.Contains("lsa.memberpro.net")).Any(y => y.IsRequired != null);
                            if (IsSelecyedlawSocityAlberta && !Isrequired)
                            {
                                MessageBox.Show($"Please select the required lawyer from the suggestion(s) in the \"Law socity Alberta\" sheet for the this lawyer \"{lastInputs[i][0].LastName}\"");
                                return;
                            }
                        }
                        if (url.Contains("www.canlii.org"))
                        {
                            var IsOutecomesSelectedUrls = lastInputs[i].Where(x => x.Url.Contains("www.canlii.org")).Any(y => y.Url != null);
                            var Isrequired = lastInputs[i].Where(x => x.Url.Contains("www.canlii.org")).Any(y => y.IsRequired != null);
                            if (IsOutecomesSelectedUrls && !Isrequired)
                            {
                                MessageBox.Show($"Please select the required Outecome(s) from the suggestions in the \"Outcomes\" sheet for the this lawyer \"{lastInputs[i][0].LastName}\"");
                                return;
                            }
                        }
                        if (url.Contains("www.lawsociety.ab.ca"))
                        {
                            var IsHearungsSeletedUrls = lastInputs[i].Where(x => x.Url.Contains("www.lawsociety.ab.ca")).Any(y => y.Url != null);
                            var Isrequired = lastInputs[i].Where(x => x.Url.Contains("www.lawsociety.ab.ca")).Any(y => y.IsRequired != null);
                            if (IsHearungsSeletedUrls && !Isrequired)
                            {
                                MessageBox.Show($"Please select the required Hearing Schedule(s) from the suggestions in the \"Hearing Schedule\" sheet for the this lawyer \"{lastInputs[i][0].LastName}\"");
                                return;
                            }
                        }
                    }

                }
                await Task.Run(CreateReports);
                SaveSearch();
                DeletTheInputFiles(filledLawyerInputFiles);
                Reporter.Log("the report(s) ready and The last search saved successfilly");
            }
            else
            {
                if (File.Exists("Saved Searches.txt"))
                {
                    _lawyersReportElements = JsonConvert.DeserializeObject<List<LawyerReportElements>>(File.ReadAllText("Saved Searches.txt"));
                }
                if (_lawyersReportElements.Count > 0)
                {
                    newInputs = CheckForNewInputs();
                    if (newInputs.Count > 0)
                    {
                        Reporter.Log("The bot detected a new lawyer(s) that you want to create a report for, excel file(s) preparaion in progress...");
                        _newSearch = new NewSearch();
                        await _newSearch.PreparingTheRequiredFilesForSearch(newInputs, _lawyersFromlsaMemberpro);
                        MessageBox.Show("The excel file(s) for the new search is/are ready please chose the required lawyer(s) so the bot can create the report(s)");
                        return;
                    }
                    else
                    {
                        await Task.Run(CreateReports);
                        SaveSearch();
                        DeletTheInputFiles(filledLawyerInputFiles);
                        Reporter.Log("the report(s) ready and The last search saved successfilly");
                    }
                }
            }
            if (_lawyersReportElements.Count == 0)
            {
                Reporter.Log("Prepaing the ecxcel file(s) in progress to choose the required lawyers, please wait");
                _newSearch = new NewSearch();
                await _newSearch.PreparingTheRequiredFilesForSearch(_inputs, _lawyersFromlsaMemberpro);
                MessageBox.Show("The excel file(s) are ready please chose the required lawyer(s) so the bot can create the report(s)");
                return;
            }
        }

        private void DeletTheInputFiles(List<string> filledInputFiles)
        {
            foreach (var file in filledInputFiles)
            {
                var fullPathFile = new FileInfo(file).FullName;
                File.Delete(fullPathFile);
            }
        }

        private List<LawyerInputData> CheckForNewInputs()
        {
            var newInputs = new List<LawyerInputData>();
            var lastLawyersHistorySearch = _lawyersReportElements.Select(x => x.LastName).ToList();
            foreach (var input in _inputs)
            {
                if (!lastLawyersHistorySearch.Contains(input.LastName))
                {
                    newInputs.Add(input);
                }
            }
            return newInputs;
        }

        private void SaveSearch()
        {
            foreach (var lawyerReportElements in _lawyersReportElements)
            {
                lawyerReportElements.IsNew = false;
            }
            var json = JsonConvert.SerializeObject(_lawyersReportElements, Formatting.Indented);
            File.WriteAllText("Saved Searches.txt", json);
        }

        private async Task CreateReports()
        {
            var newLawyers = new List<LawyerReportElements>();
            foreach (var reportElements in _lawyersReportElements)
            {
                if (reportElements.IsNew)
                {
                    newLawyers.Add(reportElements);
                }
            }
            var createReports = new CreateReport();
            //createReports._years = _newSearch._years;
            if (newLawyers.Count > 0)
            {
                await createReports.CreatePdfs(newLawyers);
            }
            else
            {
                await createReports.CreatePdfs(_lawyersReportElements);
            }
        }

        async Task SelectTheRquiredlawyers(List<string> files)
        {
            var LawyersHistorySearch = _lawyersReportElements.Select(x => x.LastName).ToList();
            var allLists = new List<List<Lawyer>>();
            foreach (var file in files)
            {
                var lists = file.ReadAllSheetsFromExcel<Lawyer>();
                allLists.AddRange(lists);
                var lawyerReportElements = new LawyerReportElements();
                lawyerReportElements.LastName = lists[0][0].LastName;
                if (LawyersHistorySearch.Contains(lawyerReportElements.LastName))
                {
                    continue;
                }
                foreach (var lawyers in lists)
                {
                    if (lawyers.Count > 0)
                    {
                        var url = lawyers[0].Url;
                        if (url.Contains("lsa.memberpro.net"))
                        {
                            foreach (var lawyer in lawyers)
                            {
                                if (lawyer.IsRequired != null)
                                {
                                    lawyerReportElements.LawSocietyAlberta = lawyer;
                                    break;
                                }
                            }
                        }
                        if (url.Contains("www.canlii.org"))
                        {
                            foreach (var lawyer in lawyers)
                            {
                                if (lawyer.IsRequired != null)
                                {
                                    lawyerReportElements.OutComes.Add(lawyer);
                                }
                            }
                        }
                        if (url.Contains("www.lawsociety.ab.ca"))
                        {
                            foreach (var lawyer in lawyers)
                            {
                                if (lawyer.IsRequired != null)
                                {
                                    lawyerReportElements.Hearings.Add(lawyer);
                                }
                            }
                        }
                    }
                }
                _lawyersReportElements.Add(lawyerReportElements);
            }
            if (File.Exists("Saved Searches.txt"))
            {
                _lawyersReportElements.AddRange(JsonConvert.DeserializeObject<List<LawyerReportElements>>(File.ReadAllText("Saved Searches.txt")));
            }
            var json = JsonConvert.SerializeObject(allLists, Formatting.Indented);
            File.WriteAllText("Last inputs.text", json);
        }

        private async Task ScrapeDataBase()
        {
            var doc = await HttpCaller.GetDoc("http://lsa.memberpro.net/ssl/main/body.cfm?menu=directory&submenu=directoryPractisingMember&page_no=1&records_perpage_qy=2000000&person_nm=%25%25&first_nm=&city_nm=&location_nm=&gender_cl=&area_ds=&language_cl=&LSR_in=N&member_status_cl=&mode=search");
            var nodes = doc.DocumentNode.SelectNodes("//tr[contains(@class,'table-row')]");
            _lawyersFromlsaMemberpro = new List<Lawyer>();
            foreach (var node in nodes)
            {
                var name = node.SelectSingleNode("./td[1]").InnerText.Trim().Replace(", QC", "").Replace(" QC", "");
                var x = name.LastIndexOf(" ");
                var firstName = name.Substring(0, x);
                var LasttName = name.Substring(x + 1);
                var url = node.SelectSingleNode("./td[1]/a")?.GetAttributeValue("href", "");
                if (url == null)
                {
                    continue;
                }
                #region Attributes probably needed later
                //var city = node.SelectSingleNode("./td[2]").InnerText.Trim();
                //var gender = node.SelectSingleNode("./td[3]").InnerText.Trim();
                //var practisingStatus = node.SelectSingleNode("./td[4]").InnerText.Trim();
                //var enrolmentDate = node.SelectSingleNode("./td[5]").InnerText.Trim();
                //var firm = node.SelectSingleNode("./td[6]").InnerText.Trim(); 
                #endregion
                _lawyersFromlsaMemberpro.Add(new Lawyer
                {
                    FirstName = firstName,
                    LastName = LasttName,
                    Url = url,
                });
            }
            var json = JsonConvert.SerializeObject(_lawyersFromlsaMemberpro, Formatting.Indented);
            File.WriteAllText("lawyers.txt", json);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            if (!Directory.Exists("reports"))
            {
                Directory.CreateDirectory("reports");
            }
            ServicePointManager.DefaultConnectionLimit = 65000;
            Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            inputI.Text = _path + @"\input.txt";
            outputI.Text = _path + @"\output.csv";
            LoadConfig();
        }

        void InitControls(Control parent)
        {
            try
            {
                foreach (Control x in parent.Controls)
                {
                    try
                    {
                        if (x.Name.EndsWith("I"))
                        {
                            switch (x)
                            {
                                case MetroCheckBox _:
                                case CheckBox _:
                                    ((CheckBox)x).Checked = bool.Parse(_config[((CheckBox)x).Name]);
                                    break;
                                case RadioButton radioButton:
                                    radioButton.Checked = bool.Parse(_config[radioButton.Name]);
                                    break;
                                case TextBox _:
                                case RichTextBox _:
                                case MetroTextBox _:
                                    x.Text = _config[x.Name];
                                    break;
                                case NumericUpDown numericUpDown:
                                    numericUpDown.Value = int.Parse(_config[numericUpDown.Name]);
                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }

                    InitControls(x);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        public void SaveControls(Control parent)
        {
            try
            {
                foreach (Control x in parent.Controls)
                {
                    #region Add key value to disctionarry

                    if (x.Name.EndsWith("I"))
                    {
                        switch (x)
                        {
                            case MetroCheckBox _:
                            case RadioButton _:
                            case CheckBox _:
                                _config.Add(x.Name, ((CheckBox)x).Checked + "");
                                break;
                            case TextBox _:
                            case RichTextBox _:
                            case MetroTextBox _:
                                _config.Add(x.Name, x.Text);
                                break;
                            case NumericUpDown _:
                                _config.Add(x.Name, ((NumericUpDown)x).Value + "");
                                break;
                            default:
                                Console.WriteLine(@"could not find a type for " + x.Name);
                                break;
                        }
                    }
                    #endregion
                    SaveControls(x);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        private void SaveConfig()
        {
            _config = new Dictionary<string, string>();
            SaveControls(this);
            try
            {
                File.WriteAllText("config.txt", JsonConvert.SerializeObject(_config, Formatting.Indented));
            }
            catch (Exception e)
            {
                ErrorLog(e.ToString());
            }
        }
        private void LoadConfig()
        {
            try
            {
                _config = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText("config.txt"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }
            InitControls(this);
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString(), @"Unhandled Thread Exception");
        }
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show((e.ExceptionObject as Exception)?.ToString(), @"Unhandled UI Exception");
        }
        #region UIFunctions
        public delegate void WriteToLogD(string s, Color c);
        public void WriteToLog(string s, Color c)
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new WriteToLogD(WriteToLog), s, c);
                    return;
                }
                if (LogToUi)
                {
                    if (DebugT.Lines.Length > 5000)
                    {
                        DebugT.Text = "";
                    }
                    DebugT.SelectionStart = DebugT.Text.Length;
                    DebugT.SelectionColor = c;
                    DebugT.AppendText(DateTime.Now.ToString(Utility.SimpleDateFormat) + " : " + s + Environment.NewLine);
                }
                Console.WriteLine(DateTime.Now.ToString(Utility.SimpleDateFormat) + @" : " + s);
                if (LogToFile)
                {
                    File.AppendAllText(_path + "/data/log.txt", DateTime.Now.ToString(Utility.SimpleDateFormat) + @" : " + s + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        public void NormalLog(string s)
        {
            WriteToLog(s, Color.Black);
        }
        public void ErrorLog(string s)
        {
            WriteToLog(s, Color.Red);
        }
        public void SuccessLog(string s)
        {
            WriteToLog(s, Color.Green);
        }
        public void CommandLog(string s)
        {
            WriteToLog(s, Color.Blue);
        }

        public delegate void SetProgressD(int x);
        public void SetProgress(int x)
        {
            if (InvokeRequired)
            {
                Invoke(new SetProgressD(SetProgress), x);
                return;
            }
            if ((x <= 100))
            {
                ProgressB.Value = x;
            }
        }
        public delegate void DisplayD(string s);
        public void Display(string s)
        {
            if (InvokeRequired)
            {
                Invoke(new DisplayD(Display), s);
                return;
            }
            displayT.Text = s;
        }

        #endregion
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveConfig();
        }
        private void loadInputB_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog { Filter = @"XLSX|*.xlsx", InitialDirectory = _path };
            if (o.ShowDialog() == DialogResult.OK)
            {
                inputI.Text = o.FileName;
            }
        }
        private void openInputB_Click_1(object sender, EventArgs e)
        {
            try
            {
                Process.Start(inputI.Text);
            }
            catch (Exception ex)
            {
                ErrorLog(ex.ToString());
            }
        }
        private void openOutputB_Click_1(object sender, EventArgs e)
        {
            try
            {
                Process.Start(outputI.Text);
            }
            catch (Exception ex)
            {
                ErrorLog(ex.ToString());
            }
        }
        private void loadOutputB_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                Filter = @"csv file|*.csv",
                Title = @"Select the output location"
            };
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                outputI.Text = saveFileDialog1.FileName;
            }
        }

        private async void startB_Click_1(object sender, EventArgs e)
        {
            _lawyersReportElements = new List<LawyerReportElements>();
            await MainWork();
        }
    }
}
