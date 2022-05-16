using ExcelHelperExe;
using Lawyers_informations_reporter.Models;
using MetroFramework.Controls;
using MetroFramework.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using scrapingTemplateV51.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
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
        List<string> _years = new List<string>();
        List<Lawyer> _lawyersFromlsaMemberpro = new List<Lawyer>();

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
            if (!File.Exists("lawyers.txt"))
            {
                Reporter.Log("We lost the data base please wait till we rescrape it again");
                await ScrapeDataBase();
                Reporter.Log("the scrape success, you can select the required lawyer now");
                _lawyersFromlsaMemberpro = JsonConvert.DeserializeObject<List<Lawyer>>(File.ReadAllText("lawyers.txt"));
            }
            else
            {
                _lawyersFromlsaMemberpro = JsonConvert.DeserializeObject<List<Lawyer>>(File.ReadAllText("lawyers.txt"));
            }
            if (!File.Exists(inputI.Text) || inputI.Text != "")
            {
                _inputs = ExcelHelper.ReadFromExcel<LawyerInputData>(inputI.Text);
            }
            else
            {
                MessageBox.Show("Please select the excel input data");
                return;
            }
            var filledInputFiles = Directory.GetFiles("Inputs").ToList();
            if (filledInputFiles.Count > 0)
            {
                await CreateReports(filledInputFiles);
            }
            await ScrapeYears();
            await StartScrapingThePossiblitiesInputs();
            MessageBox.Show("The excel files are ready please chose the required lawyer(s) so the bot can create the report(s)");
        }

        async Task CreateReports(List<string> filledInputFiles)
        {
            var lawyerReportElements = new List<LawyerReportElements>();
            foreach (var filledInputFile in filledInputFiles)
            {

            }
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

                #region other attribute maybe needed in the future
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

        private async Task ScrapeYears()
        {
            var doc = await HttpCaller.GetDoc("https://www.canlii.org/en/ab/abls/");
            var nodes = doc.DocumentNode.SelectNodes("//select[@id='navYearsSelector']/option");
            for (int i = 1; i < nodes.Count; i++)
            {
                _years.Add(nodes[i].InnerText.Trim());
            }
        }

        private async Task StartScrapingThePossiblitiesInputs()
        {
            foreach (var input in _inputs)
            {
                await CreateTheInputsBasedOnthelastName(input);
            }
        }

        private async Task CombinePdfs(LawyerInputData input)
        {
            var pdfFiles = Directory.GetFiles("pdfs");
            PdfDocument outputPDFDocument = new PdfDocument();
            if (pdfFiles.Count() > 0)
            {
                foreach (string pdfFile in pdfFiles)
                {
                    PdfDocument inputPDFDocument = PdfReader.Open(pdfFile, PdfDocumentOpenMode.Import);
                    outputPDFDocument.Version = inputPDFDocument.Version;
                    foreach (PdfPage page in inputPDFDocument.Pages)
                    {
                        outputPDFDocument.AddPage(page);
                    }
                }
                outputPDFDocument.Save($"reports/{input.FirstName} {input.LastName}.pfd");
                foreach (var pdfFile in pdfFiles)
                {
                    File.Delete(pdfFile);
                }
            }
            if (pdfFiles.Count() == 0)
            {
                File.AppendAllText($"lawyers not founded/lawyers not founded.txt", input.Firm + " " + input.LastName);
            }
        }

        private async Task CreateTheInputsBasedOnthelastName(LawyerInputData input)
        {
            await GetDataFromLawSocityAlberta(input);
            await GetDecisionsAndOutcomes(input);
            await GetHearingSchedule(input);
        }

        private async Task GetHearingSchedule(LawyerInputData input)
        {
            var doc = await HttpCaller.GetDoc("https://www.lawsociety.ab.ca/regulation/adjudication/hearing-schedule/");
            var nodes = doc.DocumentNode.SelectNodes("//table[@id='hearings']//tr/td[1]/a");
            var hearingCases = new List<Lawyer>();
            foreach (var node in nodes)
            {
                var fullName = node.InnerText.Trim();
                var x = fullName.LastIndexOf(" ");
                var firstName = fullName.Substring(0, x);
                var lastName = fullName.Substring(x + 1);
                var url = node.GetAttributeValue("href", "").Trim();
                hearingCases.Add(
                    new Lawyer
                    {
                        FirstName = firstName,
                        LastName = lastName,
                        Url = url
                    });
            }
            var suggestionsHearing = hearingCases.FindAll(x => x.LastName.ToLower() == input.LastName.ToLower());
            suggestionsHearing.Save($"Inputs/{input.FirstName} {input.LastName} suggestions.xlsx", "Hearing Schedule");
        }

        private async Task GetDecisionsAndOutcomes(LawyerInputData input)
        {
            #region Outcomes from lawsocity
            //var doc = await HttpCaller.GetDoc("https://www.lawsociety.ab.ca/regulation/adjudication/decisions-outcomes/");
            //var records = doc.DocumentNode.SelectNodes("//table[@id='outcomes']//tr");
            //for (int i = 1; i < records.Count; i++)
            //{
            //    var nameAndLastName = records[i].SelectSingleNode(".//td[2]").InnerText.Trim();
            //    var x = nameAndLastName.LastIndexOf(" ");
            //    var name = nameAndLastName.Substring(0, x);
            //    var lasName = nameAndLastName.Substring(x + 1);
            //} 
            #endregion
            var outcomesHistory = new List<Lawyer>();
            for (int i = 0; i < _years.Count; i++)
            {
                var json = await HttpCaller.GetHtml($"https://www.canlii.org/en/ab/abls/nav/date/{_years[i]}/items");
                var array = JArray.Parse(json);
                foreach (var node in array)
                {
                    var lastName = (string)node.SelectToken("..styleOfCause");
                    var x = lastName.LastIndexOf(" ");
                    lastName = lastName.Substring(x + 1);
                    var url = "https://www.canlii.org" + ((string)node.SelectToken("..url"));
                    var date = ((string)node.SelectToken("..judgmentDate"));
                    outcomesHistory.Add(new Lawyer { LastName = lastName, Url = url });
                }
            }
            var lawyerOutcomesHistories = outcomesHistory.FindAll(x => x.LastName.ToLower() == input.LastName.ToLower());
            lawyerOutcomesHistories.Save($"Inputs/{input.FirstName} {input.LastName} suggestions.xlsx", "Outcomes");
        }

        private async Task GetDataFromLawSocityAlberta(LawyerInputData input)
        {
            #region preceding work about socity alberta
            //var format = new List<KeyValuePair<string, string>>()
            //{
            //   new KeyValuePair<string, string>("person_nm",input.LastName),
            //   new KeyValuePair<string, string>("first_nm",input.FirstName),
            //   new KeyValuePair<string, string>("member_status_cl","PRAC"),
            //   new KeyValuePair<string, string>("city_nm",""),
            //   new KeyValuePair<string, string>("location_nm",input.Firm),
            //   new KeyValuePair<string, string>("gender_cl",""),
            //   new KeyValuePair<string, string>("language_cl",""),
            //   new KeyValuePair<string, string>("area_ds",""),
            //   new KeyValuePair<string, string>("mode","search")
            //};
            //var html = await HttpCaller.PostFormData("https://lsa.memberpro.net/main/body.cfm?menu=directory&submenu=directoryPractisingMember", format);
            //var doc = new HtmlAgilityPack.HtmlDocument();
            //doc.LoadHtml(html);
            //var suggestions = doc.DocumentNode.SelectNodes("//tr[contains(@class,'table-row')]");
            //var headerName = doc.DocumentNode.SelectSingleNode("//div[@class='content-heading']")?.InnerText.Trim();
            //if (suggestions != null && headerName == null)
            //{
            //    ReportNotFoundedLawyer(input, suggestions);
            //}
            //html = AddCssAndStyle(html);
            //doc.LoadHtml(html);
            ////doc.DocumentNode.SelectSingleNode("//div[@class='instructions-text']").Remove();
            ////doc.DocumentNode.SelectSingleNode("//div[@id='message']").Remove();
            //html = doc.DocumentNode.OuterHtml;
            //doc.Save("test.html");
            //byte[] res = null;
            //using (MemoryStream ms = new MemoryStream())
            //{
            //    var pdf = TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerator.GeneratePdf(html, PdfSharp.PageSize.RA1, -1);
            //    pdf.Save(ms);
            //    res = ms.ToArray();
            //    File.WriteAllBytes($"pdfs/{input.LastName}.pdf", res);
            //} 
            #endregion
            var socityAlbertaSuggestions = _lawyersFromlsaMemberpro.FindAll(x => x.LastName.ToLower() == input.LastName.ToLower());
            socityAlbertaSuggestions.Save($"Inputs/{input.FirstName} {input.LastName} suggestions.xlsx", "Law socity Alberta");
        }

        private void ReportNotFoundedLawyer(LawyerInputData input, HtmlAgilityPack.HtmlNodeCollection suggestions)
        {
            var suggetionsFounded = new List<Lawyer>();
            foreach (var suggestion in suggestions)
            {
                var name = suggestion.SelectSingleNode("./td[1]").InnerText.Trim();
                var url = suggestion.SelectSingleNode("./td[1]/a").GetAttributeValue("href", "");
                var city = suggestion.SelectSingleNode("./td[2]").InnerText.Trim();
                var gender = suggestion.SelectSingleNode("./td[3]").InnerText.Trim();
                var practisingStatus = suggestion.SelectSingleNode("./td[4]").InnerText.Trim();
                var enrolmentDate = suggestion.SelectSingleNode("./td[5]").InnerText.Trim();
                var firm = suggestion.SelectSingleNode("./td[6]").InnerText.Trim();
                suggetionsFounded.Add(new Lawyer
                {
                    //Name = name,
                    Url = url,
                });
            }
            suggetionsFounded.SaveToExcel($"suggestions for not founded lawyers/Suggestions for {input.FirstName} {input.LastName}.xlsx");
        }

        private string AddCssAndStyle(string html)
        {
            var x = html.IndexOf("<DIV ID=\"session-overlay\"></DIV>");
            html = html.Replace("//fonts.googleapis.com/css?family=Open+Sans", LawSocietyOfAlbertaStyleAndCss._css).
                Replace("/include/body/styles/animate.css", LawSocietyOfAlbertaStyleAndCss._animate).
                Replace("/include/body/styles/datepicker-ui.css", LawSocietyOfAlbertaStyleAndCss._datepicker).
                Replace("/include/body/styles/site_colors.cfm?licensee_cl=LSA", LawSocietyOfAlbertaStyleAndCss._site_colors).
                Replace("/include/body/styles/site_styles.css", LawSocietyOfAlbertaStyleAndCss._site_styles).
                Replace("/include/body/styles/override.cfm?licensee_cl=LSA", LawSocietyOfAlbertaStyleAndCss._override).
                Replace("/include/body/js/jquery.0.11.min.js", LawSocietyOfAlbertaStyleAndCss._jquery).
                Replace("/include/body/js/jquery-ui.js", LawSocietyOfAlbertaStyleAndCss._jqueryUi).
                Replace("/include/body/js/scripts.js", LawSocietyOfAlbertaStyleAndCss._script).
                Replace("/include/body/js/jquery.dataTables.min.js", LawSocietyOfAlbertaStyleAndCss._jquerytable).
                Replace("/include/body/styles/style.css", LawSocietyOfAlbertaStyleAndCss._style).
                Replace("/include/body/images/lsa_logo_v2.jpg", LawSocietyOfAlbertaStyleAndCss._img1).
                Replace("/include/body/images/memberpro-logo.png", LawSocietyOfAlbertaStyleAndCss._img2).
                Replace("<DIV ID=\"session-overlay\"></DIV>", "<DIV ID=\"session-overlay\" style=\"display: none;\"></DIV>").
                Replace("<DIV ID=\"sideMenu\" CLASS=\"side-menu\"", "<div id=\"sideMenu\" class=\"side-menu\" style=\"height: 1907px; \">").
                Replace("<DIV ID=\"session-message-center\"></DIV>", "");
            return html;
        }

        public static async Task WritePDF(string htmlPath, string pdfPath, string pdfConverterExe)
        {
            await Task.Delay(100);
            try
            {
                Process p;
                StreamWriter stdin;
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.UseShellExecute = false;
                psi.FileName = pdfConverterExe;
                psi.CreateNoWindow = true;
                psi.RedirectStandardInput = true;
                psi.RedirectStandardOutput = true;
                psi.RedirectStandardError = true;
                psi.Arguments = "";
                psi.Arguments += "--page-size A1";
                psi.Arguments += "--disable-smart-shrinking";
                psi.Arguments += "--print-media-type";
                psi.Arguments += "--margin-top 5mm --margin-bottom 5mm --margin-right 10mm --margin-left 30mm";
                psi.Arguments = "-q  -n - \"" + pdfPath + "\" ";


                p = Process.Start(psi);

                try
                {
                    stdin = p.StandardInput;
                    stdin.AutoFlush = true;
                    var html = File.ReadAllText(htmlPath, Encoding.UTF8);
                    StreamWriter utf8Writer = new StreamWriter(p.StandardInput.BaseStream, Encoding.UTF8);
                    utf8Writer.Write(html);
                    stdin.Close();

                    if (p.WaitForExit(15000))
                    {
                    }
                }
                finally
                {
                    p.Close();
                    p.Dispose();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            if (!Directory.Exists("pdfs"))
            {
                Directory.CreateDirectory("pdfs");
            }
            if (!Directory.Exists("reports"))
            {
                Directory.CreateDirectory("reports");
            }
            if (!Directory.Exists("lawyers not founded"))
            {
                Directory.CreateDirectory("lawyers not founded");
            }
            if (!Directory.Exists("suggestions for not founded lawyers"))
            {
                Directory.CreateDirectory("suggestions for not founded lawyers");
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
            var files = Directory.GetFiles("Inputs");
            var lawyersReportElements = new List<LawyerReportElements>();
            foreach (var file in files)
            {
                var lists = file.ReadAllSheetsFromExcel<Lawyer>();
                var lawyerReportElements = new LawyerReportElements();
                lawyerReportElements.LastName = lists[0][0].LastName;
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
                                    lawyerReportElements.LawocityAlberta = lawyer;
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
                lawyersReportElements.Add(lawyerReportElements);
            }

            return;
            await MainWork();
        }
    }
}
