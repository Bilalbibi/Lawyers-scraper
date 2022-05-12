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
        private int _maxConcurrency;
        private Dictionary<string, string> _config;
        public HttpCaller HttpCaller = new HttpCaller();
        List<LawyerInputData> _inputs;
        List<string> _years = new List<string>();
        string _css = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/css").FullName;
        string _style = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/style.css").FullName;
        string _animate = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/animate.css").FullName;
        string _datepicker = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/datepicker-ui.css").FullName;
        string _site_colors = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/site_colors.cfm").FullName;
        string _site_styles = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/site_styles.css").FullName;
        string _override = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/override.cfm").FullName;
        string _img1 = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/lsa_logo_v2.jpg").FullName;
        string _img2 = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/memberpro-logo.png").FullName;
        string _jquery = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/jquery.0.11.min.js.download").FullName;
        string _jquerytable = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/jquery.dataTables.min.js.download").FullName;
        string _jqueryUi = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/jquery-ui.js.download").FullName;
        string _script = new FileInfo("Law Society of Alberta style and css/Law Society of Alberta_files/scripts.js.download").FullName;
        public MainForm()
        {
            InitializeComponent();
        }


        private async Task MainWork()
        {
            if (!File.Exists(inputI.Text) || inputI.Text != "")
            {
                _inputs = ExcelHelperExe.ExcelHelper.ReadFromExcel<LawyerInputData>(inputI.Text);
            }
            else
            {
                MessageBox.Show("Please select the excel input data");
                return;
            }
            await ScrapeYears();
            await StartScraping();
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

        private async Task StartScraping()
        {
            foreach (var input in _inputs)
            {
                await CreateReport(input);
                await CombinePdfs(input);
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

        private async Task CreateReport(LawyerInputData input)
        {
            var toContinue = await GetDataFromLawSocityAlberta(input);
            //if (true)
            //{
            //    await GetDecisionsAndOutcomes(input);
            //}
            //var htmlPath = Path.GetFullPath("test.html");
            //var pdfPath = Path.GetFullPath("test.pdf");
            //var exePath = Path.GetFullPath("wkhtmltopdf.exe");
            //await WritePDF(htmlPath, pdfPath, exePath);
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
            var outcomesHistory = new List<OutcomeHistory>();
            for (int i = 0; i < _years.Count; i++)
            {
                var json = await HttpCaller.GetHtml($"https://www.canlii.org/en/ab/abls/nav/date/{_years[i]}/items");
                var array = JArray.Parse(json);
                foreach (var node in array)
                {
                    var styleOfCause = ((string)node.SelectToken("..styleOfCause"));
                    var url = "https://www.canlii.org" + ((string)node.SelectToken("..url"));
                    var date = ((string)node.SelectToken("..judgmentDate"));
                    outcomesHistory.Add(new OutcomeHistory { StyleOfCause = styleOfCause, Url = url, Date = date });
                }
            }
            var lawyerOutcomesHistories = outcomesHistory.FindAll(x => x.StyleOfCause.ToLower().Contains(input.LastName.ToLower()));
            if (lawyerOutcomesHistories.Count > 0)
            {
                foreach (var lawyerOutcomeHistory in lawyerOutcomesHistories)
                {
                    WebClient webClient = new WebClient();
                    webClient.DownloadFile(lawyerOutcomeHistory.Url.Replace(".html", ".pdf"), $"pdfs/{input.FirstName} {input.LastName} {lawyerOutcomeHistory.Date}.pdf");
                }
            }
            #region MyRegion
            //var array = JArray.Parse(json);
            //foreach (var node in array)
            //{
            //    var name = ((string)node.SelectToken("..styleOfCause")).Replace(" v ", "$").Split('$')[1];
            //    var url = ((string)node.SelectToken("..url")).Replace(".html", ".pdf");
            //    var doc1 = await HttpCaller.GetDoc(url);
            //    url = "https://www.canlii.org" + url;
            //    WebClient webClient = new WebClient();
            //    webClient.DownloadFile(url, $"pdfs/{input.FirstName} {name}.pdf");
            //} 
            #endregion
        }

        private async Task<bool> GetDataFromLawSocityAlberta(LawyerInputData input)
        {
            var format = new List<KeyValuePair<string, string>>()
            {
               new KeyValuePair<string, string>("person_nm",input.LastName),
               new KeyValuePair<string, string>("first_nm",""),
               new KeyValuePair<string, string>("member_status_cl","PRAC"),
               new KeyValuePair<string, string>("city_nm",""),
               new KeyValuePair<string, string>("location_nm",input.Firm),
               new KeyValuePair<string, string>("gender_cl",""),
               new KeyValuePair<string, string>("language_cl",""),
               new KeyValuePair<string, string>("area_ds",""),
               new KeyValuePair<string, string>("mode","search")
            };
            var html = await HttpCaller.PostFormData("https://lsa.memberpro.net/main/body.cfm?menu=directory&submenu=directoryPractisingMember", format);
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);
            var suggestions = doc.DocumentNode.SelectNodes("//tr[contains(@class,'table-row')]");
            var headerName = doc.DocumentNode.SelectSingleNode("//div[@class='content-heading']")?.InnerText.Trim();
            if (suggestions != null && headerName == null)
            {
                 ReportNotFoundedLawyer(input, suggestions);
            }
            return true;
            html = AddCssAndStyle(html);
            doc.LoadHtml(html);
            //doc.DocumentNode.SelectSingleNode("//div[@class='instructions-text']").Remove();
            //doc.DocumentNode.SelectSingleNode("//div[@id='message']").Remove();
            html = doc.DocumentNode.OuterHtml;
            doc.Save("test.html");

            byte[] res = null;
            using (MemoryStream ms = new MemoryStream())
            {
                var pdf = TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerator.GeneratePdf(html, PdfSharp.PageSize.RA1, -1);
                pdf.Save(ms);
                res = ms.ToArray();
                File.WriteAllBytes($"{input.LastName}.pdf", res);
            }
            return true;
        }

        private void ReportNotFoundedLawyer(LawyerInputData input, HtmlAgilityPack.HtmlNodeCollection suggestions)
        {
            var suggetionsFounded = new List<SuggestedLawyers>();
            foreach (var suggestion in suggestions)
            {

            }
        }

        private string AddCssAndStyle(string html)
        {
            var x = html.IndexOf("<DIV ID=\"session-overlay\"></DIV>");
            html = html.Replace("//fonts.googleapis.com/css?family=Open+Sans", _css).
                Replace("/include/body/styles/animate.css", _animate).
                Replace("/include/body/styles/datepicker-ui.css", _datepicker).
                Replace("/include/body/styles/site_colors.cfm?licensee_cl=LSA", _site_colors).
                Replace("/include/body/styles/site_styles.css", _site_styles).
                Replace("/include/body/styles/override.cfm?licensee_cl=LSA", _override).
                Replace("/include/body/js/jquery.0.11.min.js", _jquery).
                Replace("/include/body/js/jquery-ui.js", _jqueryUi).
                Replace("/include/body/js/scripts.js", _script).
                Replace("/include/body/js/jquery.dataTables.min.js", _jquerytable).
                Replace("/include/body/styles/style.css", _style).
                Replace("/include/body/images/lsa_logo_v2.jpg", _img1).
                Replace("/include/body/images/memberpro-logo.png", _img2).
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
            await MainWork();
        }
    }
}
