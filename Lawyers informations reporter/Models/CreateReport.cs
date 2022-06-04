using CoreHtmlToImage;
using Newtonsoft.Json.Linq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using scrapingTemplateV51.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using TheArtOfDev.HtmlRenderer.PdfSharp;

namespace Lawyers_informations_reporter.Models
{
    public class CreateReport
    {
        HttpCaller HttpCaller = new HttpCaller();
        public List<string> _years = new List<string>();
        public async Task CreatePdfs(List<LawyerReportElements> reportsElements)
        {
            var index = 1;
            var nbr = 1;
            foreach (var reportElements in reportsElements)
            {
                await CreateLawSocietyAlbertaPdf(reportElements.LawSocietyAlberta);
                foreach (var hearingCase in reportElements.Hearings)
                {
                    await CreateHearingCasesPdf(hearingCase, index);
                    index++;
                }
                await GetNewOutcomes(reportElements);
                foreach (var outcome in reportElements.OutComes)
                {
                    await DownLoadOutcomesPdfs(outcome, index);
                    index++;
                }
                index++;
                await CombinePdfs(reportElements.LawSocietyAlberta);
                Reporter.OnProgress(null,(nbr, reportsElements.Count, $"{reportElements.LastName}'s Report created"));
                nbr++;
            }

        }

        private async Task GetNewOutcomes(LawyerReportElements lawyer)
        {
            for (int i = 2006; i < 20000; i++)
            {
                var json = await HttpCaller.GetHtml($"https://www.canlii.org/en/ab/abls/nav/date/{i}/items");
                var array = JArray.Parse(json);
                if (array.Count==0)
                {
                    break;
                }
                var urls = lawyer.OutComes.Select(xx => xx.Url).ToList();
                var lasNames = lawyer.OutComes.Select(xx => xx.LastName).ToList();
                foreach (var node in array)
                {
                    var lastName = (string)node.SelectToken("..styleOfCause");
                    var x = lastName.LastIndexOf(" ");
                    lastName = lastName.Substring(x + 1);
                    var url = "https://www.canlii.org" + ((string)node.SelectToken("..url"));
                    if (!urls.Contains(url) && lawyer.LawSocietyAlberta.LastName.ToLower() == lastName.ToLower())
                    {
                        var isMatch = await CkechIsMatch(url, lawyer.OutComes);
                        if (isMatch)
                        {
                            var date = ((string)node.SelectToken("..judgmentDate"));
                            lawyer.OutComes.Add(new Lawyer { LastName = lastName, Url = url });
                        }
                    }
                }
            }
        }
        private async Task<bool> CkechIsMatch(string url, List<Lawyer> outcomes)
        {
            var ismatch = false;
            var firsNames = outcomes.Select(x => x.FirstName).ToList();
            var html = await HttpCaller.GetHtml(url);
            foreach (var firsName in firsNames)
            {
                if (firsName != null)
                {
                    if (html.ToLower().Contains(firsName.ToLower()))
                    {
                        ismatch = true;
                        break;
                    }
                }
            }
            return ismatch;
        }

        private async Task CombinePdfs(Lawyer input)
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
                outputPDFDocument.Save($"reports/{input.FirstName} {input.LastName}.pdf");
                foreach (var pdfFile in pdfFiles)
                {
                    File.Delete(pdfFile);
                }
            }
        }
        public async Task CreateHearingCasesPdf(Lawyer hearingCase, int index)
        {
            var doc = await HttpCaller.GetDoc(hearingCase.Url);
            var scripts = doc.DocumentNode.SelectNodes("//script").ToList();
            foreach (var script in scripts)
                script.Remove();
            var divs = doc.DocumentNode.SelectNodes("//div[contains(@id,'pum')]");
            foreach (var div in divs)
                div.Remove();
            var header = doc.DocumentNode.SelectSingleNode("//div[@class='header-container']");
            header.Remove();
            var html = doc.DocumentNode.OuterHtml;
            var exePath = Path.GetFullPath("wkhtmltopdf.exe");
            await WritePDF(html, $"pdfs/{hearingCase.LastName}{index} hearing cases.pdf", exePath);
            do
            {
                if (File.Exists($"pdfs/{hearingCase.LastName}{index} hearing cases.pdf"))
                {
                    break;
                }
            } while (true);
            AwaitTillTheFileGetDisposed($"pdfs/{hearingCase.LastName}{index} hearing cases.pdf");

        }

        private void AwaitTillTheFileGetDisposed(string path)
        {
            var file = new FileInfo(path);
            do
            {
                try
                {
                    using (FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        stream.Close();
                        break;
                    }
                }
                catch (IOException)
                {
                }
            } while (true);
        }

        public static async Task WritePDF(string html, string pdfPath, string pdfConverterExe)
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
                //psi.Arguments = "";
                //psi.Arguments += "--page-size A1";
                //psi.Arguments += "--disable-smart-shrinking";
                //psi.Arguments += "--print-media-type";
                //psi.Arguments += "--margin-top 5mm --margin-bottom 5mm --margin-right 10mm --margin-left 30mm";
                psi.Arguments = "-q  -n - \"" + pdfPath + "\" ";


                p = Process.Start(psi);

                try
                {
                    stdin = p.StandardInput;
                    stdin.AutoFlush = true;
                    StreamWriter utf8Writer = new StreamWriter(p.StandardInput.BaseStream, Encoding.UTF8);
                    utf8Writer.Write(html);
                    stdin.Close();
                }
                finally
                {
                    p.Close();
                    p.Dispose();
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        public async Task DownLoadOutcomesPdfs(Lawyer outcome, int index)
        {
            using (WebClient client = new WebClient())
            {
                await client.DownloadFileTaskAsync(outcome.Url.Replace(".html", ".pdf"), $"pdfs/{outcome.LastName}{index} outcomes cases.pdf");
            }
        }

        public async Task CreateLawSocietyAlbertaPdf(Lawyer lawSocietyAlberta)
        {
            var html = await HttpCaller.GetHtml(lawSocietyAlberta.Url);
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);
            html = AddCssAndStyle(html);
            doc.LoadHtml(html);
            html = doc.DocumentNode.OuterHtml;
            byte[] res = null;
            using (MemoryStream ms = new MemoryStream())
            {
                var pdf = PdfGenerator.GeneratePdf(html, PdfSharp.PageSize.RA1, -1);
                pdf.Save(ms);
                res = ms.ToArray();
                File.WriteAllBytes($"pdfs/{lawSocietyAlberta.LastName} law society Alberta.pdf", res);
            }
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
    }
}
