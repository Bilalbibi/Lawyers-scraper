using Newtonsoft.Json.Linq;
using scrapingTemplateV51.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lawyers_informations_reporter.Models
{
    public class NewSearch
    {
        public HttpCaller HttpCaller = new HttpCaller();
        public List<string> _years = new List<string>();
        List<Lawyer> _lawyersFromlsaMemberpro = new List<Lawyer>();

        public async Task PreparingTheRequiredFilesForSearch(List<LawyerInputData> inputs, List<Lawyer> reportElements)
        {
            _lawyersFromlsaMemberpro = reportElements;
            await ScrapeYears();
            await ScrapingThePossiblitiesInputs(inputs);
        }
        public async Task ScrapeYears()
        {
            //var doc = await HttpCaller.GetDoc("https://www.canlii.org/en/ab/abls/");
            //doc.Save("canlii.html");
            //var nodes = doc.DocumentNode.SelectNodes("//select[@id='navYearsSelector']/option");
            //for (int i = 1; i < nodes.Count; i++)
            //{
            //    _years.Add(nodes[i].InnerText.Trim());
            //}
        }
        private async Task ScrapingThePossiblitiesInputs(List<LawyerInputData> inputs)
        {
            var nbr = 1;
            foreach (var input in inputs)
            {
                await CreateTheInputsBasedOnthelastName(input);
                Reporter.OnProgress(null, (nbr, inputs.Count, $"{input.LastName}'s suggestions excel file created"));
                nbr++;
            }
        }
        private async Task CreateTheInputsBasedOnthelastName(LawyerInputData input)
        {
            await GetDataFromLawSocityAlberta(input);
            await GetDecisionsAndOutcomes(input);
            await GetHearingSchedule(input);
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
            for (int i = 2006; i < 200000; i++)
            {
                var json = await HttpCaller.GetHtml($"https://www.canlii.org/en/ab/abls/nav/date/{i}/items");
                var array = JArray.Parse(json);
                if (array.Count==0)
                {
                    break;
                }
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
    }
}
