using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lawyers_informations_reporter.Models
{
    public class LawyerReportElements
    {
        public string LastName { get; set; }
        public Lawyer LawocityAlberta { get; set; }
        public List<Lawyer> OutComes { get; set; } = new List<Lawyer>();
        public List<Lawyer> Hearings { get; set; } = new List<Lawyer>();
    }
}
