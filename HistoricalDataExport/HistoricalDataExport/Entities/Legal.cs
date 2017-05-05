using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HistoricalDataExport.Entities
{
    public class Legal
    {
        public int Id { get; set; }

        public string PromulgationDate { get; set; }

        public string PromulgationDepartment { get; set; }

        public string PromulgationNumber { get; set; }

        public string SubTitle { get; set; }

        public string Body { get; set; }

        public string Tag { get; set; }

        public string Title { get; set; }

        public string Classification { get; set; }

        public string PublishDate { get; set; }

        public LegalType Type { get; set; }
    }

    public enum LegalType
    {
        Inward,

        Outward
    }
}
