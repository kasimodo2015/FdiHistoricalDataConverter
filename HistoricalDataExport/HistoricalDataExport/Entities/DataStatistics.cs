using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HistoricalDataExport.Entities
{
    public class DataStatistics
    {
        public int Id { get; set; }

        public string YearAndMonthLabel { get; set; }

        public string Classification { get; set; }

        public string ChineseEconomyStatisticsCategory { get; set; }

        public string ForeignInvestmentStatisticsCategory { get; set; }

        public string OverseasInvestmentStatisticsCategory { get; set; }

        public string Source { get; set; }

        public string SubTitle { get; set; }

        public string Body { get; set; }

        public string Tag { get; set; }

        public string Title { get; set; }

        public string PublishDate { get; set; }

        /// <summary>
        /// 发布日期
        /// </summary>
        public string DeclareDate { get; set; }
    }
}
