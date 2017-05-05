using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HistoricalDataExport.Entities
{
    public class News
    {
        public int Id { get; set; }

        public string PublishDate { get; set; }

        public string Classification { get; set; }

        /// <summary>
        /// 区域经济地区
        /// </summary>
        public string AreaEconomyRegion { get; set; }

        /// <summary>
        /// 产业经济行业
        /// </summary>
        public string IndustrialEconomyVocation { get; set; }

        public string Url { get; set; }

        /// <summary>
        /// 区域经济中国
        /// </summary>
        public string AreaEconomyChina { get; set; }

        /// <summary>
        /// 区域经济境外
        /// </summary>
        public string AreaEconomyOversea { get; set; }

        public string Source { get; set; }

        public string Author { get; set; }

        public string Intro { get; set; }

        public string SubTitle { get; set; }

        public string Body { get; set; }

        public string Tag { get; set; }

        public string Title { get; set; }
    }
}
