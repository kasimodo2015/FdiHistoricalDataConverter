using HistoricalDataExport.Entities;
using log4net;
using log4net.Config;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace HistoricalDataExport
{
    public partial class Form1 : Form
    {
        private const string xmlRootNodeName = "Orchard";
        private const string xmlDocumentVersion = "1.0";
        private const string xmlEncoding = "utf-8";
        private const string xmlStandalone = "yes";
        private BackgroundWorker _worker = new BackgroundWorker();
        private int _maxGenerationDataCount;
        private bool _isRunning = false;

        public Form1()
        {
            InitializeComponent();

            InitLog4Net();

            _worker.WorkerSupportsCancellation = true;
            _worker.DoWork += _worker_DoWork;
            _worker.RunWorkerCompleted += _worker_RunWorkerCompleted;
        }

        private void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _isRunning = false;
            button2.Text = "开 始";
            MessageBox.Show(string.Format("Success, total {0}", _maxGenerationDataCount));
        }

        private void _worker_DoWork(object sender, DoWorkEventArgs e)
        {
            var logger = LogManager.GetLogger(typeof(Program));

            FileInfo fileInfo = new FileInfo(openFileDialog1.FileName);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                if (package.Workbook.Worksheets.Count > 0)
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets[1];

                    int minColumnNum = ws.Dimension.Start.Column;//工作区开始列
                    int maxColumnNum = ws.Dimension.End.Column; //工作区结束列
                    int minRowNum = ws.Dimension.Start.Row + 1; //工作区开始行号
                    int maxRowNum = ws.Dimension.End.Row; //工作区结束行号

                    _maxGenerationDataCount = maxRowNum - 1;
                    logger.Info(string.Format("总行数：{0}", _maxGenerationDataCount));

                    var fileCategory = string.Empty;
                    var fileNames = new List<string>();
                    if (openFileDialog1.FileName.ToLower().Contains("news"))
                    {
                        fileCategory = DataCategory.News.ToString();
                        fileNames.Add("fdi-news-recipe-1.0.recipe");
                    }
                    else if (openFileDialog1.FileName.ToLower().Contains("laws"))
                    {
                        fileCategory = DataCategory.Legal.ToString();
                        fileNames.Add("fdi-inward-legal-recipe-1.0.recipe");
                        fileNames.Add("fdi-outward-legal-recipe-1.0.recipe");
                    }
                    else if (openFileDialog1.FileName.ToLower().Contains("data-statistics"))
                    {
                        fileCategory = DataCategory.DataStatistics.ToString();
                        fileNames.Add("fdi-data-statistics-recipe-1.0.recipe");
                    }

                    #region 循环
                    foreach (var fileName in fileNames)
                    {
                        CreateFileAndBuildRecipeContent(fileName);

                        var index = 0;
                        var inwardLegalIndex = 0;
                        var outwardLegalIndex = 0;

                        for (int n = minRowNum; n <= maxRowNum; n++)
                        {
                            var id = ws.Cells[string.Format("A{0}", n)].Value;

                            if (id == null)
                                continue;

                            Regex reg = new Regex("\\d{7}");
                            if (!reg.IsMatch(id.ToString()) || !int.TryParse(id.ToString(), out int result))
                            {
                                continue;
                            }

                            if (fileCategory == DataCategory.News.ToString())
                            {
                                var news = new News()
                                {
                                    Id = int.Parse(id.ToString())
                                };

                                #region Set Data

                                var publishDate = ws.Cells[string.Format("J{0}", n)];
                                if (publishDate.Value != null)
                                {
                                    news.PublishDate = publishDate.Value.ToString();
                                }

                                var title = ws.Cells[string.Format("Y{0}", n)];
                                if (title.Value != null)
                                {
                                    news.Title = title.Value.ToString();
                                }

                                var body = ws.Cells[string.Format("Z{0}", n)];
                                if (body.Value != null)
                                {
                                    news.Body = body.Value.ToString();
                                }

                                var classification = ws.Cells[string.Format("AC{0}", n)];
                                if (classification.Value != null)
                                {
                                    news.Classification = "Business Economy";//classification.Value.ToString();
                                }

                                var tag = ws.Cells[string.Format("AD{0}", n)];
                                if (tag.Value != null)
                                {
                                    news.Tag = tag.Value.ToString();
                                }

                                var source = ws.Cells[string.Format("AE{0}", n)];
                                if (source.Value != null)
                                {
                                    news.Source = source.Value.ToString();
                                }

                                var author = ws.Cells[string.Format("AF{0}", n)];
                                if (author.Value != null)
                                {
                                    news.Author = author.Value.ToString();
                                }

                                var vocation = ws.Cells[string.Format("AG{0}", n)];
                                if (vocation.Value != null)
                                {
                                    news.IndustrialEconomyVocation = vocation.Value.ToString();
                                }

                                var areaRegion = ws.Cells[string.Format("AH{0}", n)];
                                if (areaRegion.Value != null)
                                {
                                    news.AreaEconomyRegion = areaRegion.Value.ToString();
                                }

                                var areaChina = ws.Cells[string.Format("AI{0}", n)];
                                if (areaChina.Value != null)
                                {
                                    news.AreaEconomyChina = areaChina.Value.ToString();
                                }

                                var areaOversea = ws.Cells[string.Format("AJ{0}", n)];
                                if (areaOversea.Value != null)
                                {
                                    news.AreaEconomyOversea = areaOversea.Value.ToString();
                                }

                                var subtitle = ws.Cells[string.Format("AK{0}", n)];
                                if (subtitle.Value != null)
                                {
                                    news.SubTitle = subtitle.Value.ToString();
                                }
                                #endregion

                                index++;

                                BuildNewsContentXmlFile(fileName, news, index);
                            }
                            else if (fileCategory == DataCategory.Legal.ToString())
                            {
                                var legal = new Legal()
                                {
                                    Id = int.Parse(id.ToString())
                                };

                                #region Set Data

                                var publishDate = ws.Cells[string.Format("J{0}", n)];
                                if (publishDate.Value != null)
                                {
                                    legal.PublishDate = publishDate.Value.ToString();
                                }

                                var title = ws.Cells[string.Format("Y{0}", n)];
                                if (title.Value != null)
                                {
                                    legal.Title = title.Value.ToString();
                                }

                                var body = ws.Cells[string.Format("Z{0}", n)];
                                if (body.Value != null)
                                {
                                    legal.Body = body.Value.ToString();
                                }

                                var classification = ws.Cells[string.Format("AC{0}", n)];
                                if (classification.Value != null)
                                {
                                    var classStr = classification.Value.ToString().Trim();
                                    if (classStr == "“Going global” Laws and Regulations")
                                    {
                                        legal.Type = LegalType.Outward;
                                        legal.Classification = GetOutwardLegalClassification(ws.Cells[string.Format("AL{0}", n)].Value);

                                        outwardLegalIndex++;
                                    }
                                    else
                                    {
                                        legal.Type = LegalType.Inward;

                                        if (classStr == "“Bringing in” Laws and Regulations")
                                        {
                                            legal.Classification = GetInwardLegalClassification(ws.Cells[string.Format("AK{0}", n)].Value);
                                        }
                                        else if (classStr == "Comprehensive Laws and Regulations")
                                        {
                                            legal.Classification = GetInwardLegalClassification(ws.Cells[string.Format("AJ{0}", n)].Value);
                                        }
                                        else if (classStr == "Other")
                                        {
                                            legal.Classification = @"/alias=inward-laws-classfication\/other";
                                        }

                                        inwardLegalIndex++;
                                    }
                                }

                                var promulgationDate = ws.Cells[string.Format("AD{0}", n)];
                                if (promulgationDate.Value != null)
                                {
                                    legal.PromulgationDate = promulgationDate.Value.ToString();
                                }

                                var promulgationNumber = ws.Cells[string.Format("AE{0}", n)];
                                if (promulgationNumber.Value != null)
                                {
                                    legal.PromulgationNumber = promulgationNumber.Value.ToString();
                                }

                                var promulgationDepartment = ws.Cells[string.Format("AF{0}", n)];
                                if (promulgationDepartment.Value != null)
                                {
                                    legal.PromulgationDepartment = promulgationDepartment.Value.ToString();
                                }

                                var subtitle = ws.Cells[string.Format("AH{0}", n)];
                                if (subtitle.Value != null)
                                {
                                    legal.SubTitle = subtitle.Value.ToString();
                                }

                                var tag = ws.Cells[string.Format("AI{0}", n)];
                                if (tag.Value != null)
                                {
                                    legal.SubTitle = tag.Value.ToString();
                                }

                                #endregion

                                if (fileName.Contains(legal.Type.ToString().ToLower()))
                                {
                                    if(legal.Type == LegalType.Inward)
                                        BuildLegalContentXmlFile(fileName, legal, inwardLegalIndex);
                                    else
                                        BuildLegalContentXmlFile(fileName, legal, outwardLegalIndex);
                                }
                            }
                            else if (fileCategory == DataCategory.DataStatistics.ToString())
                            {
                                var dataStatistics = new DataStatistics()
                                {
                                    Id = int.Parse(id.ToString())
                                };

                                #region Set Data
                                var publishDate = ws.Cells[string.Format("J{0}", n)];
                                if (publishDate.Value != null)
                                {
                                    dataStatistics.PublishDate = publishDate.Value.ToString();
                                }

                                var title = ws.Cells[string.Format("Y{0}", n)];
                                if (title.Value != null)
                                {
                                    dataStatistics.Title = title.Value.ToString();
                                }

                                var body = ws.Cells[string.Format("Z{0}", n)];
                                if (body.Value != null)
                                {
                                    dataStatistics.Body = body.Value.ToString();
                                }

                                var classification = ws.Cells[string.Format("AC{0}", n)];
                                if (classification.Value != null)
                                {
                                    dataStatistics.Classification = classification.Value.ToString();
                                }

                                var foreignCategory = ws.Cells[string.Format("AD{0}", n)];
                                if(foreignCategory.Value != null)
                                {
                                    dataStatistics.ForeignInvestmentStatisticsCategory = foreignCategory.Value.ToString();
                                }

                                var chineseCategory = ws.Cells[string.Format("AE{0}", n)];
                                if (chineseCategory.Value != null)
                                {
                                    dataStatistics.ChineseEconomyStatisticsCategory = chineseCategory.Value.ToString();
                                }

                                var tag = ws.Cells[string.Format("AG{0}", n)];
                                if (tag.Value != null)
                                {
                                    dataStatistics.SubTitle = tag.Value.ToString();
                                }

                                var declareDate = ws.Cells[string.Format("AH{0}", n)];
                                if (declareDate.Value != null)
                                {
                                    dataStatistics.DeclareDate = declareDate.Value.ToString();
                                }

                                var overseaCategory = ws.Cells[string.Format("AI{0}", n)];
                                if (overseaCategory.Value != null)
                                {
                                    dataStatistics.OverseasInvestmentStatisticsCategory = overseaCategory.Value.ToString();
                                }

                                var subtitle = ws.Cells[string.Format("AJ{0}", n)];
                                if (subtitle.Value != null)
                                {
                                    dataStatistics.SubTitle = subtitle.Value.ToString();
                                }

                                var yearMonthLabel = ws.Cells[string.Format("AL{0}", n)];
                                if (yearMonthLabel.Value != null)
                                {
                                    dataStatistics.YearAndMonthLabel = yearMonthLabel.Value.ToString();
                                }

                                var source = ws.Cells[string.Format("AM{0}", n)];
                                if (source.Value != null)
                                {
                                    dataStatistics.Source = source.Value.ToString();
                                }
                                #endregion

                                index++;
                                BuildDataStatisticsContentXmlFile(fileName, dataStatistics, index);
                            }

                            if (_worker.CancellationPending)
                            {
                                e.Cancel = true;
                                return;
                            }
                        }
                    }
                    #endregion

                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "(*.xlsx)|*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string extension = Path.GetExtension(openFileDialog1.FileName);

                string[] str = new string[] { ".xlsx" };

                if (!((IList)str).Contains(extension))
                {
                    MessageBox.Show("请选择Excel文件！");
                    return;
                }

                this.textBox1.Text = openFileDialog1.FileName;
            }
        }

        private static string GetOutwardLegalClassification(object realCell)
        {
            var result = string.Empty;
            if (realCell != null)
            {
                var realCellValue = realCell.ToString();
                if (!string.IsNullOrEmpty(realCellValue))
                {
                    var classificationArray = realCellValue.Split(',');

                    if (classificationArray.Length > 0)
                    {
                        result = string.Format(@"/alias=inward-laws-classfication\/{0}",
                                classificationArray[0].ToLower());
                    }
                }
            }

            return result;
        }

        private static string GetInwardLegalClassification(object realCell)
        {
            var result = string.Empty;
            if (realCell != null)
            {
                var realCellValue = realCell.ToString();
                if (!string.IsNullOrEmpty(realCellValue))
                {
                    var classificationArray = realCellValue.Split(',');

                    if (classificationArray.Length > 0)
                    {
                        if (classificationArray.Length > 1)
                        {
                            result = string.Format(@"/alias=inward-laws-classfication\/{0}\/{1}",
                                classificationArray[0].ToLower(), classificationArray[1].ToLower());
                        }
                        else
                        {
                            result = string.Format(@"/alias=inward-laws-classfication\/{0}",
                                classificationArray[0].ToLower());
                        }
                    }
                }
            }

            return result;
        }

        private void CreateFileAndBuildRecipeContent(string fileName)
        {
            XMLHelper.CreateXmlDocument(fileName, xmlRootNodeName, xmlDocumentVersion, xmlEncoding, xmlStandalone);
            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, "/Orchard", "Recipe", "");
            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, "/Orchard/Recipe", "ExportUtc", "2017-04-14T01:33:13.3099906Z");
            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, "/Orchard/Recipe", "Name", "Fdi News Recipe");
            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, "/Orchard/Recipe", "Description", "新闻导入模板");
            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, "/Orchard/Recipe", "Author", "admin");
            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, "/Orchard/Recipe", "Version", "1.0");
            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, "/Orchard/Recipe", "IsSetupRecipe", "false");
            XMLHelper.CreateXmlNodeByXPath(fileName, "/Orchard", "Content", "", "", "");
        }

        private void BuildNewsContentXmlFile(string fileName, News news, int index)
        {
            BuildContentTypeRoot(fileName, "News", index, out string identifier);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "EnumerationField.Classification", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/EnumerationField.Classification", index), "Value", news.Classification);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "EnumerationField.AreaEconomyRegion", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/EnumerationField.AreaEconomyRegion", index), "Value", news.AreaEconomyRegion);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "EnumerationField.IndustrialEconomyVocation", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/EnumerationField.IndustrialEconomyVocation", index), "Value", news.IndustrialEconomyVocation);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "LinkField.Url", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/LinkField.Url", index), "Url", news.Url);
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/LinkField.Url", index), "Target", "_top");

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "TaxonomyField.AreaEconomyChina", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/TaxonomyField.AreaEconomyChina", index), "Terms", news.AreaEconomyChina);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "TaxonomyField.AreaEconomyOversea", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/TaxonomyField.AreaEconomyOversea", index), "Terms", news.AreaEconomyOversea);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "TextField.Source", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/TextField.Source", index), "Text", news.Source);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "TextField.Author", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/TextField.Author", index), "Text", news.Author);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "TextField.Intro", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/TextField.Intro", index), "Text", news.Intro);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "TextField.SubTitle", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/TextField.SubTitle", index), "Text", news.SubTitle);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "BodyPart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/BodyPart", index), "Text", news.Body);

            BuildCommonPart(fileName, index, news.PublishDate, identifier, "News");
            BuildPositionAndCulture(fileName, index);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "TagsPart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/TagsPart", index), "Tags", news.Tag);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "TitlePart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/TitlePart", index), "Title", news.Title);      
        }

        private void BuildLegalContentXmlFile(string fileName, Legal legal, int index)
        {
            BuildContentTypeRoot(fileName, "Legal", index, out string identifier);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]", index), "Legal.PromulgationDate", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]/Legal.PromulgationDate", index), "Value", DateTime.Parse(legal.PromulgationDate).ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss").Replace('/', '-').Replace(' ', 'T').Trim() + 'Z');

            if (legal.Type == LegalType.Inward)
            {
                XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]", index), "TaxonomyField.Classification", "");
                XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]/TaxonomyField.Classification", index), "Terms", legal.Classification);
            }
            else if (legal.Type == LegalType.Outward)
            {
                XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]", index), "EnumerationField.Classification", "");
                XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]/EnumerationField.Classification", index), "Value", legal.Classification);
            }

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]", index), "TextField.PromulgationDepartment", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]/TextField.PromulgationDepartment", index), "Text", legal.PromulgationDepartment);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]", index), "TextField.PromulgationNumber", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]/TextField.PromulgationNumber", index), "Text", legal.PromulgationNumber);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]", index), "TextField.SubTitle", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]/TextField.SubTitle", index), "Text", legal.SubTitle);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]", index), "BodyPart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]/BodyPart", index), "Text", legal.Body);

            BuildCommonPart(fileName, index, legal.PublishDate, identifier, "Legal");
            BuildPositionAndCulture(fileName, index);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]", index), "TagsPart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]/TagsPart", index), "Tags", legal.Tag);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]", index), "TitlePart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/Legal[{0}]/TitlePart", index), "Title", legal.Title);
        }

        private void BuildDataStatisticsContentXmlFile(string fileName, DataStatistics dataStatistifs, int index)
        {
            BuildContentTypeRoot(fileName, "DataStatistics", index, out string identifier);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]", index), "DataStatistics.PublishDate", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]/DataStatistics.PublishDate", index), "Value", dataStatistifs.DeclareDate);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]", index), "DataStatistics.YearAndMonthLabel", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]/DataStatistics.YearAndMonthLabel", index), "Value", dataStatistifs.YearAndMonthLabel);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]", index), "EnumerationField.Classification", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]/EnumerationField.Classification", index), "Value", dataStatistifs.Classification);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]", index), "EnumerationField.ChineseEconomyStatisticsCategory", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]/EnumerationField.ChineseEconomyStatisticsCategory", index), "Value", dataStatistifs.ChineseEconomyStatisticsCategory);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]", index), "EnumerationField.ForeignInvestmentStatisticsCategory", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]/EnumerationField.ForeignInvestmentStatisticsCategory", index), "Value", dataStatistifs.ForeignInvestmentStatisticsCategory);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]", index), "EnumerationField.OverseasInvestmentStatisticsCategory", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]/EnumerationField.OverseasInvestmentStatisticsCategory", index), "Value", dataStatistifs.OverseasInvestmentStatisticsCategory);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]", index), "TextField.Source", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]/TextField.Source", index), "Text", dataStatistifs.Source);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]", index), "TextField.SubTitle", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]/TextField.SubTitle", index), "Text", dataStatistifs.SubTitle);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]", index), "BodyPart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]/BodyPart", index), "Text", dataStatistifs.Body);

            BuildCommonPart(fileName, index, dataStatistifs.PublishDate, identifier, "DataStatistics");
            BuildPositionAndCulture(fileName, index);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]", index), "TagsPart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]/TagsPart", index), "Tags", dataStatistifs.Tag);

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]", index), "TitlePart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/DataStatistics[{0}]/TitlePart", index), "Title", dataStatistifs.Title);
        }

        private static void BuildContentTypeRoot(string fileName, string contentTypeTag, int index, out string identifier)
        {
            identifier = Guid.NewGuid().ToString();

            XMLHelper.CreateXmlNodeByXPath(fileName, "/Orchard/Content", contentTypeTag, "", "", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/{1}[{0}]", index, contentTypeTag), "Id", string.Format("/Identifier={0}", identifier));
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/{1}[{0}]", index, contentTypeTag), "Status", "Published");
        }

        private static void BuildPositionAndCulture(string fileName, int index)
        {
            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "ContainablePart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/ContainablePart", index), "Position", "0");

            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]", index), "LocalizationPart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/News[{0}]/LocalizationPart", index), "Culture", "en-US");
        }

        private static void BuildCommonPart(string fileName, int index, string publishDate, string identifier, string contentTypeTag)
        {
            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/{1}[{0}]", index, contentTypeTag), "IdentityPart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/{1}[{0}]/IdentityPart", index, contentTypeTag), "Identifier", identifier);

            publishDate = Convert.ToDateTime(publishDate).ToUniversalTime().ToString();
            XMLHelper.CreateOrUpdateXmlNodeByXPath(fileName, string.Format("/Orchard/Content/{1}[{0}]", index, contentTypeTag), "CommonPart", "");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/{1}[{0}]/CommonPart", index, contentTypeTag), "Owner", "/User.UserName=admin");
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/{1}[{0}]/CommonPart", index, contentTypeTag), "CreatedUtc", publishDate);
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/{1}[{0}]/CommonPart", index, contentTypeTag), "PublishedUtc", publishDate);
            XMLHelper.CreateOrUpdateXmlAttributeByXPath(fileName, string.Format("/Orchard/Content/{1}[{0}]/CommonPart", index, contentTypeTag), "ModifiedUtc", publishDate);
        }

        private static void InitLog4Net()
        {
            var logCfg = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + "log4net.config");
            XmlConfigurator.ConfigureAndWatch(logCfg);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!_isRunning)
            {
                _worker.RunWorkerAsync();
                _isRunning = true;
                button2.Text = "停 止";
            }
            else
            {
                _worker.CancelAsync();
                _isRunning = false;
                button2.Text = "开 始";
            }
        }
    }

    public enum DataCategory
    {
        News,

        Legal,

        DataStatistics
    }
}
