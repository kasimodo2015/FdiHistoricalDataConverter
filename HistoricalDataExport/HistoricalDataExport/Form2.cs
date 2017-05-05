using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace HistoricalDataExport
{
    public partial class Form2 : Form
    {
        private BackgroundWorker _worker = new BackgroundWorker();
        private static string _changedField = string.Empty;

        public Form2()
        {
            InitializeComponent();

            _worker.DoWork += _worker_DoWork;
            _worker.RunWorkerCompleted += _worker_RunWorkerCompleted;
        }

        private void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Complete!");
        }

        private void _worker_DoWork(object sender, DoWorkEventArgs e)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(openFileDialog1.FileName);

            //var list = XMLHelper.GetXmlNodeListByXpath(doc, "/Orchard/Content/Legal/Legal.PromulgationDate");

            //foreach (XmlNode node in list)
            //{
            //    var newValue = node.Attributes["Value"] as XmlAttribute;

            //    if (newValue == null || string.IsNullOrEmpty(newValue.Value))
            //        continue;

            //    var dt = DateTime.Parse(newValue.Value).ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss");
            //    newValue.Value = dt.Replace('/', '-').Replace(' ', 'T').Trim() + 'Z';
            //}

            var list = XMLHelper.GetXmlNodeListByXpath(doc, "/Orchard/Content/News/EnumerationField.IndustrialEconomyVocation");

            foreach (XmlNode node in list)
            {
                var newValue = node.Attributes["Value"] as XmlAttribute;

                if (newValue == null || string.IsNullOrEmpty(newValue.Value))
                    continue;

                //newValue.Value = newValue.Value.Replace(" ", "-").Replace("“", string.Empty).Replace("”", string.Empty);
                newValue.Value = string.Format("{0}{1}{2}", ";;", newValue.Value.Replace(',', ';'), ";");
            }

            var newFileName = openFileDialog1.FileName.Replace("part1", "part1_new").Replace("part2", "part2_new");
            doc.Save(newFileName);
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "(*.recipe)|*.recipe";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string extension = Path.GetExtension(openFileDialog1.FileName);

                string[] str = new string[] { ".recipe" };

                if (!((IList)str).Contains(extension))
                {
                    MessageBox.Show("请选择xml文件！");
                    return;
                }

                this.textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            _worker.RunWorkerAsync();
        }
    }
}
