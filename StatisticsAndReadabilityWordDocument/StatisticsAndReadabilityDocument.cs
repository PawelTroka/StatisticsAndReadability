using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace StatisticsAndReadabilityWordDocument
{
    public partial class StatisticsAndReadabilityDocument
    {
        private void StatisticsAndReadabilityDocument_Startup(object sender, System.EventArgs e)
        {
            this.Application.DocumentOpen += Application_DocumentOpen;
        }

        private void Application_DocumentOpen(Word.Document Doc)
        {
         //   Object start = Doc.Content.Start;
         //   Object end = Doc.Content.End;
          //  Doc.Range(ref start, ref end).Select();
            if (string.IsNullOrEmpty(Doc.Content.Text)||string.IsNullOrWhiteSpace(Doc.Content.Text))
            {
                var s = File.ReadAllText(System.IO.Path.Combine(Application.CustomDictionaries.ActiveCustomDictionary.Path, Application.CustomDictionaries.ActiveCustomDictionary.Name) );
                Doc.Content.Text =s;
            }
            try
            {
                var sb = new StringBuilder("");
                foreach (Word.ReadabilityStatistic readabilityStatistic in Doc.ReadabilityStatistics)
                {
                    sb.AppendLine($"{readabilityStatistic.Name}: {readabilityStatistic.Value}");
                }
                MessageBox.Show(sb.ToString(), "STATISTICS AND READABILITY");
            }
            catch { }
        }

        private void StatisticsAndReadabilityDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(StatisticsAndReadabilityDocument_Startup);
            this.Shutdown += new System.EventHandler(StatisticsAndReadabilityDocument_Shutdown);
        }

        #endregion
    }
}
