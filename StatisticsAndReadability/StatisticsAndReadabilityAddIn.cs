using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace StatisticsAndReadability
{
    public partial class StatisticsAndReadabilityAddIn
    {


        private StatisticsAndReadabilityRibbon statisticsAndReadabilityRibbon;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            statisticsAndReadabilityRibbon = Globals.Ribbons.StatisticsAndReadabilityRibbon;
            this.Application.DocumentChange += Application_DocumentChange;
            this.statisticsAndReadabilityRibbon.UpdateStatsRequested += (o, ea) => Application_DocumentChange();
        }

        private void Application_DocumentChange()
        {
            documentLoaded = this.Application.Documents.Count > 0;
            TextChanged();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        

        private bool documentLoaded = false;

        private void TextChanged()
        {
            //this.Application.ActiveDocument.Readabil
            if(documentLoaded)
                try
                {
                    statisticsAndReadabilityRibbon.UpdateStats(this.Application.ActiveDocument.ReadabilityStatistics);
                }
                catch { }
        }



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }


}
