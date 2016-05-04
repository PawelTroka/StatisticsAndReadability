using System;
using Office = Microsoft.Office.Core;

namespace StatisticsAndReadability
{
    public partial class StatisticsAndReadabilityAddIn
    {
        private bool _documentLoaded;


        private StatisticsAndReadabilityRibbon _statisticsAndReadabilityRibbon;

        private void StatisticsAndReadabilityAddIn_Startup(object sender, EventArgs e)
        {
            _statisticsAndReadabilityRibbon = Globals.Ribbons.StatisticsAndReadabilityRibbon;
            Application.DocumentChange += Application_DocumentChange;
            _statisticsAndReadabilityRibbon.UpdateStatsRequested += (o, ea) => Application_DocumentChange();
        }

        private void Application_DocumentChange()
        {
            _documentLoaded = Application.Documents.Count > 0;
            TextChanged();
        }

        private void StatisticsAndReadabilityAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        private void TextChanged()
        {
            //this.Application.ActiveDocument.Readabil
            if (!_documentLoaded) return;
            try
            {
                _statisticsAndReadabilityRibbon.UpdateStats(Application.ActiveDocument.ReadabilityStatistics);
            }
            catch
            {
            }
        }

        #region VSTO generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += StatisticsAndReadabilityAddIn_Startup;
            Shutdown += StatisticsAndReadabilityAddIn_Shutdown;
        }

        #endregion
    }
}