using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace StatisticsAndReadability
{
    public partial class StatisticsAndReadabilityRibbon
    {
        private void StatisticsAndReadabilityRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            ribbonUI = e.RibbonUI; 
        }

        private static IRibbonUI ribbonUI;
        public StatisticsAndReadabilityRibbon()
    : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            statisticsLabels = new[]
{
                this.wordsCountLabel,
                this.charactersCountLabel,
                paragraphsCountLabel,//Paragraphs,
                this.sentencesCountLabel,

                this.sentencesPerParagraphLabel,
                this.wordsPerSentenceLabel,
                this.charactersPerWordLabel,

                this.passiveSentencesLabel,//Passive Sentences,
                this.fleschReadingEaseLabel,
                this.fleschKincaidGradeLevelLabel,
            };
        }

        /*
         *  The statistics are ordered as follows:
         *  Words,
         *  Characters,
         *  Paragraphs,
         *  Sentences,
         *  Sentences per Paragraph,
         *  Words per Sentence,
         *  Characters per Word,
         *  Passive Sentences,
         *  Flesch Reading Ease,
         *  and Flesch-Kincaid Grade Level.
         */

        private RibbonLabel[] statisticsLabels;

        public void UpdateStats(ReadabilityStatistics readabilityStatistics)
        {
            try
            {
                int i = 0;
                foreach (ReadabilityStatistic readabilityStatistic in readabilityStatistics)
                {
                    if (statisticsLabels[i] != null)
                    {
                        //    try
                        {

                            statisticsLabels[i].Label = readabilityStatistic.Value.ToString();
                            if (statisticsLabels[i] == fleschReadingEaseLabel)
                            {
                                SetFleschReadingEaseDescription(readabilityStatistic.Value);
                            }


                            ribbonUI?.InvalidateControl(statisticsLabels[i].Id);
                        }
                        //    catch
                        {
                        }

                    }
                    i++;
                }
                ribbonUI?.Invalidate();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not run UpdateStats(), reason:\n" + ex.Message, "Statistics And Readability AddIn error");
            }
            //     this.RibbonUI.Invalidate();
            //    this.RibbonUI.InvalidateControl(this.RibbonId);
            //this.RibbonUI.
        }

        private void SetFleschReadingEaseDescription(double fleshReadingEase)
        {

            if (fleshReadingEase <= 30.0)
                fleschReadingEaseLabel.Label += " (college graduate)";
            else if (fleshReadingEase <= 50.0)
                fleschReadingEaseLabel.Label += " (college)";
            else if (fleshReadingEase <= 60.0)
                fleschReadingEaseLabel.Label += " (10th to 12th grade)";
            else if (fleshReadingEase <= 70.0)
                fleschReadingEaseLabel.Label += " (8th to 9th grade)";
            else if (fleshReadingEase <= 80.0)
                fleschReadingEaseLabel.Label += " (7th grade)";
            else if (fleshReadingEase <= 90.0)
                fleschReadingEaseLabel.Label += " (6th grade)";
            else
                fleschReadingEaseLabel.Label += " (5th grade)";
        }

        /* Flesch reading ease
        90.0–100.0	5th grade	Very easy to read. Easily understood by an average 11-year-old student.
        80.0–90.0	6th grade	Easy to read. Conversational English for consumers.
        70.0–80.0	7th grade	Fairly easy to read.
        60.0–70.0	8th & 9th grade	Plain English. Easily understood by 13- to 15-year-old students.
        50.0–60.0	10th to 12th grade	Fairly difficult to read.
        30.0–50.0	college	Difficult to read.
        0.0–30.0	college graduate	Very difficult to read. Best understood by university graduates.
        */


        public event EventHandler UpdateStatsRequested;

        private void recalculateButton_Click(object sender, RibbonControlEventArgs e)
        {
            UpdateStatsRequested?.Invoke(sender,new EventArgs());
        }
    }
}
