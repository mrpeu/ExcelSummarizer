using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Excel;
using ExcelDna.Integration;
using ExcelSummarizer.Properties;
using System.Threading;
using System.Diagnostics;
using ExcelSummarizer.Parser;

namespace ExcelSummarizer
{
    public class AddIn : IExcelAddIn
    {
        #region var
        static Configuration Configuration;

        static Ribbon ribbon;

        static SettingsPanel _settingsPanel;
        static SettingsPanel SettingsPanel
        {
            get
            {
                if ( _settingsPanel == null )
                {
                    _settingsPanel = new SettingsPanel();
                }

                return _settingsPanel;
            }
        }

        static StringBuilder LogBuilder = new StringBuilder();
        #endregion


        #region init

        public void AutoOpen()
        {
            Configuration = new Configuration();
        }

        public void AutoClose()
        {
            //throw new NotImplementedException();
        }

        internal static void RegisterRibbon( Ribbon ribbon )
        {
            ribbon.Texts[ Ribbon.EControlIds.txt_template ] = Configuration.TemplatePath;
            ribbon.Texts[ Ribbon.EControlIds.txt_target ] = Configuration.TargetPath;

            ribbon.Images[ Ribbon.EControlIds.txt_template ] = Configuration.IsTemplateValid ? Resources.bullet_green : Resources.bullet_pink;
            ribbon.Images[ Ribbon.EControlIds.txt_target ] = Configuration.IsTargetValid ? Resources.bullet_green : Resources.bullet_pink;

            ribbon.Invalidate();

            ribbon.SettingsClicked += ribbon_SettingsClicked;
            ribbon.SummaryClicked += ribbon_SummaryClicked;

            ribbon.TemplatePathChanged += ribbon_TemplatePathChanged;

            ribbon.TargetPathChanged += ribbon_TargetPathChanged;
        }

        #endregion


        static void ribbon_TemplatePathChanged( object sender, EventArgs<string> e )
        {
            Configuration.TemplatePath = e.Value;
            ribbon.UpdateConfiguration( Configuration );
        }

        static void ribbon_TargetPathChanged( object sender, EventArgs<string> e )
        {
            Configuration.TargetPath = e.Value;
            ribbon.UpdateConfiguration( Configuration );
        }

        static void ribbon_SettingsClicked( object sender, EventArgs e )
        {
            SettingsPanel.ShowDialog();
        }

        static CancellationTokenSource _cts = null;

        static void ribbon_SummaryClicked( object sender, EventArgs e )
        {
            if ( _cts == null )
            {
                _cts = new CancellationTokenSource();
                var sw = new Stopwatch();
                sw.Start();

                var task = Task.Factory.StartNew( () =>
                {
                    //BeginInvoke( (MethodInvoker)( () =>
                    //{
                    //    LogSB.Clear();
                    //    toolStripProgressBar1.Visible = toolStripStatusLabel1.Visible = true;
                    //    toolStripProgressBar1.Value = 0;
                    //    toolStripStatusLabel1.Text = "reading bills...";
                    //    button2.Text = "Cancel";
                    //    Cursor = Cursors.WaitCursor;
                    //} ) );
                    LogBuilder.Clear();

                    if ( _cts.IsCancellationRequested ) { EndGeneration(); return; }


                    //============
                    // parse files

                    Log( "//=============" );
                    Log( "// start." );

                    List<InputDoc> bills = Summarist.Summarize( Configuration.TargetPath, _cts, Log, Progress );

                    if ( _cts.IsCancellationRequested ) { EndGeneration(); return; }


                    //============
                    // create summary

                    Log( "//=============" );
                    bool summaryCreated = false;

                    //BeginInvoke( (MethodInvoker)( () =>
                    //{
                    //    toolStripProgressBar1.Value = 0;
                    //    toolStripStatusLabel1.Text = "summarizing...";
                    //} ) );

                    try
                    {
                        summaryCreated = Summarist.CreateSummary(
                            bills,
                             _cts, Log, Progress
                        );
                    }
                    catch { }


                    //=============
                    // the end

                    if ( summaryCreated )
                    {
                        Log( String.Format( "Summary created [{0}]", Configuration.OutputPath ) );
                    }
                    else
                    {
                        Log( String.Format( "Summary creation failed! [{0}]", Configuration.OutputPath ) );
                    }

                    //BeginInvoke( (MethodInvoker)( () =>
                    //{
                    //    Log( String.Format( "// done. ({0})", sw.Elapsed.ToString() ) );
                    //    Log( "//=============" );
                    //} ) );

                    EndGeneration();

                }, _cts.Token );
            }
            else
            {
                _cts.Cancel();
            }
        }

        static void EndGeneration()
        {
            if ( _cts.IsCancellationRequested )
            {
                Log( "\nUser canceled." );
            }

            //toolStripProgressBar1.Value = 0;
            //toolStripProgressBar1.Visible = false;
            //toolStripStatusLabel1.Visible = false;
            //richTextBox1.AppendText( LogSB.ToString() );
            //richTextBox1.ScrollToCaret();
            //richTextBox1.SelectAll();
            //button2.Text = "Generate summary";
            //Cursor = Cursors.Default;
            _cts = null;


            var logPath = Path.Combine( Path.GetTempPath(), "ExcelSummarizerLog.txt" );
            File.WriteAllText( logPath, LogBuilder.ToString() );
            System.Diagnostics.Process.Start( logPath );
        }

        static void Log( string message )
        {
            LogBuilder.AppendLine( message );
        }

        static void Progress( int i, int max )
        {
            //if ( toolStripProgressBar1.Maximum != max )
            //    toolStripProgressBar1.Maximum = max;

            //toolStripProgressBar1.Value = i;
        }
    }
}
