using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using ExcelDna.Integration;
using ExcelSummarizer.Properties;

namespace ExcelSummarizer
{
    public class Configuration
    {
        #region var

        /// <summary>
        /// Path of the file used as template for the summary
        /// </summary>
        public String TemplatePath
        {
            get { return Settings.Default.TemplatePath; }
            set
            {
                Settings.Default.TemplatePath = value;
                IsTemplateValid = File.Exists( value );
                Settings.Default.Save();

                if ( IsTemplateValid )
                {
                    InitTemplate();
                }
            }
        }

        /// <summary>
        /// Resource object: embbeded summary template file.
        /// </summary>
        /// <remarks>Used in case TemplatePath is null or points to an invalid file.</remarks>
        byte[] TemplateDefault { get { return Resources.templateDefault; } }

        bool _isTemplateValid;
        public bool IsTemplateValid
        {
            get { return _isTemplateValid; }
            protected set { _isTemplateValid = value; }
        }

        /// <summary>
        /// Path to the folder containing the target files to summarize.
        /// </summary>
        public String TargetPath
        {
            get { return Settings.Default.TargetPath; }
            set
            {
                Settings.Default.TargetPath = value;
                IsTargetValid = Directory.Exists( value );
                Settings.Default.Save();
            }
        }

        bool _isTargetValid;
        public bool IsTargetValid
        {
            get { return _isTargetValid; }
            protected set { _isTargetValid = value; }
        }


        public String OutputPath
        {
            get
            {
                return Path.Combine( "Summary.xlsx" );
            }
        }

        #endregion


        #region init
        public Configuration()
        {
            InitTemplate();

            InitTarget();
        }
        #endregion

        public bool InitTemplate()
        {
            bool valid = false;
            var ExcelApp = (Application)ExcelDnaUtil.Application;

            //---------------
            // init template path

            string templatePath = TemplatePath;

            if ( String.IsNullOrWhiteSpace( templatePath ) )
            {
                // create a temporary file from the embedded template
                templatePath = Path.Combine( Path.GetTempPath(), "~template.xlsx" );

                if ( File.Exists( templatePath ) )
                {
                    using ( var writer = new FileInfo( templatePath ).OpenWrite() )
                    {
                        var bytes = Resources.templateDefault;
                        writer.Write( bytes, 0, bytes.Length );
                        writer.Close();
                    }
                }
            }

            //---------------
            // open template

            try
            {
                if ( ExcelApp.Workbooks.Count > 0 && ExcelApp.ActiveWorkbook != null )
                    ExcelApp.ActiveWorkbook.Close();

                ExcelApp.Workbooks.Open( templatePath );

                valid = true;
            }
            catch
            {
                valid = false;
            }



            return IsTemplateValid = valid;
        }

        public bool InitTarget()
        {
            bool valid = false;

            if ( Directory.Exists( TargetPath ) )
            {
                valid = true;
            }

            return IsTargetValid = valid;
        }
    }
}
