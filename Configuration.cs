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
        String templatePath = Resources.TemplatePathDefault;
        public String TemplatePath
        {
            get { return templatePath; }
            set { templatePath = value; }
        }

        /// <summary>
        /// Resource object: embbeded summary template file.
        /// </summary>
        /// <remarks>Used in case TemplatePath is null or points to an invalid file.</remarks>
        byte[] TemplateDefault { get { return Resources.template; } }

        bool _isTemplateValid;
        public bool IsTemplateValid
        {
            get { return _isTemplateValid; }
            protected set { _isTemplateValid = value; }
        }

        /// <summary>
        /// Path of the generated summary
        /// </summary>
        String outputPath = Resources.OutputPathDefault;
        public String OutputPath
        {
            get { return outputPath; }
            set { outputPath = value; }
        }

        /// <summary>
        /// Path to the folder containing the target files to summarize.
        /// </summary>
        String targetPath;
        public String TargetPath
        {
            get { return targetPath; }
            set { targetPath = value; }
        }

        bool _isTargetValid;
        public bool IsTargetValid
        {
            get { return _isTargetValid; }
            protected set { _isTargetValid = value; }
        }

        #endregion


        #region init
        public Configuration()
        {
            PrepareTemplate();

            PrepareTarget();
        }
        #endregion

        public bool PrepareTemplate()
        {
            bool valid = false;

            var ExcelApp = (Application)ExcelDnaUtil.Application;


            //---------------
            // init template path

            if ( !File.Exists( TemplatePath ) )
            {
                // create a temporary file from the embedded template
                TemplatePath = Path.Combine( Path.GetTempPath(), "summary.xlsx" );

                using ( var writer = new FileInfo( TemplatePath ).OpenWrite() )
                {
                    var bytes = Resources.template;
                    writer.Write( bytes, 0, bytes.Length );
                    writer.Close();
                }
            }



            //---------------
            // open template

            try
            {
                var wb = ExcelApp.Workbooks.Open( TemplatePath );
                valid = true;
            }
            catch
            {
                valid = false;
            }



            return IsTemplateValid = valid;
        }

        public bool PrepareTarget()
        {
            bool valid = false;

            if ( !Directory.Exists( TargetPath ) )
            {
                valid = true;
            }

            return IsTargetValid = valid;
        }
    }
}
