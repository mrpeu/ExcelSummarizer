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

namespace ExcelSummarizer
{
    public class AddIn : IExcelAddIn
    {
        #region var
        static Configuration Configuration;

        static Ribbon ribbon;
        #endregion

        #region IExcelAddIn Members

        public void AutoOpen()
        {
            //Configuration = new Configuration();
        }

        public void AutoClose()
        {
            //throw new NotImplementedException();
        }

        internal static void RegisterRibbon( Ribbon ribbon )
        {
            //AddIn.ribbon = ribbon;

            //ribbon.UpdateConfiguration( Configuration );

            //ribbon.TemplatePathChanged += ( s, e ) => { Configuration.TemplatePath = e.Value; };
            //ribbon.TargetPathChanged += ( s, e ) => { Configuration.TargetPath = e.Value; };
        }
        #endregion
    }
}
