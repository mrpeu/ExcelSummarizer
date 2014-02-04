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

            ribbon.TemplatePathChanged += ( s, e ) => { Configuration.TemplatePath = e.Value; };
            ribbon.TargetPathChanged += ( s, e ) => { Configuration.TargetPath = e.Value; };
        }
        #endregion
    }
}
