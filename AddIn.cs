using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelSummarizer
{
    public class AddIn : IExcelAddIn
    {
        #region IExcelAddIn Members

        public void AutoOpen()
        {
            //throw new NotImplementedException();
        }

        public void AutoClose()
        {
            //throw new NotImplementedException();
        }

        #endregion
    }
}
