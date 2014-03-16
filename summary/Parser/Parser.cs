using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelSummarizer.Parser
{
    internal class Parser
    {
        static Dictionary<InputDoc.Field, String> _fieldCells;

        public Parser()
        {
            _fieldCells = new Dictionary<InputDoc.Field, String>()
            {
                {InputDoc.Field.CustomerName, "A7" },
                {InputDoc.Field.Address, "A8" },
                {InputDoc.Field.Address1, "A9" },
                {InputDoc.Field.ID, "E15" },
                {InputDoc.Field.Date, "K15" }
            };
        }


        public InputDoc Parse( FileInfo file, CancellationTokenSource cts, Action<string> log )
        {
            InputDoc doc = null;

            using ( var pck = new ExcelPackage( file ) )
            {
                if ( !ValidateExcelPackage( pck ) ) throw new Exception( "Invalid Excel package!" );
                var wks = pck.Workbook.Worksheets.First();

                doc = new InputDoc();
                log( String.Format( "New bill( {0} )", file.FullName ) );

                object val = null;
                foreach ( var kvp in _fieldCells )
                {
                    val = wks.Cells[ kvp.Value ].Value;
                    if ( val != null )
                    {
                        doc.SetData( kvp.Key, val.ToString() );
                        //log( String.Format( "  {0}: {1}", kvp.Key.ToString(), val ) );
                    }
                    else
                    {
                        doc.SetData( kvp.Key, "#ERROR" );
                        log( String.Format( "\"{0}\" wasn't found in [{1}]", kvp.Key, kvp.Value ) );
                    }
                }

                if ( !cts.IsCancellationRequested )
                {
                    //------------
                    // look for Bill.Price
                    // Its legend cell should be in column 'H', between row 29 and 300
                    string col = "H";
                    int row, minRow = 30, maxRow = 300;
                    row = minRow;
                    ExcelRange cell = wks.Cells[ col + row ];
                    val = cell.Value;
                    string sVal = null;
                    while ( ( sVal == null || !sVal.Contains( "Rechnungsbetrag" ) )
                        && row < maxRow )
                    {
                        cell = wks.Cells[ col + ++row ];
                        sVal = cell.Value as String;
                    }
                    // then the price should be on its right
                    if ( sVal != null && sVal.Contains( "Rechnungsbetrag" ) )
                    {
                        cell = wks.Cells[ "M" + row ];
                        val = cell.Value;

                        doc.SetData( InputDoc.Field.Price, val );
                    }
                    else
                    {
                        doc.SetData( InputDoc.Field.Price, Double.NaN );

                        //log( String.Format( "\"{0}\" wasn't found in [M{1}-M{2}]", Bill.Field.Price, minRow, minRow + maxRow ) );
                        //log( "Cancelling parsing of " + file );
                        //return null;
                    }
                }
            }

            return doc;
        }

        private bool ValidateExcelPackage( ExcelPackage pck )
        {
            try
            {
                var wkb = pck.Workbook;

                if ( wkb.Worksheets.Count < 1 )
                    return false;


                var wks = wkb.Worksheets.First();

                if ( wks == null || wks.Dimension == null )
                    return false;
            }
            catch ( System.Runtime.InteropServices.COMException )
            {
                return false;
            }

            return true;
        }
    }
}
