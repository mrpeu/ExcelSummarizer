using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Threading;

namespace ExcelSummarizer.Parser
{
    public class Summarist
    {
        #region summarize
        public static List<InputDoc> Summarize( string path, CancellationTokenSource cts, Action<string> log, Action<int, int> progress )
        {
            List<InputDoc> bills = null;

            if ( ( File.GetAttributes( path ) & FileAttributes.Directory ) == FileAttributes.Directory )
            {
                bills = Summarize( new DirectoryInfo( path ), cts, log, progress );
            }

            return bills;
        }

        public static List<InputDoc> Summarize( DirectoryInfo dir, CancellationTokenSource cts, Action<string> log, Action<int, int> progress )
        {
            if ( cts.IsCancellationRequested ) return null;

            var parser = new Parser();

            var list = new List<InputDoc>();

            var extensions = new string[] { ".xls", ".xlsx", ".xlsm" };
            var files = Directory.EnumerateFiles( dir.FullName, "*", SearchOption.TopDirectoryOnly )
                .Where( file =>
                {
                    if ( extensions.Contains( Path.GetExtension( file ) ) )
                    {
                        return true;
                    }
                    else
                    {
                        log( "File ignored (wrong extension): " + file );
                        return false;
                    }
                }
            );

            int max = files.Count();
            progress( 0, max );
            InputDoc bill;
            string path;
            int i = 0;
            var enumerator = files.GetEnumerator();
            while ( enumerator.MoveNext() )
            {
                log( "//-------------" );

                if ( cts.IsCancellationRequested ) return null;

                path = enumerator.Current;
                bill = parser.Parse( new FileInfo( path ), cts, log );

                if ( bill != null )
                {
                    list.Add( bill );
                }
                i++;
                progress( i, max );
            }

            return list;
        }

        #endregion

        #region Save List<Bill> to output
        internal static bool CreateSummary( List<InputDoc> bills, CancellationTokenSource cts, Action<string> log, Action<int, int> progress )
        {
            //============
            // create new file
            ExcelPackage pck=null;

            //============
            // find headers and map them to our Bill.Field
            var fields = (InputDoc.Field[])Enum.GetValues( typeof( InputDoc.Field ) );
            var sFields = fields.Select( f => f.ToString() );
            var headers = pck.Workbook.Names
                .Where( c => sFields.Contains( c.Name ) )
                .ToDictionary( c => Enum.Parse( typeof( InputDoc.Field ), c.Name )
            );

            //============
            // pour in
            var sheet = pck.Workbook.Worksheets[ "Zusammenfassung" ];
            int i=1;
            object data;
            ExcelRangeBase cell = null;
            foreach ( InputDoc bill in bills )
            {
                if ( cts.IsCancellationRequested )
                {
                    pck.Dispose();
                    return false;
                }

                foreach ( InputDoc.Field field in fields )
                {
                    //todo: delete
                    if ( !headers.ContainsKey( field ) )
                        System.Windows.Forms.MessageBox.Show( "Check your header " + field, "Summarist", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation );
                    else
                    {
                        cell = headers[ field ].Offset( i, 0 );
                        data = bill.GetData( field );

                        switch ( field )
                        {
                            case InputDoc.Field.Date:
                            case InputDoc.Field.Price:
                                try { cell.Value = double.Parse( data.ToString() ); }
                                catch { cell.Value = data; }
                                break;

                            case InputDoc.Field.ID:
                                try { cell.Value = ushort.Parse( data.ToString() ); }
                                catch { cell.Value = data; }
                                break;

                            case InputDoc.Field.CustomerName:
                            case InputDoc.Field.Address:
                            case InputDoc.Field.Address1:
                            default:
                                cell.Value = data;
                                break;
                        }
                    }
                }
                progress( i, bills.Count );
                i++;
            }

            //============
            // save created file

            pck.Save();
            pck.Dispose();

            return true;
        }
        #endregion
    }
}
