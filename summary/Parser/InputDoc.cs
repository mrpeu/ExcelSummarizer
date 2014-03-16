using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSummarizer.Parser
{
    public class InputDoc
    {
        public enum Field { CustomerName, Address, Address1, ID, Date, Price }

        Dictionary<Field, object> Data;

        public InputDoc()
        {
            Data = new Dictionary<Field, object>();
            foreach ( var field in (Field[])Enum.GetValues( typeof( Field ) ) ) { Data.Add( field, null ); }
        }

        public object GetData( Field field )
        {
            return Data[ field ];
        }

        public object SetData( Field field, object value )
        {
            return Data[ field ] = value;
        }
    }
}
