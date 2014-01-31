using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSummarizer
{
    [System.Diagnostics.DebuggerStepThrough]
    public class EventArgs<T> : EventArgs
    {
        public T Value { get; protected set; }

        public EventArgs( T value )
        {
            this.Value = value;
        }
    }
}
