using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertXMLtoXLS.Logic
{
    public class CustomTuple : Tuple<string, string, string, string, string>
    {
      

        public CustomTuple(string item1, string item2, string item3, string item4, string item5) : base(item1, item2, item3, item4, item5)
        {
        }
    }
}
