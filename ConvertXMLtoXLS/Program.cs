using ConvertXMLtoXLS.Logic;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertXMLtoXLS
{
    class Program
    {
        static void Main(string[] args)
        {
            FileReader xx = new FileReader();
            List<CustomTuple> table = xx.readDataFromXml();
            xx.SaveDataToExcel(table);
        }
        
           
        
    }
}
