using Microsoft.Office.Interop.Excel;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ElibraryInionApi
{
    class Program
    {     

        static void Main(string[] args)
        {
            ElibraryApiConsole elibraryApi = new ElibraryApiConsole();
            elibraryApi.GetItemsData();                   
        }

       
    }


}
