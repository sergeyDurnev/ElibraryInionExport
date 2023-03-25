using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using ElibraryInionApi.Core;

namespace ElibraryInionApi
{
    public class ElibraryApiConsole
    {
        //private readonly ILogger _logger;

        private const string APP_PATH = "http://elibrary.ru/projects/API-NEB/API_NEB.aspx";
        //private const string Code = "1f45bb59-f811-4b59-8401-476dcc5310b6";
        private const string Sid1 = "011";
        private const string Sid2 = "026";
        public void GetItemsData()
        {
            ILogger _logger = LogManager.GetLogger(GetType().Name);

            //GetTest();

            try
            {
                ApplicationConfiguration applicationConfiguration = new ApplicationConfiguration();

               // Console.WriteLine("Получение данных из ELIBRARY");

                IElibraryApi elibraryApi = new ElibraryApi();
                elibraryApi.GetItemIdsData(true, applicationConfiguration.ApiCode, applicationConfiguration.InputPath, applicationConfiguration.OutputPath, null, null, null, null, null);

                Console.WriteLine("Завершение работы");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Призошли ошибки получении данных, см. лог");

                Console.WriteLine("Призошли ошибки получении данных, см. лог, нажмите клавишу Enter для завершения работы");
                Console.ReadLine();
                Console.WriteLine("Завершение работы");
            }
        }
      
    }
}
