using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElibraryInionApi.Core
{
    public interface IElibraryApi
    {
        string GetItemIdsData(bool consoleType, 
            string apiCode, 
            string inputPath, 
            string outputPath, 
            string parentId, 
            string subjectCode, 
            string titleId, 
            string yearPubl, string rubricCode);

        Item GetItem(string apiCode,
            int itemOrder,
            string itemId,
            //int? lastItemId, int lastItemIdTemp, 
            string subjectCode,
            //string subject, string inionCode, 
            string parentId,
            //string titleId, 
            string destinationPath,
            string typeCode
            //string logHistoryFile 
            //string dateNowExcel
            );

        string GetApiData(string path);

        IEnumerable<Journal> GetTitles(string filePath);
        void GenerateTxtDocument(IEnumerable<Item> items, string destinationPath);

        void GenerateExcelItemsLog(string logPath, string apiKey, IEnumerable<ProtocolLog> historyLogs, IEnumerable<Item> items);
        void WriteHistory(string writePath, string text);
    }
}
