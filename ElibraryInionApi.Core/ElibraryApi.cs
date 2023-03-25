using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Xml.Linq;

namespace ElibraryInionApi.Core
{
    public class ElibraryApi : IElibraryApi
    {
        private const string APP_PATH = "http://elibrary.ru/projects/API-NEB/API_NEB.aspx";
        private const string Sid1 = "011";
        private const string Sid2 = "026";
        public string GetItemIdsData(bool consoleType, string apiCode, string inputPath, string outputPath, string parentId, string subjectCode, string titleId, string yearPublication, string rubricCode)
        {
            string app_path2 = APP_PATH + "?ucode=" + apiCode + "&sid=" + Sid1;

            string dateNow = DateTime.Now.Date.Year + "-" + DateTime.Now.Date.Month + "-" + DateTime.Now.Date.Day;

            string outputDirectory = outputPath;

            string logHistory = "log_history_" + dateNow + ".txt";
            string logHistoryFile = Path.Combine(outputDirectory, logHistory);

            string dateNowExcel = null;

            int dayExcel = DateTime.Now.Date.Day;
            if (dayExcel < 10)
            {
                dateNowExcel = "0" + DateTime.Now.Date.Day + ".";
            }
            else
            {
                dateNowExcel = DateTime.Now.Date.Day + ".";
            }

            int monthExcel = DateTime.Now.Date.Month;
            if (monthExcel < 10)
            {
                dateNowExcel = dateNowExcel + "0" + DateTime.Now.Date.Month + ".";
            }
            else
            {
                dateNowExcel = dateNowExcel + DateTime.Now.Date.Month + ".";
            }

            dateNowExcel += DateTime.Now.Date.Year;

            List<ProtocolLog> protocolLogs = new List<ProtocolLog>();

            string itemIdsData = null;

            if (consoleType)
            {

                Console.WriteLine("Получение данных из ELIBRARY");

                int.TryParse(File.ReadAllText(Path.Combine(inputPath, "lastItemId.txt")), out int lastItemId);

                File.WriteAllText(Path.Combine(inputPath, "lastItemId_old.txt"), lastItemId.ToString());
                int lastItemIdTemp = lastItemId;

                // получение типов статей, которые нужно выгрузить
                List<string> articleTypes = new List<string>();
                //Create COM Objects. Create a COM object for everything that is referenced
                string typeCodeFilePath = Path.Combine(inputPath, "article types.xlsx");
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(typeCodeFilePath);
                Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        //new line
                        if (j == 1)
                            Console.Write("\r\n");

                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            articleTypes.Add(xlRange.Cells[i, j].Value2.ToString());
                            Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                        }
                            
                    }
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);


                IEnumerable<string> files = Directory.GetFiles(inputPath).Where(x => !x.EndsWith(".txt") && !x.EndsWith("article types.xlsx"));

                string logProtocol = "log_protocol_" + dateNow + ".xls";
                string logProtocolFile = Path.Combine(outputDirectory, logProtocol);

                string itemLogProtocol = "log_protocol_items_" + dateNow + ".xls";
                string itemLogProtocolFile = Path.Combine(outputDirectory, itemLogProtocol);

                List<Item> logItems = new List<Item>();

                foreach (string file in files)
                {
                    FileInfo fileInfo = new FileInfo(file);
                    Console.WriteLine($"Получение списка TitleId из файла:{file}");
                    string subjectPath = Path.Combine(outputDirectory, fileInfo.Name.Replace(".xls", "").Replace(".xlsx", ""));
                    if (Directory.Exists(subjectPath) == false)
                    {
                        Directory.CreateDirectory(subjectPath);
                    }

                    string dateNowPath = Path.Combine(subjectPath, dateNow);

                    if (!Directory.Exists(dateNowPath))
                    {
                        Directory.CreateDirectory(dateNowPath);
                    }
                    else
                    {
                        Directory.Delete(dateNowPath, true);
                        Directory.CreateDirectory(dateNowPath);
                        if (File.Exists(logHistoryFile))
                        {
                            File.Delete(logHistoryFile);
                        }
                    }
                    IEnumerable<Journal> titles = GetTitles(file);

                    foreach (Journal title in titles)
                    {
                        //string number = null;
                        Console.WriteLine($"Получаем список ItemId по TitleId:{title.TitleId}");
                        string path1 = app_path2 + "&titleid=" + title.TitleId;
                        if (string.IsNullOrEmpty(title.RubricCode) == false)
                        {
                            //string rubricCode = title.RubricCode.Replace(".", "").Trim();
                            path1 = path1 + "&mrc=" + title.RubricCode.Replace(".", "").Trim();
                        }
                        DateTime dateTime = DateTime.Now;
                        //int yearPubl = 2022;
                        //int yearPubl = dateTime.Year;
                        //if (dateTime.Month == 1)
                        //{
                        //    yearPubl--;
                        //}

                        //path1 = path1 + "&yearpubl=" + yearPubl;

                        if (lastItemId > 0)
                        {
                            path1 = path1 + "&lastid=" + lastItemId;
                        }

                        WriteHistory(logHistoryFile, path1);
                        itemIdsData = GetApiData(path1);
                        XDocument xDoc = XDocument.Parse(itemIdsData);
                        IEnumerable<XElement> itemIds = xDoc.Root.Descendants("item");
                        List<Item> items = new List<Item>();
                        int itemOrder = 0;
                        foreach (XElement itemId in itemIds)
                        {
                            string typeCode = itemId.Attribute("typeCode").Value.Trim();
                            if (articleTypes.Contains(typeCode) || articleTypes.Count() == 0)
                            {
                                itemOrder++;
                                int.TryParse(itemId.Value, out int itemIdValue);
                                if (itemIdValue > lastItemId && itemIdValue > lastItemIdTemp)
                                {
                                    lastItemIdTemp = itemIdValue;
                                }

                                string path2 = APP_PATH + "?ucode=" + apiCode + "&sid=" + Sid2 + "&itemid=" + itemId.Value;
                                WriteHistory(logHistoryFile, path2);

                                string destinationPath = Path.Combine(dateNowPath, title.TitleId);
                                if (!Directory.Exists(destinationPath))
                                {
                                    Directory.CreateDirectory(destinationPath);
                                }

                                Item item = GetItem(apiCode, itemOrder, itemId.Value, title.SubjectCode, null, destinationPath, typeCode);
                                items.Add(item);

                                if (!protocolLogs.Exists(x => x.Number?.Trim() == item.IssueNumber?.Trim()
                                   && x.Volume?.Trim() == item.VolumeNumber?.Trim()
                                   && x.TitleId == title.TitleId
                                   && x.Year == item.Year))
                                {

                                    if (item.Complete)
                                    {
                                        protocolLogs.Add(new ProtocolLog
                                        {
                                            Date = dateNowExcel,
                                            SubjectCode = title.SubjectCode,
                                            Subject = title.Subject,
                                            TitleId = title.TitleId,
                                            Name = title.Name,
                                            InionCode = title.InionCode,
                                            Year = item.Year,
                                            Volume = item.VolumeNumber,
                                            Number = item.IssueNumber,
                                            ParentId = item.ParentId,
                                            ArticlesCount = 1,
                                            ArticlesCompleteCount = 1,
                                            ArticlesIncompleteCount = 0,
                                            RubricCode = title.RubricCode
                                        });
                                    }
                                    else
                                    {
                                        protocolLogs.Add(new ProtocolLog
                                        {
                                            Date = dateNowExcel,
                                            SubjectCode = title.SubjectCode,
                                            Subject = title.Subject,
                                            TitleId = title.TitleId,
                                            Name = title.Name,
                                            InionCode = title.InionCode,
                                            Year = item.Year,
                                            Volume = item.VolumeNumber,
                                            Number = item.IssueNumber,
                                            ParentId = item.ParentId,
                                            ArticlesCount = 1,
                                            ArticlesCompleteCount = 0,
                                            ArticlesIncompleteCount = 1,
                                            RubricCode = title.RubricCode
                                        });
                                    }

                                }
                                else
                                {
                                    //if (string.IsNullOrEmpty(item.VolumeNumber) == false)
                                    //{
                                    var index = protocolLogs.FindIndex(x => x.Number?.Trim() == item.IssueNumber?.Trim() && x.Volume == item.VolumeNumber?.Trim()
                            && x.TitleId == title.TitleId
                            && x.Year == item.Year);
                                    protocolLogs[index].ArticlesCount = protocolLogs[index].ArticlesCount + 1;
                                    if (item.Complete)
                                    {
                                        protocolLogs[index].ArticlesCompleteCount = protocolLogs[index].ArticlesCompleteCount + 1;
                                    }
                                    else
                                    {
                                        protocolLogs[index].ArticlesIncompleteCount = protocolLogs[index].ArticlesIncompleteCount + 1;
                                    }
                                    //}
                                    //    else
                                    //    {
                                    //        var index = protocolLogs.FindIndex(x => x.Number?.Trim() == item.IssueNumber?.Trim()
                                    //&& x.TitleId == title.TitleId
                                    //&& x.Year == item.Year);
                                    //        protocolLogs[index].ArticlesCount = protocolLogs[index].ArticlesCount + 1;
                                    //        if (item.Complete)
                                    //        {
                                    //            protocolLogs[index].ArticlesCompleteCount = protocolLogs[index].ArticlesCompleteCount + 1;
                                    //        }
                                    //        else
                                    //        {
                                    //            protocolLogs[index].ArticlesIncompleteCount = protocolLogs[index].ArticlesIncompleteCount + 1;
                                    //        }
                                    //    }

                                }

                            }
                        }

                        if (items.Count() > 0)
                        {
                            string destinationPath = Path.Combine(dateNowPath, title.TitleId);
                            if (!Directory.Exists(destinationPath))
                            {
                                Directory.CreateDirectory(destinationPath);
                            }

                            var groupedItems = items.GroupBy(p => p.ParentId);
                            foreach (var groupedItem in groupedItems)
                            {
                                string journalName = title.Name;
                                if (journalName.Length >= 45)
                                {
                                    journalName = journalName.Substring(0, 45);
                                }

                                journalName = journalName.Replace(@"\", "").Replace("/", "").Replace("|", "")
                                    .Replace(":", "").Replace("*", "").Replace("?", "").Replace(";", "").Replace("\"", "");

                                string destinationIssuePath = null;
                                if (string.IsNullOrEmpty(groupedItem.FirstOrDefault().VolumeNumber))
                                {
                                    destinationIssuePath = Path.Combine(destinationPath,
                                    title.SubjectCode + "_" + journalName + "_" + groupedItem.Key + "_"
                                    + groupedItem.FirstOrDefault()?.IssueNumber + "_"
                                    + groupedItem.Count() + ".rtf");
                                }
                                else
                                {
                                    destinationIssuePath = Path.Combine(destinationPath,
                                    title.SubjectCode + "_" + journalName + "_" + groupedItem.Key + "_Том_"
                                    + groupedItem.FirstOrDefault()?.VolumeNumber + "_Выпуск_" + groupedItem.FirstOrDefault()?.IssueNumber + "_"
                                    + groupedItem.Count() + ".rtf");
                                }                               
                                GenerateTxtDocument(groupedItem, destinationIssuePath);
                            }

                            logItems.AddRange(items);
                        }
                    }
                }

                if (protocolLogs.Count() > 0)
                {
                    GenerateExcelLogDocument(logProtocolFile, apiCode, protocolLogs);
                    GenerateExcelItemsLog(itemLogProtocolFile, apiCode, protocolLogs, logItems);
                }


                if (lastItemIdTemp > lastItemId)
                {
                    File.WriteAllText(Path.Combine(inputPath, "lastItemId.txt"), lastItemIdTemp.ToString());
                    Console.WriteLine($"Записываем последний ItemId в LastItemId:{lastItemId}");
                }
            }
            else
            {
                string path1 = null;
                if (parentId != null)
                {
                    path1 = app_path2 + "&parentid=" + parentId;
                }
                else
                {
                    path1 = app_path2 + "&titleid=" + titleId + "&yearpubl=" + yearPublication;

                }

                if (rubricCode != null)
                {
                    path1 = path1 + "&mrc=" + rubricCode;
                }

                itemIdsData = GetApiData(path1);
            }

            return itemIdsData;
        }

        public Item GetItem(string apiCode,
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
            )
        {

            string app_path3 = APP_PATH + "?ucode=" + apiCode + "&sid=" + Sid2 + "&itemid=";
            //List<Item> items = new List<Item>();

            // перенести в консольное приложение
            //int.TryParse(itemId.Value, out int itemIdValue);
            //if (itemIdValue > lastItemId && itemIdValue > lastItemIdTemp)
            //{
            //    lastItemIdTemp = itemIdValue;
            //}

            string path2 = app_path3 + itemId;
            // WriteHistory(logHistoryFile, path2);
            string xml = GetApiData(path2);
            XDocument xDoc2 = XDocument.Parse(xml);

            if (!Directory.Exists(destinationPath))
            {
                Directory.CreateDirectory(destinationPath);
            }

            //string destinationFilePath = Path.Combine(destinationPath, itemId + ".xml");

            int abstractCount = xDoc2.Root.Descendants("abstract").Count();
            int keywordCount = xDoc2.Root.Descendants("keyword").Count();

            // определяем полный комплект
            Console.WriteLine($"abstractCount={abstractCount} , keywordCount={keywordCount}");
            Console.WriteLine(xDoc2.ToString());
            bool hasComplete = false;
            if (abstractCount > 0 && keywordCount > 0)
            {
                hasComplete = true;
            }
            else
            {
                if (xDoc2.Root.Element("item").Descendants("abstract").Count() == 0)
                {
                    XElement abstractsElement = new XElement("abstracts",
                  new XElement("abstract",
                  new XAttribute("lang", "RU"),
                  "нет")
                  );
                    xDoc2.Root.Element("item").Add(abstractsElement);
                    xml = xDoc2.ToString();
                }
            }

            if (typeCode == null)
            {
                typeCode = xDoc2.Root.Element("item").Attribute("typecode").Value;
            }
            List<Field> fields = new List<Field> { new Field { FieldName = "Тип_", FieldValue = typeCode} };
           


            string authors = null;

            string language = xDoc2.Root.Element("item").Attribute("lang").Value;

            IEnumerable<XElement> authorsElement = from author in xDoc2.Root.Element("item").Descendants("author")
                                                   where author.Attribute("lang").Value == language
                                                   select author;

            XElement authorElement = authorsElement.FirstOrDefault();

            if (authorElement == null)
            {
                authorElement = xDoc2.Root.Element("item").Descendants("author").FirstOrDefault();
            }

            if (authorElement != null)
            {
                string lastname = authorElement.Element("lastname")?.Value;
                string initials = authorElement.Element("initials")?.Value;
                string initialsNorm = null;
                if (!string.IsNullOrEmpty(initials))
                {
                    initials = Regex.Replace(initials, @"\b(оглы|оглу|кызы|кизи|оол|ugli|kizi|oglu|ogli)\b", @"", RegexOptions.IgnoreCase);
                    initials = initials.Trim();
                    if (initials.Contains(".") && !initials.Contains(" "))
                    {
                        initialsNorm = initials.Replace(" ", "");
                    }
                    else
                    {
                        RegexOptions options = RegexOptions.None;
                        Regex regex = new Regex("[ ]{2,}", options);
                        initials = regex.Replace(initials, " ");

                        string[] initialList = initials.Split(' ');
                        foreach (string initial in initialList)
                        {
                            initialsNorm = initialsNorm + initial.Substring(0, 1) + ".";
                        }
                    }

                    authors = lastname + ", " + initialsNorm;
                }
                else
                {
                    authors = lastname;
                }
            }


            XElement refElement = xDoc2.Root.Element("item").Element("ref");

            if (refElement != null)
            {
                string refValue = HttpUtility.HtmlDecode(refElement.Value).Trim();
                string authorsNorm = null;
                if (authors != null)
                {
                    authorsNorm = "<b>" + authors + "</b>";
                }

                if (authorsElement.Count() <= 3)
                {
                    var regex1 = new Regex(@"\S+, \S+");
                    refValue = regex1.Replace(HttpUtility.HtmlDecode(refElement.Value).Trim(), "", 1);
                    fields.Add(new Field { FieldName = "10000 ", FieldValue = authorsNorm + "%24500 " + refValue.Trim() });
                }
                else
                {
                    fields.Add(new Field { FieldName = "24500 ", FieldValue = refValue.Trim() });
                }

            }


            //if (refElement != null)
            //{
            //    string refValue = HttpUtility.HtmlDecode(refElement.Value).Trim();
            //    string authorsNorm = "<b>" + authors + "</b>";

            //    int pos = refValue.IndexOf(authors);
            //    string refNorm = null;
            //    if (pos >= 0)
            //    {
            //        refNorm = refValue.Substring(pos + authors.Length).Trim();
            //    }

            //    refNorm = "%24500 " + refNorm;

            //}

            if (language == "RU")
            {
                fields.Add(new Field { FieldName = "О4000 ", FieldValue = "570" });
            }
            else if (language == "EN")
            {
                fields.Add(new Field { FieldName = "О4000 ", FieldValue = "045" });
            }
            else if (language == "IT")
            {
                fields.Add(new Field { FieldName = "О4000 ", FieldValue = "235" });
            }
            else if (language == "DE")
            {
                fields.Add(new Field { FieldName = "О4000 ", FieldValue = "481" });
            }
            else if (language == "FR")
            {
                fields.Add(new Field { FieldName = "О4000 ", FieldValue = "745" });
            }
            else
            {
                fields.Add(new Field { FieldName = "О4000 ", FieldValue = "000" });
            }

            XElement abstractElement = xDoc2.Root.Element("item").Descendants("abstract").FirstOrDefault(x => x.Attribute("lang").Value == "RU");
            if (abstractElement == null)
            {
                abstractElement = xDoc2.Root.Element("item").Descendants("abstract").First();
            }
            string abstractValue = "Аннотация_" + HttpUtility.HtmlDecode(abstractElement.Value);
            fields.Add(new Field { FieldName = "82000\r", FieldValue = abstractValue });

            fields.Add(new Field { FieldName = "$_" + subjectCode });

            IEnumerable<XElement> keywordsElement = from keywordElement in xDoc2.Root.Element("item").Descendants("keyword")
                                                    where keywordElement.Attribute("lang").Value == "RU"
                                                    select keywordElement;
            if (keywordsElement.Count() == 0)
            {
                keywordsElement = xDoc2.Root.Element("item").Descendants("keyword");
            }
            foreach (XElement keywordElement in keywordsElement)
            {
                string keyword = HttpUtility.HtmlDecode(keywordElement.Value);
                fields.Add(new Field { FieldName = "$", FieldValue = keyword });
            }

            XElement itemCodesElement = xDoc2.Root.Element("item").Element("itemCodes");
            if (itemCodesElement != null)
            {
                XElement doiElement = (from code in itemCodesElement.Elements("code")
                                       where code.Attribute("type").Value == "DOI"
                                       select code).FirstOrDefault();
                if (doiElement != null)
                {
                    string doi = "DOI " + HttpUtility.HtmlDecode(doiElement.Value);
                    fields.Add(new Field { FieldName = "85000 ", FieldValue = doi });
                }
            }

            XElement linkElement = xDoc2.Root.Element("item").Descendants("link").FirstOrDefault();
            if (linkElement != null)
            {
                string link = HttpUtility.HtmlDecode(linkElement.Value);
                fields.Add(new Field { FieldName = "85100 ", FieldValue = link });
            }

            parentId = xDoc2.Root.Element("item")?.Attribute("parentId")?.Value?.Trim();

            string issueNumber = xDoc2.Root.Descendants("number").FirstOrDefault()?.Value.Trim();
            string volumeNumber = xDoc2.Root.Descendants("volume").FirstOrDefault()?.Value.Trim();
            string year = xDoc2.Root.Descendants("year").FirstOrDefault()?.Value.Trim();

            return new Item
            {
                Number = itemOrder,
                ItemId = itemId,
                ParentId = parentId,
                VolumeNumber = volumeNumber,
                IssueNumber = issueNumber,
                Year = year,
                Complete = hasComplete,
                Fields = fields
            };
        }

        // обращаемся по маршруту api/values 
        public string GetApiData(string path)
        {
            var client = new HttpClient();

            var response = client.GetAsync(path).Result;
            return response.Content.ReadAsStringAsync().Result;

        }

        public IEnumerable<Journal> GetTitles(string filePath)
        {
            List<Journal> titles = new List<Journal>();

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = excel.Workbooks.Open(filePath);
            Worksheet excelSheet = wb.ActiveSheet;
            int i = 2;
            bool isTitleId = true;
            while (isTitleId)
            {
                string titleId = excelSheet.Cells[i, 1].Value?.ToString();
                string name = excelSheet.Cells[i, 2].Value?.ToString();
                string inionCode = excelSheet.Cells[i, 3].Value?.ToString();
                string subjectCode = excelSheet.Cells[i, 4].Value?.ToString();
                string subject = excelSheet.Cells[i, 5].Value?.ToString();
                string rubricCode = excelSheet.Cells[i, 6].Value?.ToString();
                i++;
                if (!string.IsNullOrEmpty(titleId))
                {

                    titles.Add(new Journal
                    {
                        TitleId = titleId,
                        Name = name,
                        InionCode = inionCode,
                        SubjectCode = subjectCode,
                        Subject = subject,
                        RubricCode = rubricCode
                    });
                }
                else
                {
                    isTitleId = false;
                }
            }

            wb.Close();
            return titles;
        }

        //private void GenerateWordDocument(IEnumerable<Item> items, string destinationPath)
        //{

        //    object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

        //    Microsoft.Office.Interop.Word.Application winword =
        //      new Microsoft.Office.Interop.Word.Application();

        //    winword.Visible = false;

        //    //Заголовок документа
        //    //winword.Documents.Application.Caption = "www.CSharpCoderR.com";

        //    object missing = System.Reflection.Missing.Value;

        //    //Создание нового документа
        //    Microsoft.Office.Interop.Word.Document document =
        //        winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

        //    foreach (Item item in items)
        //    {
        //        Microsoft.Office.Interop.Word.Paragraph oPara6;
        //        object oRng2 = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
        //        oPara6 = document.Content.Paragraphs.Add(ref oRng2);
        //        oPara6.Range.Text = "#" + item.Number;
        //        oPara6.Range.Font.Bold = 1;
        //        oPara6.Range.InsertParagraphAfter();

        //        foreach (Field field in item.Fields)
        //        {

        //            Microsoft.Office.Interop.Word.Paragraph oPara1;
        //            object oRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
        //            oPara1 = document.Content.Paragraphs.Add(ref oRng);

        //            oPara1.Range.Text = field.FieldName + field.FieldValue;
        //            oPara1.Range.Font.Bold = 0;
        //            oPara1.Range.InsertParagraphAfter();

        //        }
        //    }

        //    //Microsoft.Office.Interop.Word.Paragraph oPara1;
        //    //oPara1 = document.Content.Paragraphs.Add(ref missing);
        //    //oPara1.Range.Text = "#1";
        //    //oPara1.Range.Font.Bold = 0;
        //    //oPara1.Range.InsertParagraphAfter();

        //    //foreach (Field field in fields)
        //    //{              

        //    //    Microsoft.Office.Interop.Word.Paragraph oPara6;
        //    //    object oRng2 = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
        //    //    oPara6 = document.Content.Paragraphs.Add(ref oRng2);

        //    //    oPara6.Range.Text = field.FieldName + field.FieldValue;               
        //    //    oPara6.Range.InsertParagraphAfter();

        //    //}

        //    //string textWord = null;

        //    //string recordWord = null;
        //    //foreach (Field field in fields)
        //    //{
        //    //    string text = null;

        //    //    if (field.FieldName != "$")
        //    //    {
        //    //        text = field.FieldName + " " + field.FieldValue;
        //    //    }
        //    //    else
        //    //    {
        //    //        text = field.FieldName + field.FieldValue;
        //    //    }

        //    //    if (recordWord == null)
        //    //    {
        //    //        recordWord = text;
        //    //    }
        //    //    else
        //    //    {
        //    //        recordWord = recordWord + "\r" + text;
        //    //    }

        //    //}
        //    ////recordWord = recordWord + "\r";
        //    //textWord += recordWord;

        //    //Microsoft.Office.Interop.Word.Paragraph oPara3;
        //    //object oRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
        //    //oPara3 = document.Content.Paragraphs.Add(ref oRng);
        //    //// oPara3.Range.Text = textWord;
        //    ////oPara3.Range.Font.Bold = 0;
        //    //oPara3.Range.InsertParagraphAfter();

        //    Object begin = Type.Missing;
        //    Object end = Type.Missing;
        //    Microsoft.Office.Interop.Word.Range wordrange = document.Range(ref begin, ref end);
        //    wordrange.Select();

        //    //wordrange.Font.Size = 10;
        //    ////wordrange.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed;
        //    //wordrange.Font.Name = "Times New Roman";

        //    ////document.Content.Font.Size = 10;
        //    ////document.Content.Font.Bold = 0;
        //    ////document.Content.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
        //    ////document.Content.ParagraphFormat.Alignment =
        //    ////Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
        //    //document.Content.ParagraphFormat.LeftIndent =
        //    // document.Content.Application.CentimetersToPoints((float)0);
        //    //document.Content.ParagraphFormat.RightIndent =
        //    // document.Content.Application.CentimetersToPoints((float)0);

        //    string s = string.Empty;
        //    for (int i = 1; i < document.Paragraphs.Count; i++)
        //    {
        //        s = document.Paragraphs[i].Range.Text;
        //        //Object begin1 = 42;
        //        //Object end1 = 49;
        //        //wordrange = document.Range(ref begin1, ref end1);
        //        //wordrange.Select();
        //        //На Рис.5. Слева выведенный  текст на данном этапе
        //        //Меняем характеристики текста выделенного фрагмента
        //        wordrange.Font.Size = 10;
        //        document.Paragraphs[i].SpaceAfter = (float)0.5;
        //        //wordrange.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed;
        //        //wordrange.Text = "Текст который мы выводим в выделенный участок ";
        //    }


        //    //Microsoft.Office.Interop.Word.Paragraph oPara4;
        //    //oRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
        //    //oPara4 = document.Content.Paragraphs.Add(ref oRng);
        //    //oPara4.Range.Text = "This is a sentence of normal text. Now here is a table:";
        //    //oPara4.Range.Font.Bold = 0;
        //    //oPara4.Range.InsertParagraphAfter();

        //    // document.Content.Text = textWord;


        //    //winword.Visible = true;

        //    //Сохранение документа          

        //    FileInfo fullFileNameInfo = new FileInfo(destinationPath);
        //    object filename = destinationPath;
        //    document.SaveAs(ref filename);

        //    //Закрытие текущего документа
        //    document.Close(ref missing, ref missing, ref missing);
        //    document = null;
        //    //Закрытие приложения Word
        //    winword.Quit(ref missing, ref missing, ref missing);
        //    winword = null;

        //}

        public void GenerateTxtDocument(IEnumerable<Item> items, string destinationPath)
        {
            using (StreamWriter sw = new StreamWriter(destinationPath, true, Encoding.UTF8))
            {
                foreach (Item item in items)
                {
                    string recordNumber = "#" + item.Number;
                    sw.WriteLine(recordNumber);
                    //File.WriteAllText(fullFileName, recordNumber, Encoding.Unicode);

                    foreach (Field field in item.Fields)
                    {
                        string text = field.FieldName + field.FieldValue;
                        //File.WriteAllText(fullFileName, text, Encoding.Unicode);
                        sw.WriteLine(text);

                    }
                    //string endRecord = "%%%%%" + Environment.NewLine;
                    //File.WriteAllText(fullFileName, endRecord, Encoding.Unicode);
                    //sw.WriteLine(endRecord);                    
                }
            }
        }

        public void GenerateExcelLogDocument(string logPath, string apiKey, IEnumerable<ProtocolLog> historyLogs)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            object misvalue = System.Reflection.Missing.Value;

            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;

            //Get a new workbook.
            oWB = oXL.Workbooks.Add();
            oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;

            //Add table headers going cell by cell.
            oSheet.Cells[1, 1] = "ДАТА";
            oSheet.Cells[1, 2] = "Код темы";
            oSheet.Cells[1, 3] = "Тема ИНИОН";
            oSheet.Cells[1, 4] = "TitleId";
            oSheet.Cells[1, 5] = "Название журнала";
            oSheet.Cells[1, 6] = "Код ИНИОН";
            oSheet.Cells[1, 7] = "Год";
            oSheet.Cells[1, 8] = "Номер";
            oSheet.Cells[1, 9] = "Идентификатор выпуска в НЭБ";
            oSheet.Cells[1, 10] = "Количество статей (всего)";
            oSheet.Cells[1, 11] = "Количество статей с аннотациями и ключевыми словами";
            oSheet.Cells[1, 12] = "Количество статей без аннотаций и ключевых слов";
            oSheet.Cells[1, 13] = "Код ГРНТИ";

            //Format A1:D1 as bold, vertical alignment = center.
            oSheet.get_Range("A1", "M1").Font.Bold = true;
            oSheet.get_Range("A1", "M1").VerticalAlignment =
                Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            int i = 2;
            foreach (ProtocolLog historyLog in historyLogs)
            {
                oSheet.Cells[i, 1] = historyLog.Date;
                oSheet.Cells[i, 2] = historyLog.SubjectCode;
                oSheet.Cells[i, 3] = historyLog.Subject;
                oSheet.Cells[i, 4] = historyLog.TitleId;
                oSheet.Cells[i, 5] = historyLog.Name;
                oSheet.Cells[i, 6] = historyLog.InionCode;
                oSheet.Cells[i, 7] = historyLog.Year;
                if (string.IsNullOrEmpty(historyLog.Volume) == false)
                {
                    oSheet.Cells[i, 8] = "Т." + historyLog.Volume + " №" + historyLog.Number;
                }
                else
                {
                    oSheet.Cells[i, 8] = historyLog.Number;
                }
                oSheet.Cells[i, 9] = historyLog.ParentId;
                oSheet.Cells[i, 10] = historyLog.ArticlesCount;
                oSheet.Cells[i, 11] = historyLog.ArticlesCompleteCount;
                oSheet.Cells[i, 12] = historyLog.ArticlesIncompleteCount;
                oSheet.Cells[i, 13] = historyLog.RubricCode;
                i++;
            }


            oXL.Visible = false;
            oXL.UserControl = false;
            // oXL.DisplayAlerts = false;
            //oXL.ScreenUpdating = false;
            //oXL.Interactive = false;           

            if (!System.IO.File.Exists(logPath))
            {
                logPath = logPath.Replace("/", "");
                oWB.SaveAs(logPath);
            }
            else
            {
                logPath = logPath.Replace("/", "").Replace(".xls", "") + "(Copy).xls";
                oWB.SaveAs(logPath);
            }


            oWB.Close();
            oXL.Quit();
        }


        public void GenerateExcelItemsLog(string logPath, string apiKey, IEnumerable<ProtocolLog> historyLogs, IEnumerable<Item> items)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            object misvalue = System.Reflection.Missing.Value;

            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;

            //Get a new workbook.
            oWB = oXL.Workbooks.Add();
            oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;

            //Add table headers going cell by cell.
            oSheet.Cells[1, 1] = "ДАТА";
            oSheet.Cells[1, 2] = "Код темы";
            oSheet.Cells[1, 3] = "Тема ИНИОН";
            oSheet.Cells[1, 4] = "TitleId";
            oSheet.Cells[1, 5] = "Название журнала";
            oSheet.Cells[1, 6] = "Код ИНИОН";
            oSheet.Cells[1, 7] = "Год";
            oSheet.Cells[1, 8] = "Номер";
            oSheet.Cells[1, 9] = "Идентификатор публикации в НЭБ";
            oSheet.Cells[1, 10] = "Код ГРНТИ";

            //Format A1:D1 as bold, vertical alignment = center.
            oSheet.get_Range("A1", "M1").Font.Bold = true;
            oSheet.get_Range("A1", "M1").VerticalAlignment =
                Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            int i = 2;
            foreach (Item item in items)
            {
                ProtocolLog historyLog = historyLogs.FirstOrDefault(x => x.ParentId == item.ParentId);
                oSheet.Cells[i, 1] = historyLog.Date;
                oSheet.Cells[i, 2] = historyLog.SubjectCode;
                oSheet.Cells[i, 3] = historyLog.Subject;
                oSheet.Cells[i, 4] = historyLog.TitleId;
                oSheet.Cells[i, 5] = historyLog.Name;
                oSheet.Cells[i, 6] = historyLog.InionCode;
                oSheet.Cells[i, 7] = historyLog.Year;
                if (string.IsNullOrEmpty(item.VolumeNumber) == false)
                {
                    oSheet.Cells[i, 8] = "Т." + item.VolumeNumber + " №" + historyLog.Number;
                }
                else
                {
                    oSheet.Cells[i, 8] = historyLog.Number;
                }
                oSheet.Cells[i, 9] = item.ItemId;
                oSheet.Columns[10].NumberFormat = "@";
                oSheet.Cells[i, 10] = historyLog.RubricCode;
                i++;
            }


            oXL.Visible = false;
            oXL.UserControl = false;
            // oXL.DisplayAlerts = false;
            //oXL.ScreenUpdating = false;
            //oXL.Interactive = false;           

            if (!System.IO.File.Exists(logPath))
            {
                logPath = logPath.Replace("/", "");
                oWB.SaveAs(logPath);
            }
            else
            {
                logPath = logPath.Replace("/", "").Replace(".xls", "") + "(Copy).xls";
                oWB.SaveAs(logPath);
            }


            oWB.Close();
            oXL.Quit();
        }

        public void WriteHistory(string writePath, string text)
        {
            using (StreamWriter sw = new StreamWriter(writePath, true, Encoding.Default))
            {
                sw.WriteLine(text);
            }
        }

        private Dictionary<string, string> BuildSubjects()
        {
            Dictionary<string, string> subjects = new Dictionary<string, string>
            {
                { "$_A02", "Философия" },
                { "$_A03", "История" },
                { "$_A04", "Социология" },
                { "$_A05", "Экономика" },
                { "$_A10", "Правоведение" },
                { "$_A11", "Политология" },
                { "$_A12", "Науковедение" },
                { "$_A16", "Языкознание" },
                { "$_A17", "Литературоведение" },
                { "$_A21", "Религиоведение" }
            };

            return subjects;
        }

    }
}
