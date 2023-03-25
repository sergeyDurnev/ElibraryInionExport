using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElibraryInionApi
{
    class ApplicationConfiguration
    {
        public string ApiCode { get; }
        public string GetItemCode { get; }
        public string InputPath { get; }
        public string OutputPath { get; }
        public int? ParentId { get; }
        public string Subject { get; }

        public ApplicationConfiguration()
        {
            ApiCode = (ConfigurationManager.AppSettings["ApiCode"] ?? "").Trim();
            GetItemCode = (ConfigurationManager.AppSettings["GetItemCode"] ?? "").Trim();
            InputPath = IfAbsolutePath((ConfigurationManager.AppSettings["inputPath"] ?? "").Trim());
            OutputPath = IfAbsolutePath((ConfigurationManager.AppSettings["outputPath"] ?? "").Trim());
            int.TryParse(ConfigurationManager.AppSettings["ParentId"], out int parentId);
            ParentId = parentId;
            Subject = (ConfigurationManager.AppSettings["Subject"] ?? "").Trim();

            if (string.IsNullOrWhiteSpace(InputPath) || !Directory.Exists(InputPath))
                throw new InvalidOperationException("Неверно указана директория \"InputPath\"");
            if (string.IsNullOrWhiteSpace(OutputPath) || !Directory.Exists(OutputPath))
                throw new InvalidOperationException("Неверно указана директория \"OutputPath\"");
        }

        private static string IfAbsolutePath(string path)
        {
            return path.Replace("|BaseDirectory|", AppDomain.CurrentDomain.BaseDirectory);
        }

    }
}
