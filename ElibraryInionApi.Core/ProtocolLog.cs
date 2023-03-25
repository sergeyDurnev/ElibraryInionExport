using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElibraryInionApi.Core
{
    public class ProtocolLog
    {
        public string Date { get; set; }
        public string SubjectCode { get; set; }
        public string Subject { get; set; }
        public string TitleId { get; set; }
        public string Name { get; set; }
        public string InionCode { get; set; }
        public string Year { get; set; }
        public string Volume { get; set; }
        public string Number { get; set; }
        public string ParentId { get; set; }
        public int ArticlesCount { get; set; }

        // Количество статей с аннотациями и ключевыми словами
        public int ArticlesCompleteCount { get; set; }

        //  Количество статей без аннотаций и ключевых слов
        public int ArticlesIncompleteCount { get; set; }

        // Рубрика ГРНТИ
        public string RubricCode { get; set; }
    }
}
