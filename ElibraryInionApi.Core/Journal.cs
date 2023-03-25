using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElibraryInionApi.Core
{
    public class Journal
    {
        public string TitleId { get; set; }
        public string Name { get; set; }
        public string InionCode { get; set; }
        public string SubjectCode { get; set; }
        public string Subject { get; set; }
        public string RubricCode { get; set; }
    }
}
