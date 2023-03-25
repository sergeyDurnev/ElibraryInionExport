using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElibraryInionApi.Core
{
    public class Item
    {
        /// <summary>
        /// номер публикации
        /// </summary>
        public int Number { get; set; }
        public string ItemId { get; set; }
        public string ParentId { get; set; }
        public string VolumeNumber { get; set; }
        public string IssueNumber { get; set; }
        public string Year { get; set; }
        public bool Complete { get; set; }
        public IEnumerable<Field> Fields { get; set; }
    }
}
