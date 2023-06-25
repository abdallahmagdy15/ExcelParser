using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser.Models
{
    public class DomainModelResultList<T>
    {
        public List<DomainModelResult<T>> ResultList { get; set; } = new List<DomainModelResult<T>>();
        public List<string> Errors { get; set; } = new List<string>();
    }

    public class DomainModelResult<T>
    {
        public T Result { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public int RowIndex { get; set; }
    }
}
