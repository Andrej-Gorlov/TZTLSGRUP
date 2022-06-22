using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TZTLSGRUP
{
    internal class FinExample
    {
        public int Id { get; set; }
        public string Product { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public DateOnly Date { get; set; }
        public decimal Profit { get; set; }
    }
}
