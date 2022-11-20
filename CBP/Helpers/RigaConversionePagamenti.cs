using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CBP.Helpers
{
    public class RigaConversionePagamenti
    {
        public string Insurer { get; set; }
        public string ProductCode { get; set; }
        public string SiglaProdotto { get; set; }
        public string Warranty { get; set; }
        public int Tipologia { get; set; }

        public RigaConversionePagamenti() { }
    }
}
