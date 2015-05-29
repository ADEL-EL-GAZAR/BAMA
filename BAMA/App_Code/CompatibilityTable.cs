using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BAMA.App_Code
{
    public class CompatibilityTable
    {
        public string prefix;
        public string stem;
        public string suffix;

        public CompatibilityTable(string prefix, string stem, string suffix)
        {
            this.prefix = prefix;
            this.stem = stem;
            this.suffix = suffix;
        }
    }
}
