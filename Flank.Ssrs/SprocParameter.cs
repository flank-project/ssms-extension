using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Flank.Ssrs
{
    public sealed class SprocParameter
    {
        public string Name { get; set; }
        public string SqlType { get; set; }
        public short MaxLength { get; set; }
        public byte Precision { get; set; }
        public byte Scale { get; set; }
        public bool IsOutput { get; set; }
    }
}
