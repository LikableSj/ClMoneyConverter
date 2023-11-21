using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClMoneyConverter.Models
{
    internal class T_ChangAccount
    {
        [Description("변경전")]
        public string OldName { get; set; }
        [Description("변경후")]
        public string NewName { get; set; }
    }
}
