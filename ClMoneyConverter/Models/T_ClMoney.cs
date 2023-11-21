using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClMoneyConverter.Models
{
    internal class T_ClMoney
    {
        [Description("일자")]
        public string C_Date { get; set; }

        [Description("내역")]
        public string History { get; set; }

        [Description("금액")]
        public string Amount { get; set; }

        [Description("청구할인")]
        public string Discount { get; set; }

        [Description("할부")]
        public string Installment { get; set; }


        [Description("대분류")]
        public string MainCategory { get; set; }

        [Description("소분류")]
        public string Subcategory { get; set; }

        [Description("계좌")]
        public string Account { get; set; }

        [Description("구분")]
        public string Division { get; set; }

        [Description("비고")]
        public string Note { get; set; }
    }
}
