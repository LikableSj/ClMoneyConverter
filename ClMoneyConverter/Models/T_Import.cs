using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClMoneyConverter.Models
{
    internal class T_Import
    {
        [Description("날짜")]
        public string C_Date { get; set; }


        [Description("자산")]
        public string Account { get; set; }


        [Description("대분류")]
        public string MainCategory { get; set; }

        [Description("소분류")]
        public string Subcategory { get; set; }


        [Description("내용")]
        public string History { get; set; }

        [Description("금액")]
        public string Amount { get; set; }


        /// <summary>
        /// 수입/지출/이체출금
        /// </summary>
        [Description("수입/지출")]
        public string Category { get; set; }

        [Description("비고")]
        public string Note { get; set; }
    }
}
