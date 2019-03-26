using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Excel_Html_Previewer.Models
{
    public class ExcelSheetHtml
    {
        public int Sn { set; get; }

        public string TabTitle { set; get; }

        public string Content { set; get; }

        public ExcelSheetHtml()
        {
            Sn = 0;
            TabTitle = string.Empty;
            Content = string.Empty;
        }

    }
}