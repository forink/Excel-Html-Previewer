using Excel_Html_Previewer.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Excel_Html_Previewer.ViewModels
{
    public class VM_ExcelPreviewer
    {
        public List<SelectListItem> XlsSelectList { get; set; }

        public List<ExcelSheetHtml> SheetList { set; get; }

        public VM_ExcelPreviewer()
        {
            XlsSelectList = new List<SelectListItem>();
            SheetList = new List<ExcelSheetHtml>();
        }
    }
}