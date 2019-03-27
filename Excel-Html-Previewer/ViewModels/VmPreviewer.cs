using Excel_Html_Previewer.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Excel_Html_Previewer.ViewModels
{
    public abstract class VmPreviewer
    {
        public List<SelectListItem> FileSelectList { get; set; }

        public List<ExcelSheetHtml> SheetList { set; get; }

        public VmPreviewer()
        {
            FileSelectList = new List<SelectListItem>();
            SheetList = new List<ExcelSheetHtml>();
        }
    }

    public class VmPreviewerForExcel : VmPreviewer { }

    public class VmPreviewerForHtmlPack : VmPreviewer { }
}