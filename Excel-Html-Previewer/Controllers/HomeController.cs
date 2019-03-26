using Excel_Html_Previewer.Models;
using Excel_Html_Previewer.Services;
using Excel_Html_Previewer.ViewModels;
using NPOI.HSSF.UserModel;
using NPOI.SS.Converter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel_Html_Previewer.Helper;

namespace Excel_Html_Previewer.Controllers
{
    public class HomeController : Controller
    {
        [AcceptVerbs(HttpVerbs.Get | HttpVerbs.Post)]
        public ActionResult Index(FormCollection form)
        {
            VM_ExcelPreviewer ViewModel = new VM_ExcelPreviewer();
            string XlsFileName = (!string.IsNullOrEmpty(form["select-excel-file"])) ? form["select-excel-file"] : string.Empty;

            List<string> XlsList = Directory.GetFiles(Server.MapPath(GlobalVars.LOADING_FILES), "*.xls*",
            SearchOption.AllDirectories).ToList();
            Dictionary<string, string> XlsDict = XlsList.ToDictionary(x => Path.GetFileName(x), x => x);
            foreach (KeyValuePair<string, string> kv in XlsDict)
            {
                SelectListItem Item = new SelectListItem();
                Item.Text = kv.Key;
                Item.Value = kv.Key;
                Item.Selected = (kv.Key == XlsFileName);
                ViewModel.XlsSelectList.Add(Item);
            }

            if (this.Request.RequestType == "POST" && !string.IsNullOrEmpty(XlsFileName))
            {
                ConvertService ConvertServ = new ConvertService();
                string XmlPath = XlsDict[XlsFileName];
                ViewModel.SheetList = ConvertServ.ExcelToHtmlByNPOI(XmlPath);
            }

            return View(ViewModel);
        }
    }
}