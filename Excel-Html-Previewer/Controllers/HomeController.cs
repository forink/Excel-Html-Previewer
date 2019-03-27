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
            return View();
        }

        [AcceptVerbs(HttpVerbs.Get | HttpVerbs.Post)]
        public ActionResult PreviewerForExcel(FormCollection form)
        {
            VmPreviewerForExcel ViewModel = new VmPreviewerForExcel();
            string TargetFileName = (!string.IsNullOrEmpty(form["select-excel-file"])) ? form["select-excel-file"] : string.Empty;

            List<string> TargetFileList = Directory.GetFiles(Server.MapPath(GlobalVars.LOADING_FILES), "*.xls*",
            SearchOption.AllDirectories).ToList();
            Dictionary<string, string> TargetFileDict = TargetFileList.ToDictionary(x => Path.GetFileName(x), x => x);
            foreach (KeyValuePair<string, string> kv in TargetFileDict)
            {
                SelectListItem Item = new SelectListItem();
                Item.Text = kv.Key;
                Item.Value = kv.Key;
                Item.Selected = (kv.Key == TargetFileName);
                ViewModel.FileSelectList.Add(Item);
            }

            if (this.Request.RequestType == "POST" && !string.IsNullOrEmpty(TargetFileName))
            {
                ConvertService ConvertServ = new ConvertService();
                string XmlPath = TargetFileDict[TargetFileName];
                ViewModel.SheetList = ConvertServ.ExcelToHtmlByNPOI(XmlPath);
            }

            return View(ViewModel);
        }

        [AcceptVerbs(HttpVerbs.Get | HttpVerbs.Post)]
        public ActionResult PreviewerForHtmlPack(FormCollection form)
        {
            VmPreviewerForExcel ViewModel = new VmPreviewerForExcel();
            string TargetFileName = (!string.IsNullOrEmpty(form["select-excel-file"])) ? form["select-excel-file"] : string.Empty;

            List<string> TargetFileList = Directory.GetFiles(Server.MapPath(GlobalVars.LOADING_FILES), "*.xls*",
            SearchOption.AllDirectories).ToList();
            Dictionary<string, string> TargetFileDict = TargetFileList.ToDictionary(x => Path.GetFileName(x), x => x);
            foreach (KeyValuePair<string, string> kv in TargetFileDict)
            {
                SelectListItem Item = new SelectListItem();
                Item.Text = kv.Key;
                Item.Value = kv.Key;
                Item.Selected = (kv.Key == TargetFileName);
                ViewModel.FileSelectList.Add(Item);
            }

            if (this.Request.RequestType == "POST" && !string.IsNullOrEmpty(TargetFileName))
            {
                ConvertService ConvertServ = new ConvertService();
                string XmlPath = TargetFileDict[TargetFileName];
                ViewModel.SheetList = ConvertServ.ExcelToHtmlByNPOI(XmlPath);
            }

            return View(ViewModel);
        }
    }
}