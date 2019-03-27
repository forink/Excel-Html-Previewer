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
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Text;

namespace Excel_Html_Previewer.Controllers
{
    public class HomeController : Controller
    {
        /// <summary>
        /// 首頁
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        [AcceptVerbs(HttpVerbs.Get | HttpVerbs.Post)]
        public ActionResult Index(FormCollection form)
        {
            return View();
        }

        /// <summary>
        /// 讀取XLS與XLSX格式並以HTML方式顯示
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        [AcceptVerbs(HttpVerbs.Get | HttpVerbs.Post)]
        public ActionResult PreviewerForExcel(FormCollection form)
        {
            VmPreviewerForExcel ViewModel = new VmPreviewerForExcel();
            ConvertService ConvertServ = new ConvertService();

            //掃描目標資料夾並取出列表放在下拉式選單中
            string SelectedItem = (!string.IsNullOrEmpty(form["select-file-or-directory"])) ? Path.GetFileName(form["select-file-or-directory"]) : string.Empty;
            ViewModel.FileSelectList = ConvertServ.GetSelectList(GlobalVars.LOADING_FILES_EXCEL, "*.xls*", SearchOption.AllDirectories);

            //取得索引並轉換目標檔案
            if (Request.RequestType == "POST" && !string.IsNullOrEmpty(SelectedItem))
            {
                ViewModel.FileSelectList.Find(x => x.Text == SelectedItem).Selected = true;

                string ConvertPath = ViewModel.FileSelectList.Single(x => x.Text == SelectedItem).Value;
                ViewModel.SheetList = ConvertServ.GetObjectFromExcelFile(ConvertPath);
            }

            return View(ViewModel);
        }

        /// <summary>
        /// 讀取純HTML包並無損顯示於模擬框架中
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        [AcceptVerbs(HttpVerbs.Get | HttpVerbs.Post)]
        public ActionResult PreviewerForHtmlPack(FormCollection form)
        {
            VmPreviewerForHtmlPack ViewModel = new VmPreviewerForHtmlPack();
            ConvertService ConvertServ = new ConvertService();

            //掃描目標資料夾並取出列表放在下拉式選單中
            string SelectedItem = (!string.IsNullOrEmpty(form["select-file-or-directory"])) ? Path.GetFileName(form["select-file-or-directory"]) : string.Empty;
            ViewModel.FileSelectList = ConvertServ.GetSelectList(GlobalVars.LOADING_FILES_HTMLPACK, "*", SearchOption.TopDirectoryOnly, true);

            //取得索引並轉換目標檔案
            if (Request.RequestType == "POST" && !string.IsNullOrEmpty(SelectedItem))
            {
                ViewModel.FileSelectList.Find(x => x.Text == SelectedItem).Selected = true;

                string ConvertPath = ViewModel.FileSelectList.Single(x => x.Text == SelectedItem).Value;
                ViewModel.CssContent = ConvertServ.GetNewHtmlPackCssContent(ConvertPath);
                ViewModel.SheetList = ConvertServ.GetObjectFromHtmlPack(ConvertPath);
            }

            return View(ViewModel);
        }
    }
}