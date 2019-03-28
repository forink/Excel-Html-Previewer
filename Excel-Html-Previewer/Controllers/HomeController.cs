using Excel_Html_Previewer.Helper;
using Excel_Html_Previewer.Services;
using Excel_Html_Previewer.ViewModels;
using System;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;

namespace Excel_Html_Previewer.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Login()
        {
            return View();
        }

        /// <summary>
        /// 用於測試用的極簡登入驗證
        /// </summary>
        /// <param name="userId"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Login(string userId, string password)
        {
            if (!userId.Equals(GlobalVars.TEST_LOGIN_USERID) || !password.Equals(GlobalVars.TEST_LOGIN_PASSWORD))
            {
                TempData["Error"] = "您輸入的帳號不存在或者密碼錯誤!";
                return View();
            }

            // 登入時清空所有 Session 資料
            Session.RemoveAll();

            FormsAuthenticationTicket ticket = new FormsAuthenticationTicket(1,
              "test", DateTime.Now, DateTime.Now.AddMinutes(30), false, "",
              FormsAuthentication.FormsCookiePath);

            string encTicket = FormsAuthentication.Encrypt(ticket);

            Response.Cookies.Add(new HttpCookie(FormsAuthentication.FormsCookieName, encTicket));

            return RedirectToAction("Index", "Home");

        }

        /// <summary>
        /// 首頁
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        [Authorize]
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
        [Authorize]
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
        [Authorize]
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