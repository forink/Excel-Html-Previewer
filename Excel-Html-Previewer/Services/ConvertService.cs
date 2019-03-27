using NPOI.HSSF.UserModel;
using NPOI.SS.Converter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Xml;
using Excel_Html_Previewer.ViewModels;
using Excel_Html_Previewer.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;
using System.Text;
using HtmlAgilityPack;
using System.Web.Mvc;
using Excel_Html_Previewer.Helper;

namespace Excel_Html_Previewer.Services
{
    public class ConvertService
    {
        /// <summary>
        /// 使用NPOI元件將Excel格式化為Html
        /// </summary>
        /// <param name="targetPath">目標資料夾路徑</param>
        /// <returns></returns>
        public List<ExcelSheetHtml> GetObjectFromExcelFile(string targetPath)
        {
            List<ExcelSheetHtml> SheetList = new List<ExcelSheetHtml>();
            IWorkbook Workbooks;

            using (var fs = new FileStream(targetPath, FileMode.Open, FileAccess.Read))
            {
                if (Path.GetExtension(targetPath).ToLower() == ".xls")
                {
                    Workbooks = new HSSFWorkbook(fs);
                }
                else
                {
                    Workbooks = new XSSFWorkbook(fs);
                }
            }

            ExcelToHtmlConverter ExcelConverter = new ExcelToHtmlConverter();

            // Set output parameters
            ExcelConverter.OutputColumnHeaders = true;
            ExcelConverter.OutputHiddenColumns = true;
            ExcelConverter.OutputHiddenRows = true;
            ExcelConverter.OutputLeadingSpacesAsNonBreaking = true;
            ExcelConverter.OutputRowNumbers = true;
            ExcelConverter.UseDivsToSpan = true;
            ExcelConverter.ProcessWorkbook(Workbooks);

            XmlNodeList XlsTitleList = ExcelConverter.Document.GetElementsByTagName("h2");
            XmlNodeList XlsTableList = ExcelConverter.Document.GetElementsByTagName("table");

            for (int i = 0; i < Workbooks.NumberOfSheets; i++)
            {
                SheetList.Add(new ExcelSheetHtml()
                {
                    Sn = i + 1,
                    TabTitle = XlsTitleList[i].InnerText,
                    Content = XlsTableList[i].OuterXml
                });
            }

            return SheetList;
        }

        /// <summary>
        /// 將Excel轉存的Html包格式化為可嵌入的Html
        /// </summary>
        /// <param name="targetPath">目標資料夾路徑</param>
        /// <returns></returns>
        public List<ExcelSheetHtml> GetObjectFromHtmlPack(string targetPath)
        {
            List<ExcelSheetHtml> SheetList = new List<ExcelSheetHtml>();
            string DirName = Path.GetFileName(targetPath);

            //先爬附件資料夾的tabstrip.htm以取得TAB清單
            string TabIndex = string.Format(@"{0}\\{1}.files\\tabstrip.htm", targetPath, DirName);
            var htmlDoc = new HtmlDocument();
            htmlDoc.Load(TabIndex);

            Dictionary<string, string> TabDicts = new Dictionary<string, string>();
            var tabTags = htmlDoc.DocumentNode.SelectSingleNode("//body").Descendants("a");

            int TagIndex = 0;
            foreach (var a in tabTags)
            {
                var tabFileName = a.Attributes["href"].Value;
                var tabTitle = a.InnerText;
                TabDicts.Add(tabTitle, tabFileName);

                var tabFilePath = string.Format(@"{0}\\{1}.files\\{2}", targetPath, DirName, tabFileName);
                var TabDoc = new HtmlDocument();
                TabDoc.Load(tabFilePath);

                SheetList.Add(new ExcelSheetHtml()
                {
                    Sn = TagIndex + 1,
                    TabTitle = tabTitle,
                    Content = TabDoc.DocumentNode.SelectSingleNode("//body").InnerHtml
                });

                TagIndex++;
            }

            return SheetList;
        }

        /// <summary>
        /// 取得Excel轉存的Html包中的CSS設定並加以處理
        /// </summary>
        /// <param name="targetPath">目標資料夾路徑</param>
        /// <returns></returns>
        public string GetNewHtmlPackCssContent(string targetPath)
        {
            //取得CSS的內容
            string CssContent = System.IO.File.ReadAllText(string.Format(@"{0}\\{1}.files\\stylesheet.css", targetPath, Path.GetFileName(targetPath)));
            MatchCollection MatchCss = Regex.Matches(CssContent, @"[^}]?([^{]*{[^}]*})", RegexOptions.Multiline);
            StringBuilder NewCssBuilder = new StringBuilder();
            foreach (Match m in MatchCss)
            {
                NewCssBuilder.Append("\n #viewer-frame>#viewer-box>#tab-box .sheet-box ").Append(m.Value);
            }
            return NewCssBuilder.ToString();
        }

        /// <summary>
        /// 取得目標資料夾要預覽內容的檔案或資料夾清單
        /// </summary>
        /// <param name="targetPath">目標資料夾路徑</param>
        /// <param name="searchPatten">尋檔規則</param>
        /// <param name="searchOption">搜尋選項</param>
        /// <param name="byDirectory">True:資料夾模式/False:檔案模式(Default)</param>
        /// <returns></returns>
        public List<SelectListItem> GetSelectList(string targetPath, string searchPatten, SearchOption searchOption, bool byDirectory = false)
        {
            List<SelectListItem> SelectList = new List<SelectListItem>();
            string FilePath = Regex.IsMatch(targetPath, @"^(?:[a-zA-Z]:(?:\\|/)|\\\\)") ? targetPath : HttpContext.Current.Server.MapPath(targetPath);

            List<string> TargetFileList = (byDirectory) ? Directory.GetDirectories(FilePath, searchPatten, searchOption).ToList()
                : Directory.GetFiles(FilePath, searchPatten, searchOption).ToList();
            Dictionary<string, string> TargetFileDict = TargetFileList.ToDictionary(x => Path.GetFileName(x), x => x);

            foreach (KeyValuePair<string, string> kv in TargetFileDict)
            {
                SelectListItem Item = new SelectListItem();
                Item.Text = kv.Key;
                Item.Value = kv.Value;
                SelectList.Add(Item);
            }

            return SelectList;
        }
    }
}