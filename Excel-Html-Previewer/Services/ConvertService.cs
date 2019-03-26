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

namespace Excel_Html_Previewer.Services
{
    public class ConvertService
    {
        public List<ExcelSheetHtml> ExcelToHtmlByNPOI(string path)
        {
            List<ExcelSheetHtml> SheetList = new List<ExcelSheetHtml>();
            IWorkbook Workbooks;

            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                if (Path.GetExtension(path).ToLower() == ".xls")
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
    }
}