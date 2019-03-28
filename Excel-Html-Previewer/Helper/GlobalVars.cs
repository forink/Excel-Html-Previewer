using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace Excel_Html_Previewer.Helper
{
    public static class GlobalVars
    {
        /// <summary>
        /// 讀檔來源路徑 (Excel Files)
        /// </summary>
        public static string LOADING_FILES_EXCEL { set; get; }

        /// <summary>
        /// 讀檔來源路徑 (Html Pack Files)
        /// </summary>
        public static string LOADING_FILES_HTMLPACK { set; get; }

        /// <summary>
        /// CSS 覆寫Pattern
        /// </summary>
        public static string CSS_OVERRIDE_PATTERN { set; get; }

        /// <summary>
        /// 建構式
        /// </summary>
        static GlobalVars()
        {
            try
            {
                LOADING_FILES_EXCEL = ConfigurationManager.AppSettings["LOADING_FILES_EXCEL"].Trim();
                LOADING_FILES_HTMLPACK = ConfigurationManager.AppSettings["LOADING_FILES_HTMLPACK"].Trim();
                CSS_OVERRIDE_PATTERN = ConfigurationManager.AppSettings["CSS_OVERRIDE_PATTERN"].Trim();

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}