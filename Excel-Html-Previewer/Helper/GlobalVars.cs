using System;
using System.Configuration;

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

        public static string TEST_LOGIN_USERID { set; get; }

        public static string TEST_LOGIN_PASSWORD { set; get; }

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
                TEST_LOGIN_USERID = ConfigurationManager.AppSettings["TEST_LOGIN_USERID"].Trim();
                TEST_LOGIN_PASSWORD = ConfigurationManager.AppSettings["TEST_LOGIN_PASSWORD"].Trim();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}