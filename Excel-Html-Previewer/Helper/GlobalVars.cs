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
        /// 讀檔來源路徑
        /// </summary>
        public static string LOADING_FILES { set; get; }

        /// <summary>
        /// 建構式
        /// </summary>
        static GlobalVars()
        {
            try
            {
                LOADING_FILES = ConfigurationManager.AppSettings["LOADING_FILES"].Trim();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}