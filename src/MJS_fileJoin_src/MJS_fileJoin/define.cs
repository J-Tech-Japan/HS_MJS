using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MJS_fileJoin
{
    partial class MainForm
    {
        // searchBase.jsの内容を取得するプロパティ
        private static string SearchBaseJs
        {
            get
            {
                string searchBaseJsPath = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    "Resources",
                    "searchBase.js"
                );

                return File.ReadAllText(searchBaseJsPath, Encoding.UTF8);
            }
        }

        // 動的に構築されるsearchJs
        public static string searchJs => @"var searchWords = $('♪');" + "\n" + SearchBaseJs;
    }
}