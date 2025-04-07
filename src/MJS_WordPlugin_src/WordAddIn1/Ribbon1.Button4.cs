using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Drawing.Imaging;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Drawing;
using System.Xml;
using System.Threading;
using static System.Runtime.CompilerServices.RuntimeHelpers;
using System.Reflection.Emit;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            // Word.Document Doc = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument;
            // if (File.Exists(Path.GetDirectoryName(Doc.FullName) + "\\headerFile\\" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1") + @".txt"))
            // {
            //     return;
            // }
            //button4.Enabled = false;
            loader load = new loader();
            load.Visible = false;
            if (!makeBookInfo(load))
            {
                load.Close();
                load.Dispose();
                return;
            }


            MessageBox.Show("出力が終了しました。");

            button4.Enabled = true;
            button2.Enabled = true;
            button5.Enabled = true;
        }
    }
}
