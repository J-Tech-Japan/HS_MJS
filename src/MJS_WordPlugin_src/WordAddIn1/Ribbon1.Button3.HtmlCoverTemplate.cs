using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private static string BuildHtmlCoverTemplate1(bool cover)
        {
            string htmlCoverTemplate1 = "";
            htmlCoverTemplate1 += @"<!DOCTYPE HTML>" + "\n";
            htmlCoverTemplate1 += @"<html>" + "\n";
            htmlCoverTemplate1 += @"<head>" + "\n";
            htmlCoverTemplate1 += @"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" + "\n";
            htmlCoverTemplate1 += @" <meta name=""generator"" content=""Adobe RoboHelp 2017"" />" + "\n";
            htmlCoverTemplate1 += @"<title>表紙</title>" + "\n";
            htmlCoverTemplate1 += @"<link rel=""stylesheet"" href=""cover.css"" type=""text/css"" />" + "\n";
            htmlCoverTemplate1 += @"<link rel=""stylesheet"" href=""font.css"" type=""text/css"" />" + "\n";
            htmlCoverTemplate1 += @"<link rel=""StyleSheet"" href=""resp.css"" type=""text/css"" />" + "\n";
            htmlCoverTemplate1 += @"<style type=""text/css"">" + "\n";
            htmlCoverTemplate1 += @"<!--" + "\n";
            htmlCoverTemplate1 += @"A:visited { color:#954F72; }" + "\n";
            htmlCoverTemplate1 += @"A:link { color:#000000; }" + "\n";
            htmlCoverTemplate1 += @"-->" + "\n";
            htmlCoverTemplate1 += @"</style>" + "\n";
            htmlCoverTemplate1 += @"<script type=""text/javascript"" language=""JavaScript"">" + "\n";
            htmlCoverTemplate1 += @"//<![CDATA[" + "\n";
            htmlCoverTemplate1 += @"function reDo() {" + "\n";
            htmlCoverTemplate1 += @"  if (innerWidth != origWidth || innerHeight != origHeight)" + "\n";
            htmlCoverTemplate1 += @"     location.reload();" + "\n";
            htmlCoverTemplate1 += @"}" + "\n";
            htmlCoverTemplate1 += @"if ((parseInt(navigator.appVersion) == 4) && (navigator.appName == ""Netscape"")) {" + "\n";
            htmlCoverTemplate1 += @"   origWidth = innerWidth;" + "\n";
            htmlCoverTemplate1 += @"   origHeight = innerHeight;" + "\n";
            htmlCoverTemplate1 += @"   onresize = reDo;" + "\n";
            htmlCoverTemplate1 += @"}" + "\n";
            htmlCoverTemplate1 += @"onerror = null;" + "\n";
            htmlCoverTemplate1 += @"//]]>" + "\n";
            htmlCoverTemplate1 += @"</script>" + "\n";
            htmlCoverTemplate1 += @"<style type=""text/css"">" + "\n";
            htmlCoverTemplate1 += @"<!--" + "\n";
            htmlCoverTemplate1 += @"div.WebHelpPopupMenu { position:absolute;" + "\n";
            htmlCoverTemplate1 += @"left:0px;" + "\n";
            htmlCoverTemplate1 += @"top:0px;" + "\n";
            htmlCoverTemplate1 += @"z-index:4;" + "\n";
            htmlCoverTemplate1 += @"visibility:hidden; }" + "\n";
            htmlCoverTemplate1 += @"-->" + "\n";
            if (cover)
            {
                htmlCoverTemplate1 += "\n";
                htmlCoverTemplate1 += @"p.HyousiLogo {" + "\n";
                htmlCoverTemplate1 += @"text-align       : center;" + "\n";
                htmlCoverTemplate1 += @"margin-top       : 60pt;" + "\n";
                htmlCoverTemplate1 += @"margin-bottom    : 40pt;" + "\n";
                htmlCoverTemplate1 += @"margin-right     : 0mm;" + "\n";
                htmlCoverTemplate1 += @"line-height      : 15pt;" + "\n";
                htmlCoverTemplate1 += @"}" + "\n";
                htmlCoverTemplate1 += "\n";
                htmlCoverTemplate1 += @"div.HyousiBackground {" + "\n";
                htmlCoverTemplate1 += @"display : table;" + "\n";
                htmlCoverTemplate1 += @"width   : 100%;" + "\n";
                htmlCoverTemplate1 += @"height  : 65px;" + "\n";
                htmlCoverTemplate1 += @"}" + "\n";
                htmlCoverTemplate1 += "\n";
                htmlCoverTemplate1 += @"p.HyousiText {" + "\n";
                htmlCoverTemplate1 += @"display             : table-cell;" + "\n";
                htmlCoverTemplate1 += @"background-image    : url('pict/hyousi.png');" + "\n";
                htmlCoverTemplate1 += @"background-repeat   : no-repeat;" + "\n";
                htmlCoverTemplate1 += @"background-position : center;" + "\n";
                htmlCoverTemplate1 += @"text-align          : center;" + "\n";
                htmlCoverTemplate1 += @"vertical-align      : middle;" + "\n";
                htmlCoverTemplate1 += @"font-size           : 1.8em;" + "\n";
                htmlCoverTemplate1 += @"font-weight         : bold;" + "\n";
                htmlCoverTemplate1 += @"color               : #FFF;" + "\n";
                htmlCoverTemplate1 += @"letter-spacing      : 10px;" + "\n";
                htmlCoverTemplate1 += @"}" + "\n";
            }
            htmlCoverTemplate1 += @"</style>" + "\n";
            htmlCoverTemplate1 += @"</head>" + "\n";

            return htmlCoverTemplate1;
        }

        public void BuildEdgeTrackerCoverTemplate(
            System.Reflection.Assembly assembly,
            string rootPath,
            string exportDir,
            string manualTitle,
            string trademarkTitle,
            List<string> trademarkTextList,
            string trademarkRight,
            ref string htmlCoverTemplate1)
        {
            string[] hyousiGazo = { "EdgeTracker_logo50mm.png", "MJS_LOGO_255.gif", "hyousi.png" };
            foreach (var hyousi in hyousiGazo)
            {
                Bitmap bmp = new Bitmap(assembly.GetManifestResourceStream("WordAddIn1.Resources." + hyousi));
                bmp.Save(rootPath + "\\" + exportDir + "\\pict\\" + hyousi);
            }
            htmlCoverTemplate1 += @"<body>" + "\n";
            htmlCoverTemplate1 += @"<p class=""HyousiLogo""><img style=""border: currentColor; border-image: none; width: 100%; max-width: 553px;"" alt="""" src=""pict/EdgeTracker_logo50mm.png"" border=""0""></p>" + "\n";
            htmlCoverTemplate1 += @"<div class=""HyousiBackground"">" + "\n";
            htmlCoverTemplate1 += @"<p class=""HyousiText"">" + manualTitle + "</p>\n";
            htmlCoverTemplate1 += @"</div>" + "\n";
            htmlCoverTemplate1 += @"<div class=""product_trademarks"">" + "\n";
            htmlCoverTemplate1 += string.Format(@"  <p class=""trademark_title"">{0}</p>" + "\n", trademarkTitle);
            foreach (string trademarkText in trademarkTextList)
            {
                htmlCoverTemplate1 += string.Format(@"  <p class=""trademark_text"">{0}</p>" + "\n", trademarkText);
            }
            htmlCoverTemplate1 += string.Format(@"  <p class=""trademark_right"">{0}</p>" + "\n", trademarkRight);
            htmlCoverTemplate1 += @"</div>" + "\n";
            htmlCoverTemplate1 += @"<p style=""text-align: center; margin-top: 80pt;""><a href=""https://www.mjs.co.jp"" target=""_blank""><img style=""border: currentColor; border-image: none; width: 100%; max-width: 255px;"" alt="""" src=""pict/MJS_LOGO_255.gif"" border=""0""></a></p>" + "\n";
        }

        public void BuildEasyCloudCoverTemplate(
            string rootPath,
            string exportDir,
            string manualTitle,
            string manualSubTitle,
            string manualVersion,
            string trademarkTitle,
            List<string> trademarkTextList,
            string trademarkRight,
            string subTitle,
            ref string htmlCoverTemplate1,
            ref string htmlCoverTemplate2)
        {
            if (File.Exists(rootPath + "\\" + exportDir + "\\template\\images\\cover-background.png"))
                htmlCoverTemplate1 += @"<body style=""text-justify-trim: punctuation; background-image: url('template/images/cover-background.png');background-repeat: no-repeat; background-position: 0px 300px;"">" + "\n";
            else
                htmlCoverTemplate1 += @"<body>" + "\n";

            htmlCoverTemplate1 += @"<p class=""manual_title"" style=""line-height: 130%;"">" + manualTitle + "</p>" + "\n";
            htmlCoverTemplate1 += @"<p class=""manual_subtitle"">" + manualSubTitle + "</p>" + "\n";

            if (File.Exists(rootPath + "\\" + exportDir + "\\template\\images\\cover-4.png"))
                htmlCoverTemplate1 += @"<p class=""manual_title"" style=""margin: 80px 0px 80px 100px; ""><img src=""template/images/cover-4.png"" width=""650"" /></p>" + "\n";
            else
                htmlCoverTemplate1 += @"<p class=""manual_title"" style=""margin: 80px 0px 80px 100px; ""></p>" + "\n";

            htmlCoverTemplate1 += @"<p class=""manual_version"">" + manualVersion + "</p>" + "\n";
            htmlCoverTemplate1 += @"<div class=""product_trademarks"">" + "\n";
            htmlCoverTemplate1 += string.Format(@"  <p class=""trademark_title"">{0}</p>" + "\n", trademarkTitle);
            foreach (string trademarkText in trademarkTextList)
            {
                htmlCoverTemplate1 += string.Format(@"  <p class=""trademark_text"">{0}</p>" + "\n", trademarkText);
            }
            htmlCoverTemplate1 += string.Format(@"  <p class=""trademark_right"">{0}</p>" + "\n", trademarkRight);
            htmlCoverTemplate1 += @"</div>" + "\n";

            if (!String.IsNullOrEmpty(subTitle))
            {
                htmlCoverTemplate2 += @"<p style=""margin-left: 700px; margin-top: 150px; font-size: 15pt; font-family: メイリオ;" + "\n";
                htmlCoverTemplate2 += @"    font-weight: bold;"">" + subTitle + "</p>" + "\n";
                htmlCoverTemplate2 += @"<p><a href=""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）""" + "\n";
                htmlCoverTemplate2 += @"                                        style=""margin-left: 700px; margin-top: 10px;""" + "\n";
                htmlCoverTemplate2 += @"                                        width=""132"" height=""48"" /></a>" + "\n";
            }
            else
            {
                htmlCoverTemplate2 += @"<p><a href=""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）""" + "\n";
                htmlCoverTemplate2 += @"                                        style=""margin-left: 700px; margin-top: 100px;""" + "\n";
                htmlCoverTemplate2 += @"                                        width=""132"" height=""48"" /></a>" + "\n";
            }
            htmlCoverTemplate2 += @" </p>" + "\n";
        }

        public void BuildPattern1CoverTemplate(
            string manualTitle,
            string manualTitleCenter,
            string manualSubTitle,
            string manualSubTitleCenter,
            string trademarkTitle,
            List<string> trademarkTextList,
            string trademarkRight,
            ref string htmlCoverTemplate2)
        {
            htmlCoverTemplate2 += string.Format(@"<p class=""manual_title"" style=""line-height: 130%; "">{0}</p>" + "\n",
                !string.IsNullOrWhiteSpace(manualTitle) ? manualTitle : manualTitleCenter);
            htmlCoverTemplate2 += string.Format(@"<p class=""manual_subtitle"">{0}</p>" + "\n",
                !string.IsNullOrWhiteSpace(manualSubTitle) ? manualSubTitle : manualSubTitleCenter);
            htmlCoverTemplate2 += @"<p class=""product_logo_main_nosub"">" + "\n";
            htmlCoverTemplate2 += @"  <img src = ""template/images/product_logo_main.png"" alt=""製品ロゴ（メイン）"">" + "\n";
            htmlCoverTemplate2 += @"</p>" + "\n";
            htmlCoverTemplate2 += @"<div class=""product_trademarks"">" + "\n";
            htmlCoverTemplate2 += string.Format(@"  <p class=""trademark_title"">{0}</p>" + "\n", trademarkTitle);
            foreach (string trademarkText in trademarkTextList)
            {
                htmlCoverTemplate2 += string.Format(@"  <p class=""trademark_text"">{0}</p>" + "\n", trademarkText);
            }
            htmlCoverTemplate2 += string.Format(@"  <p class=""trademark_right"">{0}</p>" + "\n", trademarkRight);
            htmlCoverTemplate2 += @"</div>" + "\n";
            htmlCoverTemplate2 += @"<p><a href = ""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）"" style=""margin-left: 700px; margin-top: 100px;"" width=""132"" height=""48"" /></a>" + "\n";
            htmlCoverTemplate2 += @"</p>" + "\n";
        }

        public void BuildPattern2CoverTemplate(
            List<List<string>> productSubLogoGroups,
            string manualTitleCenter,
            string manualTitle,
            string manualSubTitleCenter,
            string manualSubTitle,
            string manualVersionCenter,
            string manualVersion,
            string trademarkTitle,
            List<string> trademarkTextList,
            string trademarkRight,
            ref string htmlCoverTemplate2)
        {
            htmlCoverTemplate2 += @"<p class=""product_logo_main"">" + "\n";
            htmlCoverTemplate2 += @"  <img src = ""template/images/product_logo_main.png"" alt=""製品ロゴ（メイン）"">" + "\n";
            htmlCoverTemplate2 += @"</p>" + "\n";
            htmlCoverTemplate2 += @"<div class=""product_logo_sub"">" + "\n";
            foreach (List<string> subLogoGroup in productSubLogoGroups)
            {
                htmlCoverTemplate2 += @"<div>" + "\n";
                foreach (string subLogoFileName in subLogoGroup)
                {
                    htmlCoverTemplate2 += string.Format(@"  <img src = ""template/images/{0}"" alt=""製品ロゴ（サブ）"">" + "\n", subLogoFileName);
                }
                htmlCoverTemplate2 += @"</div>" + "\n";
            }
            htmlCoverTemplate2 += @"</div>" + "\n";
            htmlCoverTemplate2 += string.Format(@"<p class=""manual_title_center"" style=""line-height: 130%; "">{0}</p>" + "\n",
                !string.IsNullOrWhiteSpace(manualTitleCenter) ? manualTitleCenter : manualTitle);
            htmlCoverTemplate2 += string.Format(@"<p class=""manual_subtitle_center"">{0}</p>" + "\n",
                !string.IsNullOrWhiteSpace(manualSubTitleCenter) ? manualSubTitleCenter : manualSubTitle);
            htmlCoverTemplate2 += string.Format(@"<p class=""manual_version_center"">{0}</p>" + "\n",
                !string.IsNullOrWhiteSpace(manualVersionCenter) ? manualVersionCenter : manualVersion);
            htmlCoverTemplate2 += @"<div class=""product_trademarks"">" + "\n";
            htmlCoverTemplate2 += string.Format(@"  <p class=""trademark_title"">{0}</p>" + "\n", trademarkTitle);
            foreach (string trademarkText in trademarkTextList)
            {
                htmlCoverTemplate2 += string.Format(@"  <p class=""trademark_text"">{0}</p>" + "\n", trademarkText);
            }
            htmlCoverTemplate2 += string.Format(@"  <p class=""trademark_right"">{0}</p>" + "\n", trademarkRight);
            htmlCoverTemplate2 += @"</div>" + "\n";
            htmlCoverTemplate2 += @"<p><a href = ""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）"" style=""margin-left: 700px; margin-top: 100px;"" width=""132"" height=""48"" /></a>" + "\n";
            htmlCoverTemplate2 += @"</p>" + "\n";
        }

        // すべてのパターンに共通するHTMLテンプレートの追加
        public void AppendHtmlCoverTemplate2(ref string htmlCoverTemplate2)
        {
            htmlCoverTemplate2 += @"<script type=""text/javascript"" language=""javascript1.2"">//<![CDATA[" + "\n";
            htmlCoverTemplate2 += @"<!--" + "\n";
            htmlCoverTemplate2 += @"if (window.writeIntopicBar)" + "\n";
            htmlCoverTemplate2 += @"   writeIntopicBar(0);" + "\n";
            htmlCoverTemplate2 += @"//-->" + "\n";
            htmlCoverTemplate2 += @"//]]></script>" + "\n";
            htmlCoverTemplate2 += @"</body>" + "\n";
            htmlCoverTemplate2 += @"</html>" + "\n";
        }

    }
}
