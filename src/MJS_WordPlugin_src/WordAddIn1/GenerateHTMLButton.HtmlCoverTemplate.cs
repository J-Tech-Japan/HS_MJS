// GenerateHTMLButton.HtmlCoverTemplate.cs
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private string BuildHtmlCoverTemplate1Header()
        {
            var sb = new StringBuilder();
            sb.AppendLine(@"<!DOCTYPE HTML>");
            sb.AppendLine(@"<html>");
            sb.AppendLine(@"<head>");
            sb.AppendLine(@"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />");
            sb.AppendLine(@" <meta name=""generator"" content=""Adobe RoboHelp 2017"" />");
            sb.AppendLine(@"<title>表紙</title>");
            sb.AppendLine(@"<link rel=""stylesheet"" href=""cover.css"" type=""text/css"" />");
            sb.AppendLine(@"<link rel=""stylesheet"" href=""font.css"" type=""text/css"" />");
            sb.AppendLine(@"<link rel=""StyleSheet"" href=""resp.css"" type=""text/css"" />");
            sb.AppendLine(@"<style type=""text/css"">");
            sb.AppendLine(@"<!--");
            sb.AppendLine(@"A:visited { color:#954F72; }");
            sb.AppendLine(@"A:link { color:#000000; }");
            sb.AppendLine(@"-->");
            sb.AppendLine(@"</style>");
            sb.AppendLine(@"<script type=""text/javascript"" language=""JavaScript"">");
            sb.AppendLine(@"//<![CDATA[");
            sb.AppendLine(@"function reDo() {");
            sb.AppendLine(@"  if (innerWidth != origWidth || innerHeight != origHeight)");
            sb.AppendLine(@"     location.reload();");
            sb.AppendLine(@"}");
            sb.AppendLine(@"if ((parseInt(navigator.appVersion) == 4) && (navigator.appName == ""Netscape"")) {");
            sb.AppendLine(@"   origWidth = innerWidth;");
            sb.AppendLine(@"   origHeight = innerHeight;");
            sb.AppendLine(@"   onresize = reDo;");
            sb.AppendLine(@"}");
            sb.AppendLine(@"onerror = null;");
            sb.AppendLine(@"//]]>");
            sb.AppendLine(@"</script>");
            sb.AppendLine(@"<style type=""text/css"">");
            sb.AppendLine(@"<!--");
            sb.AppendLine(@"div.WebHelpPopupMenu { position:absolute;");
            sb.AppendLine(@"left:0px;");
            sb.AppendLine(@"top:0px;");
            sb.AppendLine(@"z-index:4;");
            sb.AppendLine(@"visibility:hidden; }");
            sb.AppendLine(@"-->");
            return sb.ToString();
        }

        private string BuildEdgeTrackerCoverCss()
        {
            var sb = new StringBuilder();
            sb.AppendLine();
            sb.AppendLine(@"p.HyousiLogo {");
            sb.AppendLine(@"text-align       : center;");
            sb.AppendLine(@"margin-top       : 60pt;");
            sb.AppendLine(@"margin-bottom    : 40pt;");
            sb.AppendLine(@"margin-right     : 0mm;");
            sb.AppendLine(@"line-height      : 15pt;");
            sb.AppendLine(@"}");
            sb.AppendLine();
            sb.AppendLine(@"div.HyousiBackground {");
            sb.AppendLine(@"display : table;");
            sb.AppendLine(@"width   : 100%;");
            sb.AppendLine(@"height  : 65px;");
            sb.AppendLine(@"}");
            sb.AppendLine();
            sb.AppendLine(@"p.HyousiText {");
            sb.AppendLine(@"display             : table-cell;");
            sb.AppendLine(@"background-image    : url('pict/hyousi.png');");
            sb.AppendLine(@"background-repeat   : no-repeat;");
            sb.AppendLine(@"background-position : center;");
            sb.AppendLine(@"text-align          : center;");
            sb.AppendLine(@"vertical-align      : middle;");
            sb.AppendLine(@"font-size           : 1.8em;");
            sb.AppendLine(@"font-weight         : bold;");
            sb.AppendLine(@"color               : #FFF;");
            sb.AppendLine(@"letter-spacing      : 10px;");
            sb.AppendLine(@"}");
            return sb.ToString();
        }

        private string BuildEdgeTrackerCoverHtml(
            System.Reflection.Assembly assembly,
            string rootPath,
            string exportDir,
            string manualTitle,
            string trademarkTitle,
            List<string> trademarkTextList,
            string trademarkRight)
        {
            string pictDir = Path.Combine(rootPath, exportDir, "pict");

            string[] hyousiGazo = { "EdgeTracker_logo50mm.png", "MJS_LOGO_255.gif", "hyousi.png" };
            foreach (var hyousi in hyousiGazo)
            {
                using (var bmp = new Bitmap(assembly.GetManifestResourceStream("WordAddIn1.Resources." + hyousi)))
                {
                    bmp.Save(Path.Combine(pictDir, hyousi));
                }
            }

            var sb = new StringBuilder();
            sb.AppendLine(@"<body>");
            sb.AppendLine(@"<p class=""HyousiLogo""><img style=""border: currentColor; border-image: none; width: 100%; max-width: 553px;"" alt="""" src=""pict/EdgeTracker_logo50mm.png"" border=""0""></p>");
            sb.AppendLine(@"<div class=""HyousiBackground"">");
            sb.AppendLine($@"<p class=""HyousiText"">{manualTitle}</p>");
            sb.AppendLine(@"</div>");
            sb.AppendLine(@"<div class=""product_trademarks"">");
            sb.AppendLine($@"  <p class=""trademark_title"">{trademarkTitle}</p>");
            foreach (string trademarkText in trademarkTextList)
            {
                sb.AppendLine($@"  <p class=""trademark_text"">{trademarkText}</p>");
            }
            sb.AppendLine($@"  <p class=""trademark_right"">{trademarkRight}</p>");
            sb.AppendLine(@"</div>");
            sb.AppendLine(@"<p style=""text-align: center; margin-top: 80pt;""><a href=""https://www.mjs.co.jp"" target=""_blank""><img style=""border: currentColor; border-image: none; width: 100%; max-width: 255px;"" alt="""" src=""pict/MJS_LOGO_255.gif"" border=""0""></a></p>");
            return sb.ToString();
        }

        private string BuildEasyCloudCoverHtml(
            string rootPath,
            string exportDir,
            string manualTitle,
            string manualSubTitle,
            string manualVersion,
            string trademarkTitle,
            List<string> trademarkTextList,
            string trademarkRight)
        {
            var sb = new StringBuilder();

            string coverBackgroundPath = Path.Combine(rootPath, exportDir, "template", "images", "cover-background.png");
            string cover4Path = Path.Combine(rootPath, exportDir, "template", "images", "cover-4.png");

            if (File.Exists(coverBackgroundPath))
                sb.AppendLine(@"<body style=""text-justify-trim: punctuation; background-image: url('template/images/cover-background.png');background-repeat: no-repeat; background-position: 0px 300px;"">");
            else
                sb.AppendLine(@"<body>");

            sb.AppendLine($@"<p class=""manual_title"" style=""line-height: 130%;"">{manualTitle}</p>");
            sb.AppendLine($@"<p class=""manual_subtitle"">{manualSubTitle}</p>");

            if (File.Exists(cover4Path))
                sb.AppendLine(@"<p class=""manual_title"" style=""margin: 80px 0px 80px 100px; ""><img src=""template/images/cover-4.png"" width=""650"" /></p>");
            else
                sb.AppendLine(@"<p class=""manual_title"" style=""margin: 80px 0px 80px 100px; ""></p>");

            sb.AppendLine($@"<p class=""manual_version"">{manualVersion}</p>");
            sb.AppendLine(@"<div class=""product_trademarks"">");
            sb.AppendLine($@"  <p class=""trademark_title"">{trademarkTitle}</p>");
            foreach (string trademarkText in trademarkTextList)
            {
                sb.AppendLine($@"  <p class=""trademark_text"">{trademarkText}</p>");
            }
            sb.AppendLine($@"  <p class=""trademark_right"">{trademarkRight}</p>");
            sb.AppendLine(@"</div>");

            return sb.ToString();
        }

        private string BuildEasyCloudSubTitleSection(string subTitle)
        {
            var sb = new StringBuilder();
            if (!string.IsNullOrEmpty(subTitle))
            {
                sb.AppendLine(@"<p style=""margin-left: 700px; margin-top: 150px; font-size: 15pt; font-family: メイリオ;");
                sb.AppendLine(@"    font-weight: bold;"">" + subTitle + "</p>");
                sb.AppendLine(@"<p><a href=""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）""");
                sb.AppendLine(@"                                        style=""margin-left: 700px; margin-top: 10px;""");
                sb.AppendLine(@"                                        width=""132"" height=""48"" /></a>");
            }
            else
            {
                sb.AppendLine(@"<p><a href=""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）""");
                sb.AppendLine(@"                                        style=""margin-left: 700px; margin-top: 100px;""");
                sb.AppendLine(@"                                        width=""132"" height=""48"" /></a>");
            }
            return sb.ToString();
        }

        private string GeneratePattern1CoverHtml(
            string manualTitle,
            string manualTitleCenter,
            string manualSubTitle,
            string manualSubTitleCenter,
            string trademarkTitle,
            List<string> trademarkTextList,
            string trademarkRight)
        {
            var sb = new StringBuilder();
            sb.AppendFormat(@"<p class=""manual_title"" style=""line-height: 130%; "">{0}</p>" + "\n",
                !string.IsNullOrWhiteSpace(manualTitle) ? manualTitle : manualTitleCenter);
            sb.AppendFormat(@"<p class=""manual_subtitle"">{0}</p>" + "\n",
                !string.IsNullOrWhiteSpace(manualSubTitle) ? manualSubTitle : manualSubTitleCenter);
            sb.AppendLine(@"<p class=""product_logo_main_nosub"">");
            sb.AppendLine(@"  <img src = ""template/images/product_logo_main.png"" alt=""製品ロゴ（メイン）"">");
            sb.AppendLine(@"</p>");
            sb.AppendLine(@"<div class=""product_trademarks"">");
            sb.AppendFormat(@"  <p class=""trademark_title"">{0}</p>" + "\n", trademarkTitle);
            foreach (string trademarkText in trademarkTextList)
            {
                sb.AppendFormat(@"  <p class=""trademark_text"">{0}</p>" + "\n", trademarkText);
            }
            sb.AppendFormat(@"  <p class=""trademark_right"">{0}</p>" + "\n", trademarkRight);
            sb.AppendLine(@"</div>");
            sb.AppendLine(@"<p><a href = ""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）"" style=""margin-left: 700px; margin-top: 100px;"" width=""132"" height=""48"" /></a>");
            sb.AppendLine(@"</p>");
            return sb.ToString();
        }

        private string GeneratePattern2CoverHtml(
            List<List<string>> productSubLogoGroups,
            string manualTitleCenter,
            string manualTitle,
            string manualSubTitleCenter,
            string manualSubTitle,
            string manualVersionCenter,
            string manualVersion,
            string trademarkTitle,
            List<string> trademarkTextList,
            string trademarkRight)
        {
            var sb = new StringBuilder();
            sb.AppendLine(@"<p class=""product_logo_main"">");
            sb.AppendLine(@"  <img src = ""template/images/product_logo_main.png"" alt=""製品ロゴ（メイン）"">");
            sb.AppendLine(@"</p>");
            sb.AppendLine(@"<div class=""product_logo_sub"">");
            foreach (List<string> subLogoGroup in productSubLogoGroups)
            {
                sb.AppendLine(@"<div>");
                foreach (string subLogoFileName in subLogoGroup)
                {
                    sb.AppendFormat(@"  <img src = ""template/images/{0}"" alt=""製品ロゴ（サブ）"">" + "\n", subLogoFileName);
                }
                sb.AppendLine(@"</div>");
            }
            sb.AppendLine(@"</div>");
            sb.AppendFormat(@"<p class=""manual_title_center"" style=""line-height: 130%; "">{0}</p>" + "\n",
                !string.IsNullOrWhiteSpace(manualTitleCenter) ? manualTitleCenter : manualTitle);
            sb.AppendFormat(@"<p class=""manual_subtitle_center"">{0}</p>" + "\n",
                !string.IsNullOrWhiteSpace(manualSubTitleCenter) ? manualSubTitleCenter : manualSubTitle);
            sb.AppendFormat(@"<p class=""manual_version_center"">{0}</p>" + "\n",
                !string.IsNullOrWhiteSpace(manualVersionCenter) ? manualVersionCenter : manualVersion);
            sb.AppendLine(@"<div class=""product_trademarks"">");
            sb.AppendFormat(@"  <p class=""trademark_title"">{0}</p>" + "\n", trademarkTitle);
            foreach (string trademarkText in trademarkTextList)
            {
                sb.AppendFormat(@"  <p class=""trademark_text"">{0}</p>" + "\n", trademarkText);
            }
            sb.AppendFormat(@"  <p class=""trademark_right"">{0}</p>" + "\n", trademarkRight);
            sb.AppendLine(@"</div>");
            sb.AppendLine(@"<p><a href = ""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）"" style=""margin-left: 700px; margin-top: 100px;"" width=""132"" height=""48"" /></a>");
            sb.AppendLine(@"</p>");
            return sb.ToString();
        }

        private string BuildHtmlCoverFooter()
        {
            var sb = new StringBuilder();
            sb.AppendLine(@"<script type=""text/javascript"" language=""javascript1.2"">//<![CDATA[");
            sb.AppendLine(@"<!--");
            sb.AppendLine(@"if (window.writeIntopicBar)");
            sb.AppendLine(@"   writeIntopicBar(0);");
            sb.AppendLine(@"//-->");
            sb.AppendLine(@"//]]></script>");
            sb.AppendLine(@"</body>");
            sb.AppendLine(@"</html>");
            return sb.ToString();
        }
    }
}
