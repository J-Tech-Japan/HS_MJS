using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // 表紙に関連する段落を収集
        public void CollectCoverParagraphs(Document docCopy, ref string manualTitle, ref string manualSubTitle, ref string manualVersion,
                                      ref string manualTitleCenter, ref string manualSubTitleCenter, ref string manualVersionCenter,
                                      ref bool coverExist)
        {
            foreach (Paragraph wp in docCopy.Sections[1].Range.Paragraphs)
            {
                string styleName = wp.get_Style().NameLocal;
                string text = wp.Range.Text.Trim();
                if (string.IsNullOrEmpty(text) || text == "/")
                    continue;

                switch (styleName)
                {
                    case "MJS_マニュアルタイトル":
                        manualTitle += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルサブタイトル":
                        manualSubTitle += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルバージョン":
                        manualVersion += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルタイトル（中央）":
                        manualTitleCenter += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルサブタイトル（中央）":
                        manualSubTitleCenter += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルバージョン（中央）":
                        manualVersionCenter += text + "<br/>";
                        coverExist = true;
                        break;
                }
            }
        }

        // 商標情報と著作権情報を収集
        public void CollectTrademarkAndCopyrightDetails(
            Document docCopy,
            int lastSectionIdx,
            StreamWriter log,
            ref string trademarkTitle,
            ref List<string> trademarkTextList,
            ref string trademarkRight)
        {
            try
            {
                bool isTradeMarksDetected = false;
                bool isRightDetected = false;

                foreach (Paragraph wp in docCopy.Sections[lastSectionIdx].Range.Paragraphs)
                {
                    string wpTextTrim = wp.Range.Text.Trim();
                    string wpStyleName = wp.get_Style().NameLocal;

                    // ログに段落の内容を記録
                    //log.WriteLine($"[Style: {wpStyleName}] {wpTextTrim}");

                    // 空行や無効な行をスキップ
                    if (string.IsNullOrEmpty(wpTextTrim) || wpTextTrim == "/")
                    {
                        continue;
                    }

                    // 商標タイトルの検出
                    if (!isTradeMarksDetected && wpTextTrim.Contains("商標") &&
                        (wpStyleName.Contains("MJS_見出し 4") || wpStyleName.Contains("MJS_見出し 5")))
                    {
                        trademarkTitle = wpTextTrim + "<br/>";
                        isTradeMarksDetected = true;
                        continue;
                    }

                    // 商標情報のリスト追加
                    if (isTradeMarksDetected && !isRightDetected &&
                        (wpStyleName.Contains("MJS_箇条書き") || wpStyleName.Contains("MJS_箇条書き2")))
                    {
                        trademarkTextList.Add(wpTextTrim + "<br/>");
                        continue;
                    }

                    // 著作権情報の検出
                    if (!isRightDetected && wpTextTrim.Contains("All rights reserved") &&
                        wpStyleName.Contains("MJS_リード文"))
                    {
                        trademarkRight = wpTextTrim + "<br/>";
                        isRightDetected = true;
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                log.WriteLine($"エラー: {ex.Message}");
                MessageBox.Show($"商標および著作権情報の収集中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CleanUpManualTitles(
            ref string manualTitle,
            ref string manualSubTitle,
            ref string manualVersion,
            ref string manualTitleCenter,
            ref string manualSubTitleCenter,
            ref string manualVersionCenter)
        {
            string bell = new string((char)7, 1);
            manualTitle = Regex.Replace(manualTitle, @"<br/>$", "").Replace(bell, "").Trim();
            manualSubTitle = Regex.Replace(manualSubTitle, @"<br/>$", "").Replace(bell, "").Trim();
            manualVersion = Regex.Replace(manualVersion, @"<br/>$", "").Replace(bell, "").Trim();
            manualTitleCenter = Regex.Replace(manualTitleCenter, @"<br/>$", "").Replace(bell, "").Trim();
            manualSubTitleCenter = Regex.Replace(manualSubTitleCenter, @"<br/>$", "").Replace(bell, "").Trim();
            manualVersionCenter = Regex.Replace(manualVersionCenter, @"<br/>$", "").Replace(bell, "").Trim();
        }
    }
}
