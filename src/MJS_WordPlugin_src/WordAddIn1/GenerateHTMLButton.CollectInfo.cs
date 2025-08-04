// GenerateHTMLButton.CollectInfo.cs

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
        public void CollectCoverParagraphs(
            Document docCopy,
            ref string manualTitle,
            ref string manualSubTitle,
            ref string manualVersion,
            ref string manualTitleCenter,
            ref string manualSubTitleCenter,
            ref string manualVersionCenter,
            ref bool coverExist)
        {
            // 第1セクションの全段落をループ処理
            foreach (Paragraph wp in docCopy.Sections[1].Range.Paragraphs)
            {
                // 段落のスタイル名とテキストを取得
                string styleName = wp.get_Style().NameLocal;
                string text = wp.Range.Text.Trim();
                
                // 空行または改行のみの段落をスキップ
                if (string.IsNullOrEmpty(text) || text == "/")
                    continue;

                // スタイル名に応じて適切な変数に追加
                switch (styleName)
                {
                    case "MJS_マニュアルタイトル":
                        // 通常位置のマニュアルタイトルを追加
                        manualTitle += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルサブタイトル":
                        // 通常位置のマニュアルサブタイトルを追加
                        manualSubTitle += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルバージョン":
                        // 通常位置のマニュアルバージョンを追加
                        manualVersion += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルタイトル（中央）":
                        // 中央配置のマニュアルタイトルを追加
                        manualTitleCenter += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルサブタイトル（中央）":
                        // 中央配置のマニュアルサブタイトルを追加
                        manualSubTitleCenter += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルバージョン（中央）":
                        // 中央配置のマニュアルバージョンを追加
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
                // 商標および著作権情報の検出フラグ
                bool isTradeMarksDetected = false;
                bool isRightDetected = false;

                // 最終セクションの全段落をループ処理
                foreach (Paragraph wp in docCopy.Sections[lastSectionIdx].Range.Paragraphs)
                {
                    string wpTextTrim = wp.Range.Text.Trim();
                    string wpStyleName = wp.get_Style().NameLocal;

                    // 空行や無効な行をスキップ
                    if (string.IsNullOrEmpty(wpTextTrim) || wpTextTrim == "/")
                    {
                        continue;
                    }

                    // 商標タイトルの検出（見出し4または見出し5で「商標」を含む）
                    if (!isTradeMarksDetected && wpTextTrim.Contains("商標") &&
                        (wpStyleName.Contains("MJS_見出し 4") || wpStyleName.Contains("MJS_見出し 5")))
                    {
                        trademarkTitle = wpTextTrim + "<br/>";
                        isTradeMarksDetected = true;
                        continue;
                    }

                    // 商標情報のリスト追加（箇条書きスタイルの段落）
                    if (isTradeMarksDetected && !isRightDetected &&
                        (wpStyleName.Contains("MJS_箇条書き") || wpStyleName.Contains("MJS_箇条書き2")))
                    {
                        trademarkTextList.Add(wpTextTrim + "<br/>");
                        continue;
                    }

                    // 著作権情報の検出（「All rights reserved」を含むリード文スタイル）
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
                // エラーをログに記録
                log.WriteLine($"エラー: {ex.Message}");
                // ユーザーにエラーダイアログを表示
                MessageBox.Show($"商標および著作権情報の収集中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // マニュアルタイトル類の末尾のタグとベル文字を除去
        private void CleanUpManualTitles(
            ref string manualTitle,
            ref string manualSubTitle,
            ref string manualVersion,
            ref string manualTitleCenter,
            ref string manualSubTitleCenter,
            ref string manualVersionCenter)
        {
            // ベル文字（ASCII 7）を定義
            string bell = new string((char)7, 1);
            
            // 各タイトルから末尾の<br/>タグとベル文字を除去
            manualTitle = Regex.Replace(manualTitle, @"<br/>$", "").Replace(bell, "").Trim();
            manualSubTitle = Regex.Replace(manualSubTitle, @"<br/>$", "").Replace(bell, "").Trim();
            manualVersion = Regex.Replace(manualVersion, @"<br/>$", "").Replace(bell, "").Trim();
            manualTitleCenter = Regex.Replace(manualTitleCenter, @"<br/>$", "").Replace(bell, "").Trim();
            manualSubTitleCenter = Regex.Replace(manualSubTitleCenter, @"<br/>$", "").Replace(bell, "").Trim();
            manualVersionCenter = Regex.Replace(manualVersionCenter, @"<br/>$", "").Replace(bell, "").Trim();
        }
    }
}
