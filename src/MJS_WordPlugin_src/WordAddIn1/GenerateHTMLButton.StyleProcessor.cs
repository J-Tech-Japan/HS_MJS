using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // 指定されたCSSスタイル定義を解析し、条件に基づいてスタイル名を抽出・処理
        private void ProcessStyleBlock(string clsName, string pattern, ref string chapterSplitClass, Dictionary<string, string> styleName)
        {
            // 正規表現でスタイル定義を解析し、スタイル名を抽出
            var match = Regex.Match(clsName, $@"^(.+?){{.*{pattern}:(?:""(.+?)\\,.*""|(.+?));.*}}");
            if (match.Success)
            {
                // スタイル定義の前半部分（クラス名）を取得
                string clsBefore = match.Groups[1].Value;

                // スタイル名を取得（ダブルクォートで囲まれている場合とそうでない場合を考慮）
                string clsAfter = match.Groups[2].Success ? match.Groups[2].Value : match.Groups[3].Value;

                // クラス名をカンマで分割して個別に処理
                foreach (string cls in clsBefore.Split(','))
                {
                    // スタイル名が「章扉タイトル」に一致する場合、章分割クラスに追加
                    if (Regex.IsMatch(clsAfter, "章[　 ]*扉.*タイトル"))
                    {
                        if (!string.IsNullOrEmpty(chapterSplitClass))
                        {
                            chapterSplitClass += "|";
                        }
                        // クラス名をXPath形式に変換して追加
                        chapterSplitClass += Regex.Replace(cls, @"^(.+?)\.(.+?)$", "$1[@class='$2']");
                    }

                    // スタイル名を辞書に登録（バックスラッシュを削除）
                    styleName[cls] = clsAfter.Replace("\\", "");
                }
            }
        }

        // CSSスタイル定義全体を解析し、スタイル名や章分割クラスを抽出
        public void ProcessStyles(string className, ref string chapterSplitClass, Dictionary<string, string> styleName)
        {
            // スタイル定義を改行で分割して個別に処理
            foreach (string clsName in className.Split('\n'))
            {
                // "mso-style-name" を含むスタイル定義を処理
                if (clsName.Contains("mso-style-name:"))
                {
                    ProcessStyleBlock(clsName, "mso-style-name", ref chapterSplitClass, styleName);
                }
                // "mso-style-link" を含むスタイル定義を処理
                else if (clsName.Contains("mso-style-link:"))
                {
                    ProcessStyleBlock(clsName, "mso-style-link", ref chapterSplitClass, styleName);
                }
            }
        }

    }
}
