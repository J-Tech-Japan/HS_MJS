using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        /*
        ProcessStyles メソッドは、CSSスタイルのような文字列（className）を解析し、特定の条件に基づいて以下のように処理します：
        1.スタイル名の抽出:
        ・mso-style-name または mso-style-link を含むスタイル定義を解析し、スタイル名を抽出します。
        2.条件に基づく処理:
        ・スタイル名が特定の条件（例: "章[　 ]*扉.*タイトル"）に一致する場合、chapterSplitClass に追加します。
        ・スタイル名を styleName 辞書に登録します。
        */

        public void ProcessStyles(string className, ref string chapterSplitClass, Dictionary<string, string> styleName)
        {
            foreach (string clsName in className.Split('\n'))
            {
                string clsBefore, clsAfter;

                if (Regex.IsMatch(clsName, "mso-style-name:"))
                {
                    clsBefore = Regex.Replace(clsName, "^(.+?){.+?}$", "$1");
                    clsAfter = Regex.Replace(clsName, @"^.+?{.*mso-style-name:""(.+?)\\,.*"";.*}", "$1");
                    clsAfter = Regex.Replace(clsAfter, "^.+?{.*mso-style-name:(.+?);.*}", "$1");

                    foreach (string cls in clsBefore.Split(','))
                    {
                        if (Regex.IsMatch(clsAfter, "章[　 ]*扉.*タイトル"))
                        {
                            if (chapterSplitClass != "")
                            {
                                chapterSplitClass += "|";
                            }
                            chapterSplitClass += Regex.Replace(cls, @"^(.+?)\.(.+?)$", "$1[@class='$2']");
                        }

                        styleName.Add(cls, Regex.Replace(clsAfter, @"\\", ""));
                    }
                }
                else if (Regex.IsMatch(clsName, "mso-style-link:"))
                {
                    clsBefore = Regex.Replace(clsName, "^(.+?){.+?}$", "$1");
                    clsAfter = Regex.Replace(clsName, @"^.+?{.*mso-style-link:""(.+?)\\,.*"";.*}", "$1");
                    clsAfter = Regex.Replace(clsAfter, "^.+?{.*mso-style-link:(.+?);.*}", "$1");

                    foreach (string cls in clsBefore.Split(','))
                    {
                        if (Regex.IsMatch(clsAfter, "章[　 ]*扉.*タイトル"))
                        {
                            if (chapterSplitClass != "")
                            {
                                chapterSplitClass += "|";
                            }
                            chapterSplitClass += Regex.Replace(cls, @"^(.+?)\.(.+?)$", "$1[@class='$2']");
                        }

                        styleName.Add(cls, Regex.Replace(clsAfter, @"\\", ""));
                    }
                }
            }
        }
    }
}
