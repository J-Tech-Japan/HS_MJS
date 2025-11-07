// Utils.WordStyles.cs

using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace MJS_fileJoin
{
    internal partial class Utils
    {
        /// <summary>
        /// 指定されたWord文書で使用されているスタイルのリストを取得します
        /// </summary>
        /// <param name="document">対象のWord文書</param>
        /// <param name="includeBuiltIn">組み込みスタイルも含めるかどうか（省略時はfalse）</param>
        /// <returns>使用されているスタイル名のリスト</returns>
        public static IList<string> GetUsedStyles(Word.Document document, bool includeBuiltIn = false)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            var usedStyles = new HashSet<string>();

            try
            {
                // 全ての段落のスタイルを収集
                foreach (Word.Paragraph paragraph in document.Paragraphs)
                {
                    try
                    {
                        var style = paragraph.get_Style();
                        string styleName = GetStyleName(style);
                        
                        if (!string.IsNullOrEmpty(styleName))
                        {
                            // 組み込みスタイルのフィルタリング
                            if (includeBuiltIn || !IsBuiltInStyle(document, styleName))
                            {
                                usedStyles.Add(styleName);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // 段落スタイルの取得に失敗した場合はスキップ
                        continue;
                    }
                }

                // リストをソートして返す
                return usedStyles.OrderBy(s => s).ToList();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Word文書からスタイル情報の取得中にエラーが発生しました。", ex);
            }
        }

        /// <summary>
        /// スタイルオブジェクトから名前を取得します
        /// </summary>
        private static string GetStyleName(object style)
        {
            if (style == null)
                return null;

            try
            {
                if (style is Word.Style wordStyle)
                {
                    return wordStyle.NameLocal;
                }
                else
                {
                    return style.ToString();
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 指定されたスタイル名が組み込みスタイルかどうかを判定します
        /// </summary>
        private static bool IsBuiltInStyle(Word.Document document, string styleName)
        {
            try
            {
                foreach (Word.Style style in document.Styles)
                {
                    if (style.NameLocal == styleName)
                    {
                        return style.BuiltIn;
                    }
                }
            }
            catch (Exception)
            {
                // エラーが発生した場合は組み込みスタイルではないと判定
            }

            return false;
        }
    }
}
