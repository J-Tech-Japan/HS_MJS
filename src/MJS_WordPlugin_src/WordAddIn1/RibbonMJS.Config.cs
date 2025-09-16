// RibbonMJS.Config.cs

using System;
using System.IO;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // HTML出力用パス一覧
        private (string rootPath,
            string docName,
            string docFullName,
            string exportDir,
            string headerDir,
            string exportDirPath,
            string logPath,
            string tmpHtmlPath,
            string indexHtmlPath,
            string tmpFolderForImagesSavedBySaveAs2Method,
            string docid, string docTitle, string zipDirPath)
            PreparePaths(Word.Document activeDocument, string webHelpFolderName)
        {
            string rootPath = activeDocument.Path;
            string docName = activeDocument.Name;
            string docFullName = activeDocument.FullName;
            string exportDir = string.IsNullOrEmpty(webHelpFolderName) ? "webHelp" : webHelpFolderName;
            string headerDir = "headerFile";
            string exportDirPath = Path.Combine(rootPath, exportDir);
            string logPath = Path.Combine(rootPath, "log.txt");
            string tmpHtmlPath = Path.Combine(rootPath, "tmp.html");
            string indexHtmlPath = Path.Combine(rootPath, exportDir, "index.html");
            string tmpFolderForImagesSavedBySaveAs2Method = Path.Combine(rootPath, "tmp.files");
            string docid = Regex.Replace(docName, "^(.{3}).+$", "$1");
            string docTitle = Regex.Replace(docName, @"^.{3}_?(.+?)(?:_.+)?\.[^\.]+$", "$1");
            string zipDirPath = Path.Combine(rootPath, $"{docid}_{exportDir}_{DateTime.Today:yyyyMMdd}");
            return (rootPath, docName, docFullName, exportDir, headerDir, exportDirPath, logPath, tmpHtmlPath, indexHtmlPath, tmpFolderForImagesSavedBySaveAs2Method, docid, docTitle, zipDirPath);
        }

        // 除去したい記号
        public static readonly char[] removeSymbols = { '\u00D8', '\u00B2', '\u00B3', '\u00B9'};

        // 新たな除去候補記号'\u00E8'
        public static readonly char[] removeSingleSymbols = { '\u00E8' };

        // ファイル名形式の規定
        private const string FileNamePattern = @"^[A-Z]{3}(_[^_]*?){2}\.docx*$";
        
        // 一般的なエラーメッセージ
        private const string ErrMsg = "エラーが発生しました。";

        // ヘルパーメソッド: ドキュメントの表示設定
        private void ConfigDocumentDisplay()
        {
            var view = Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View;
            view.ShowRevisionsAndComments = true;
            view.ShowInkAnnotations = false;
            view.ShowComments = true;
            view.ShowInsertionsAndDeletions = false;
            view.ShowFormatChanges = false;
        }

        // ヘルパーメソッド: 検索条件の設定
        private void ConfigSearchParameters()
        {
            var find = Globals.ThisAddIn.Application.Selection.Find;
            // 検索条件を設定
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Text = "^s";
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            find.Format = false;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchByte = false;
            find.MatchAllWordForms = false;
            find.MatchSoundsLike = false;
            find.MatchWildcards = false;
            find.MatchFuzzy = false;
        }
    }
}
