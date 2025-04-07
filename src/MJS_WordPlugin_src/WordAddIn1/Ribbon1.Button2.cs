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
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document thisDocument = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument;
            WordAddIn1.Globals.ThisAddIn.Application.DocumentChange -= new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
            WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

            //            WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange -= delegate (Word.Selection mySelection) { Application_WindowSelectionChange(); };

            List<string> styleList = new List<string>();

            int selStart = WordAddIn1.Globals.ThisAddIn.Application.Selection.Start;
            int selEnd = WordAddIn1.Globals.ThisAddIn.Application.Selection.End;
            WordAddIn1.Globals.ThisAddIn.Application.Selection.EndKey(Word.WdUnits.wdStory);
            Application.DoEvents();
            WordAddIn1.Globals.ThisAddIn.Application.Selection.HomeKey(Word.WdUnits.wdStory);
            Application.DoEvents();

            string attachedTemplateFile = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument.get_AttachedTemplate().Path + @"\" +
                WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument.get_AttachedTemplate().Name;
            Word.Document WD = WordAddIn1.Globals.ThisAddIn.Application.Documents.Open(attachedTemplateFile);
            using (StreamWriter log = new StreamWriter(thisDocument.Path + "\\log.txt", true, Encoding.UTF8))
            {
                log.WriteLine("Attached template file: " + attachedTemplateFile);
            }
            button3.Enabled = false;


            WordAddIn1.Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Word.Style stl in WD.Styles)
            {
                if (stl.NameLocal.Contains("MJS"))
                    styleList.Add(stl.NameLocal);
            }

            WD.Close();

            var activeDoc = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument as Microsoft.Office.Interop.Word.Document;
            Word.Selection ws = WordAddIn1.Globals.ThisAddIn.Application.Selection;

            WordAddIn1.Globals.ThisAddIn.Application.Selection.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToFirst);
            activeDoc.Revisions.AcceptAll();

            activeDoc.ActiveWindow.View.ShowRevisionsAndComments = true;
            activeDoc.ActiveWindow.View.ShowInkAnnotations = false;
            activeDoc.ActiveWindow.View.ShowComments = true;
            activeDoc.ActiveWindow.View.ShowInsertionsAndDeletions = false;
            activeDoc.ActiveWindow.View.ShowFormatChanges = false;

            bool bl = false;
            //toggleButton1.Checked = false;

            //foreach (Word.Shape wsp in activeDoc.Shapes)
            //{
            //    if (wsp.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
            //    {
            //        wsp.Select();
            //         if(WordAddIn1.Globals.ThisAddIn.Application.Selection.ShapeRange.WrapFormat.Type != Word.WdWrapType.wdWrapInline)
            //            WordAddIn1.Globals.ThisAddIn.Application.Selection.ShapeRange.WrapFormat.Type = Word.WdWrapType.wdWrapInline;
            //    }
            //}

            //foreach (Word.InlineShape wis in activeDoc.InlineShapes)
            //{
            //    try
            //    {
            //        //MessageBox.Show(wis.Field.Code.Text);
            //        if (wis.Field.Code.Text == @" SHAPE  \* MERGEFORMAT ")
            //            wis.Delete();
            //    }
            //    catch { }
            //}

            foreach (Word.Comment c in activeDoc.Comments)
            {
                if (c.Range.Text.Contains("使用できない書式です。"))
                    c.Delete();
                else if (c.Range.Text.Contains("使用できない文字列です。"))
                    c.Delete();
                else if (c.Range.Text.Contains("描画キャンバス外に行内配置でない画像があります。"))
                    c.Delete();
                else if (c.Range.Text.Contains("上の段落に【MJS_手順番号リセット用】スタイルを挿入してください。"))
                    c.Delete();
                else if (c.Range.Text.Contains("描画キャンバスが行内配置ではありません。"))
                    c.Delete();
            }

            activeDoc.Range(0, 0).Select();
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.ClearFormatting();
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.Replacement.ClearFormatting();
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.Text = "^s";
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.Forward = true;
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.Wrap = Word.WdFindWrap.wdFindStop;
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.Format = false;
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.MatchCase = false;
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.MatchWholeWord = false;
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.MatchByte = false;
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.MatchAllWordForms = false;
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.MatchSoundsLike = false;
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.MatchWildcards = false;
            WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.MatchFuzzy = false;

            Application.DoEvents();
            while (WordAddIn1.Globals.ThisAddIn.Application.Selection.Find.Execute())
            {
                WordAddIn1.Globals.ThisAddIn.Application.Selection.Range.Comments.Add(WordAddIn1.Globals.ThisAddIn.Application.Selection.Range,
                    "【改行なしスペース】\r\n使用できない文字列です。");
                bl = true;
            }

            Application.DoEvents();

            foreach (Word.Shape sp in activeDoc.Shapes)
            {
                //if(sp.Type != Microsoft.Office.Core.MsoShapeType.msoCanvas)
                //{
                sp.Select();
                //sp.Anchor.Select();
                string shpType = "";
                switch (sp.Type)
                {
                    case Microsoft.Office.Core.MsoShapeType.msoAutoShape:
                        shpType = "オートシェイプ";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoCallout:
                        shpType = "引き出し線";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoChart:
                        shpType = "グラフ";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoComment:
                        shpType = "コメント";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoDiagram:
                        shpType = "ダイアグラム";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoEmbeddedOLEObject:
                        shpType = "埋め込み OLE オブジェクト";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoFormControl:
                        shpType = "フォーム コントロール";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoFreeform:
                        shpType = "フリーフォーム";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoGroup:
                        shpType = "グループ";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoInk:
                        shpType = "インク";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoInkComment:
                        shpType = "インク コメント";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoLine:
                        shpType = "直線";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoLinkedOLEObject:
                        shpType = "リンク OLE オブジェクト";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoLinkedPicture:
                        shpType = "リンク画像";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoMedia:
                        shpType = "メディア";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoOLEControlObject:
                        shpType = "OLE コントロール オブジェクト";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoPicture:
                        shpType = "画像";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoPlaceholder:
                        shpType = "プレースホルダー";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoScriptAnchor:
                        shpType = "スクリプト アンカー";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoShapeTypeMixed:
                        shpType = "図形の種類の組み合わせ";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoSlicer:
                        shpType = "スライサー";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoSmartArt:
                        shpType = "スマートアート";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoTable:
                        shpType = "表";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoTextBox:
                        shpType = "テキストボックス";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoTextEffect:
                        shpType = "テキスト効果";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoWebVideo:
                        shpType = "Web ビデオ";
                        break;
                    case Microsoft.Office.Core.MsoShapeType.msoCanvas:
                        shpType = "描画キャンバス";
                        break;
                }
                try
                {
                    if (shpType == "描画キャンバス")
                    {
                        //                    Word.Comment canCom = WordAddIn1.Globals.ThisAddIn.Application.Selection.Range.Comments.Add(WordAddIn1.Globals.ThisAddIn.Application.Selection.Range,
                        //"【画像配置エラー】\r\n" + "画像種別：" + shpType + "\r\n描画キャンバスが行内配置ではありません。");
                        //if (sp.WrapFormat.Type.ToString() == "wdWrapInline")
                        //    //canCom.Delete();
                        //else { bl = true; }
                        //continue;
                        //Word.WrapFormat wWrap = sp.WrapFormat;
                        if (WordAddIn1.Globals.ThisAddIn.Application.Selection.ShapeRange.WrapFormat.Type != Word.WdWrapType.wdWrapInline)
                        {
                            WordAddIn1.Globals.ThisAddIn.Application.Selection.Range.Comments.Add(WordAddIn1.Globals.ThisAddIn.Application.Selection.Range,
                            "【画像配置エラー】\r\n" + "画像種別：" + shpType + "\r\n描画キャンバスが行内配置ではありません。");
                            bl = true;
                        }
                        //continue;
                    }
                    //Word.Comment newCom = WordAddIn1.Globals.ThisAddIn.Application.Selection.Range.Comments.Add(WordAddIn1.Globals.ThisAddIn.Application.Selection.Range,
                    //"【画像配置エラー】\r\n" + "画像種別：" + shpType + "\r\n描画キャンバス外に行内配置でない画像があります。");
                    else if (WordAddIn1.Globals.ThisAddIn.Application.Selection.ShapeRange.WrapFormat.Type != Word.WdWrapType.wdWrapBehind &&
                        WordAddIn1.Globals.ThisAddIn.Application.Selection.ShapeRange.WrapFormat.Type != Word.WdWrapType.wdWrapInline)
                    {
                        try
                        {
                            WordAddIn1.Globals.ThisAddIn.Application.Selection.Range.Comments.Add(WordAddIn1.Globals.ThisAddIn.Application.Selection.Range,
                        "【画像配置エラー】\r\n" + "画像種別：" + shpType + "\r\n描画キャンバス外に行内配置でない画像があります。");
                        }
                        catch
                        {
                            sp.Anchor.Comments.Add(sp.Anchor,
                        "【画像配置エラー】\r\n" + "画像種別：" + shpType + "\r\n描画キャンバス外に行内配置でない画像があります。");
                        }
                        bl = true;
                    }
                }
                catch
                {
                }
            }

            WordAddIn1.Globals.ThisAddIn.Application.ScreenUpdating = true;

            Application.DoEvents();
            ProgressBar.Show();
            ProgressBar.SetProgressBar(activeDoc.Paragraphs.Count);
            int pro = 0;
            Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
            TimeSpan ts;

            WordAddIn1.Globals.ThisAddIn.Application.Selection.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToFirst);

            bool processBl = false;
            bool processHalt = false;

            foreach (Word.Paragraph i in activeDoc.Paragraphs)
            {
                ProgressBar.SetProgressBarValue(++pro);
                ts = sw.Elapsed;
                ProgressBar.ProgressTime(ts);
                if (ProgressBar.mInstance.IsDisposed)
                {
                    foreach (Word.Comment c in activeDoc.Comments)
                    {
                        if (c.Range.Text.Contains("使用できない書式です。"))
                            c.Delete();
                        else if (c.Range.Text.Contains("使用できない文字列です。"))
                            c.Delete();
                        else if (c.Range.Text.Contains("描画キャンバス外に行内配置でない画像があります。"))
                            c.Delete();
                        else if (c.Range.Text.Contains("上の段落に【MJS_手順番号リセット用】スタイルを挿入してください。"))
                            c.Delete();
                        else if (c.Range.Text.Contains("描画キャンバスが行内配置ではありません。"))
                            c.Delete();
                    }
                    processHalt = true;
                    break;
                }
                try
                {
                    if (i.Range.ParagraphStyle() == activeDoc.Styles[-1].NameLocal && String.IsNullOrEmpty(i.Range.Text.Trim().Replace("\u0007", "")))
                    {
                        continue;
                    }
                    else if (i.Range.ParagraphStyle().Contains("見出し 7") || i.Range.ParagraphStyle().Contains("章扉-見出し1") || i.Range.ParagraphStyle().Contains("章扉-目次1") || i.Range.ParagraphStyle().Contains("奥付") || i.Range.ParagraphStyle().Contains("索引")) continue;
                    else if (!styleList.Contains(i.Range.ParagraphStyle()))
                    {
                        i.Range.Comments.Add(i.Range, "【" + i.Range.ParagraphStyle() + "】" + ":\r\n使用できない書式です。");
                        bl = true;
                    }

                    if (i.Range.ParagraphStyle() == "MJS_見出し-手順") processBl = true;
                    if (processBl && i.Range.ParagraphStyle() == "MJS_手順番号リセット用") processBl = false;
                    if (processBl && i.Range.ParagraphStyle() == "MJS_手順文")
                    {
                        processBl = false;
                        i.Range.Comments.Add(i.Range, "上の段落に【MJS_手順番号リセット用】スタイルを挿入してください。");
                        bl = true;
                    }
                }
                catch
                {
                    continue;
                }
            }
            sw.Stop();
            sw = null;

            WordAddIn1.Globals.ThisAddIn.Application.Selection.Start = selStart;
            WordAddIn1.Globals.ThisAddIn.Application.Selection.End = selEnd;

            if (processHalt)
            {
                WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

                //                WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange -= delegate (Word.Selection mySelection) { Application_WindowSelectionChange(); };
                button3.Enabled = false;
                MessageBox.Show("スタイルチェックが停止しました。\r\nチェック済み項目は全て破棄されます。", "スタイルチェック停止", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (bl == false)
            {
                //書誌情報出力
                //makeBookInfo();
                activeDoc.ShowRevisions = false;
                MessageBox.Show("スタイルチェックOKです。\r\n「HTML出力」ボタンをクリックするとHTMLが出力されます。", "スタイルチェックOK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                button3.Enabled = true;
                //mySelection = WordAddIn1.Globals.ThisAddIn.Application.Selection;
                checkOK = true;
                WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
                //WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange += delegate (Word.Selection mySelection) { Application_WindowSelectionChange(); };
                //Globals.ThisAddIn.Application.WindowSelectionChange -= Application_WindowSelectionChange;
                Application.DoEvents();
            }
            else
            {
                WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

                //                WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange -= delegate (Word.Selection mySelection) { Application_WindowSelectionChange(); };
                button3.Enabled = false;
                MessageBox.Show("スタイルチェックNGです。\r\n「校閲」タブ-「コメント」-「次へ」ボタンで\r\n使用できない書式を確認できます。", "スタイルチェックNG", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            ProgressBar.Close();
            ProgressBar.mInstance = null;
            //toggleButton1.Checked = true;
        }
    }
}
