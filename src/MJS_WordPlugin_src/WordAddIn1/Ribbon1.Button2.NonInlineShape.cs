using System;
using System.Collections.Generic;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        // 図形が行内配置になっていないことをコメント
        private void AddCommentForNonInlineShape(Word.Document activeDoc, ref bool bl)
        {
            var shapeTypeMap = new Dictionary<Microsoft.Office.Core.MsoShapeType, string>
            {
                { Microsoft.Office.Core.MsoShapeType.msoAutoShape, "オートシェイプ" },
                { Microsoft.Office.Core.MsoShapeType.msoCallout, "引き出し線" },
                { Microsoft.Office.Core.MsoShapeType.msoChart, "グラフ" },
                { Microsoft.Office.Core.MsoShapeType.msoComment, "コメント" },
                { Microsoft.Office.Core.MsoShapeType.msoDiagram, "ダイアグラム" },
                { Microsoft.Office.Core.MsoShapeType.msoEmbeddedOLEObject, "埋め込み OLE オブジェクト" },
                { Microsoft.Office.Core.MsoShapeType.msoFormControl, "フォーム コントロール" },
                { Microsoft.Office.Core.MsoShapeType.msoFreeform, "フリーフォーム" },
                { Microsoft.Office.Core.MsoShapeType.msoGroup, "グループ" },
                { Microsoft.Office.Core.MsoShapeType.msoInk, "インク" },
                { Microsoft.Office.Core.MsoShapeType.msoInkComment, "インク コメント" },
                { Microsoft.Office.Core.MsoShapeType.msoLine, "直線" },
                { Microsoft.Office.Core.MsoShapeType.msoLinkedOLEObject, "リンク OLE オブジェクト" },
                { Microsoft.Office.Core.MsoShapeType.msoLinkedPicture, "リンク画像" },
                { Microsoft.Office.Core.MsoShapeType.msoMedia, "メディア" },
                { Microsoft.Office.Core.MsoShapeType.msoOLEControlObject, "OLE コントロール オブジェクト" },
                { Microsoft.Office.Core.MsoShapeType.msoPicture, "画像" },
                { Microsoft.Office.Core.MsoShapeType.msoPlaceholder, "プレースホルダー" },
                { Microsoft.Office.Core.MsoShapeType.msoScriptAnchor, "スクリプト アンカー" },
                { Microsoft.Office.Core.MsoShapeType.msoShapeTypeMixed, "図形の種類の組み合わせ" },
                { Microsoft.Office.Core.MsoShapeType.msoSlicer, "スライサー" },
                { Microsoft.Office.Core.MsoShapeType.msoSmartArt, "スマートアート" },
                { Microsoft.Office.Core.MsoShapeType.msoTable, "表" },
                { Microsoft.Office.Core.MsoShapeType.msoTextBox, "テキストボックス" },
                { Microsoft.Office.Core.MsoShapeType.msoTextEffect, "テキスト効果" },
                { Microsoft.Office.Core.MsoShapeType.msoWebVideo, "Web ビデオ" },
                { Microsoft.Office.Core.MsoShapeType.msoCanvas, "描画キャンバス" }
            };

            foreach (Word.Shape sp in activeDoc.Shapes)
            {
                try
                {
                    // Shape の種類を取得
                    // TryGetValue()はキーが存在するかどうかを確認しつつ値を取得する
                    string shpType = shapeTypeMap.TryGetValue(sp.Type, out string typeName) ? typeName : "不明な種類";

                    // 描画キャンバスの場合の処理
                    if (shpType == "描画キャンバス")
                    {
                        if (sp.WrapFormat.Type != Word.WdWrapType.wdWrapInline)
                        {
                            sp.Anchor.Comments.Add(sp.Anchor,
                                $"【画像配置エラー】\r\n画像種別：{shpType}\r\n描画キャンバスが行内配置ではありません。");
                            bl = true;
                        }
                    }
                    // その他の Shape の処理
                    else if (sp.WrapFormat.Type != Word.WdWrapType.wdWrapBehind &&
                             sp.WrapFormat.Type != Word.WdWrapType.wdWrapInline)
                    {
                        try
                        {
                            sp.Anchor.Comments.Add(sp.Anchor,
                                $"【画像配置エラー】\r\n画像種別：{shpType}\r\n描画キャンバス外に行内配置でない画像があります。");
                            bl = true;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Shape コメントの追加中に例外が発生しました: {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Shape 処理中の例外をログに記録
                    Debug.WriteLine($"Shape 処理中に例外が発生しました: {ex.Message}");
                }
            }
        }
    }
}
