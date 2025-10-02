//Utils.WordImageProcess.cs

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    internal partial class Utils
    {
        // キャンバスに関連する図形のプロパティを調整
        // キャンバス内の図形の位置調整とレイアウト最適化を実行
        // Word文書内のキャンバス図形を対象に、HTML出力前の図形調整を実施
        public static void AdjustCanvasShapes(
            Document docCopy, 
            float heightExpansion,           // キャンバス高さの拡張量（ポイント）
            float positionOffset,            // アイテムの位置調整オフセット（ポイント）
            bool skipTablesCanvases = true,          // テーブル内キャンバスをスキップするかどうか
            bool maintainAspectRatio = true          // 処理後にアスペクト比ロックを復元するかどうか
        )
        {
            // 文書内のすべての図形をループ処理
            foreach (Shape docS in docCopy.Shapes)
            {
                // キャンバス図形のみを処理対象とする
                if (docS.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
                {
                    // テーブル内のキャンバスは処理をスキップ（レイアウト崩れを防ぐため）
                    if (skipTablesCanvases && docS.Anchor != null && docS.Anchor.Tables.Count > 0)
                    {
                        continue;
                    }

                    // キャンバス内の各アイテムの元の位置・サイズ情報を保存
                    List<float> canvasItemsTop = new List<float>();
                    List<float> canvasItemsLeft = new List<float>();
                    List<float> canvasItemsHeight = new List<float>();
                    List<float> canvasItemsWidth = new List<float>();

                    // キャンバス内の各アイテムのプロパティを取得し保存
                    for (int i = 1; i <= docS.CanvasItems.Count; i++)
                    {
                        // アスペクト比ロックを解除して調整を可能にする
                        docS.CanvasItems[i].LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        canvasItemsTop.Add(docS.CanvasItems[i].Top);
                        canvasItemsLeft.Add(docS.CanvasItems[i].Left);
                        canvasItemsHeight.Add(docS.CanvasItems[i].Height);
                        canvasItemsWidth.Add(docS.CanvasItems[i].Width);
                    }

                    // キャンバス自体のサイズ調整（指定された値で高さを拡張）
                    docS.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                    docS.Height = docS.Height + heightExpansion;
                    if (maintainAspectRatio)
                    {
                        docS.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    }

                    // 保存した位置・サイズ情報を基に各アイテムを再配置
                    for (int i = 1; i <= docS.CanvasItems.Count; i++)
                    {
                        // アスペクト比ロックを解除して位置・サイズを調整
                        docS.CanvasItems[i].LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        // 元のサイズを復元
                        docS.CanvasItems[i].Height = canvasItemsHeight[i - 1];
                        docS.CanvasItems[i].Width = canvasItemsWidth[i - 1];
                        // 指定されたオフセットで位置を調整
                        docS.CanvasItems[i].Top = canvasItemsTop[i - 1] + positionOffset;
                        docS.CanvasItems[i].Left = canvasItemsLeft[i - 1];
                        // アスペクト比ロックを設定
                        if (maintainAspectRatio)
                        {
                            docS.CanvasItems[i].LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                        }
                    }
                }
            }
        }
    }
}
