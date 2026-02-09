// ThisAddIn.cs

using System;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // ドキュメントオープンイベントを登録
            Application.DocumentOpen += Application_DocumentOpen;
            
            // 新規ドキュメント作成イベントを登録
            Application.DocumentChange += Application_DocumentChange;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // イベントハンドラーを解除
                Application.DocumentOpen -= Application_DocumentOpen;
                Application.DocumentChange -= Application_DocumentChange;

                // 必要に応じてリソースを解放
                if (Globals.ThisAddIn.Application != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Globals.ThisAddIn.Application);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"アドイン終了時にエラーが発生しました: {ex.Message}");
                // シャットダウン時はメッセージボックスを表示しない（Wordの終了を妨げる可能性があるため）
            }
        }

        /// <summary>
        /// ドキュメントが開かれたときに自動実行されるイベントハンドラー
        /// カスタムドキュメントプロパティが未設定の場合、デフォルト値を設定する
        /// </summary>
        private void Application_DocumentOpen(Document doc)
        {
            try
            {
                // カスタムドキュメントプロパティが設定されていない場合、デフォルト値を設定
                EnsureCustomDocumentProperties(doc);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ドキュメントオープン時のプロパティ設定でエラー: {ex.Message}");
                // ドキュメントを開く処理は継続させる（エラーが発生しても影響を与えない）
            }
        }

        /// <summary>
        /// ドキュメントが変更されたとき（新規作成含む）に自動実行されるイベントハンドラー
        /// カスタムドキュメントプロパティが未設定の場合、デフォルト値を設定する
        /// </summary>
        private void Application_DocumentChange()
        {
            try
            {
                // アクティブドキュメントに対してプロパティを確認・設定
                if (Application.ActiveDocument != null)
                {
                    EnsureCustomDocumentProperties(Application.ActiveDocument);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ドキュメント変更時のプロパティ設定でエラー: {ex.Message}");
                // 処理は継続させる
            }
        }

        /// <summary>
        /// カスタムドキュメントプロパティが存在しない場合、デフォルト値で自動作成する
        /// </summary>
        /// <param name="doc">対象のWordドキュメント</param>
        private void EnsureCustomDocumentProperties(Document doc)
        {
            if (doc == null)
                return;

            try
            {
                var properties = (Microsoft.Office.Core.DocumentProperties)doc.CustomDocumentProperties;

                // extractHighQualityImages プロパティの確認と設定
                EnsureProperty(properties, "extractHighQualityImages", "true", 
                    "高画質画像抽出機能のデフォルト設定");

                // isBetaMode プロパティの確認と設定
                EnsureProperty(properties, "isBetaMode", "false", 
                    "Beta版モード（詳細ログ）のデフォルト設定");

                // OutputScaleMultiplier プロパティの確認と設定
                EnsureProperty(properties, "OutputScaleMultiplier", "1.4", 
                    "通常画像の出力スケール倍率のデフォルト設定");

                // TableImageScaleMultiplier プロパティの確認と設定
                EnsureProperty(properties, "TableImageScaleMultiplier", "1.2", 
                    "表内画像の出力スケール倍率のデフォルト設定");

                // ColumnImageScaleMultiplier プロパティの確認と設定
                EnsureProperty(properties, "ColumnImageScaleMultiplier", "1.2", 
                    "コラム内画像の出力スケール倍率のデフォルト設定");

                // webHelpFolderName は既存の動作を維持（未設定の場合は自動生成されるため、ここでは設定しない）
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"カスタムプロパティの確認・設定でエラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 指定されたプロパティが存在しない場合、デフォルト値で作成する
        /// </summary>
        /// <param name="properties">カスタムドキュメントプロパティのコレクション</param>
        /// <param name="propertyName">プロパティ名</param>
        /// <param name="defaultValue">デフォルト値</param>
        /// <param name="description">プロパティの説明（デバッグ用）</param>
        private void EnsureProperty(
            Microsoft.Office.Core.DocumentProperties properties, 
            string propertyName, 
            string defaultValue,
            string description)
        {
            try
            {
                // プロパティが既に存在するかチェック
                var existingProperty = properties.Cast<Microsoft.Office.Core.DocumentProperty>()
                    .FirstOrDefault(p => p.Name == propertyName);

                if (existingProperty == null)
                {
                    // プロパティが存在しない場合、デフォルト値で作成
                    properties.Add(
                        propertyName,
                        false, // LinkToContent = false (固定値)
                        Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                        defaultValue,
                        Type.Missing
                    );

                    System.Diagnostics.Debug.WriteLine(
                        $"カスタムプロパティを自動作成: {propertyName} = {defaultValue} ({description})");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"カスタムプロパティは既に存在: {propertyName} = {existingProperty.Value}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"プロパティ '{propertyName}' の作成に失敗: {ex.Message}");
            }
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// メソッドの内容をコードエディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}