// DocumentPropertySettings.cs

using System;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// Wordドキュメントのカスタムドキュメントプロパティから
    /// 高画質画像抽出機能に関連する設定を取得するヘルパークラス
    /// </summary>
    public static class DocumentPropertySettings
    {
        // プロパティ名の定数
        private const string PropertyExtractHighQualityImages = "extractHighQualityImages";
        private const string PropertyIsBetaMode = "isBetaMode";
        private const string PropertyOutputScaleMultiplier = "OutputScaleMultiplier";
        private const string PropertyTableImageScaleMultiplier = "TableImageScaleMultiplier";
        private const string PropertyColumnImageScaleMultiplier = "ColumnImageScaleMultiplier";

        // ブール値として認識される文字列
        private static readonly string[] TrueValues = { "true", "1", "yes" };
        private static readonly string[] FalseValues = { "false", "0", "no" };

        // デフォルト値
        private const bool DefaultExtractHighQualityImages = true;
        private const bool DefaultIsBetaMode = false;
        private const float DefaultOutputScaleMultiplier = 1.4f;
        private const float DefaultTableImageScaleMultiplier = 1.2f;
        private const float DefaultColumnImageScaleMultiplier = 1.2f;

        // スケール倍率の有効範囲
        private const float MinScaleMultiplier = 0.1f;
        private const float MaxScaleMultiplier = 10.0f;

        /// <summary>
        /// カスタムドキュメントプロパティから高画質画像抽出機能の設定を取得
        /// </summary>
        /// <param name="activeDocument">アクティブなWordドキュメント</param>
        /// <returns>高画質画像抽出を実行する場合はtrue、しない場合はfalse（デフォルト: true）</returns>
        /// <remarks>
        /// プロパティ名: "extractHighQualityImages"
        /// 有効な値: "true", "false", "1", "0", "yes", "no" (大文字小文字区別なし)
        /// </remarks>
        public static bool GetExtractHighQualityImagesSetting(Document activeDocument)
        {
            return GetBooleanProperty(activeDocument, PropertyExtractHighQualityImages, DefaultExtractHighQualityImages);
        }

        /// <summary>
        /// カスタムドキュメントプロパティからbeta版モードの設定を取得
        /// </summary>
        /// <param name="activeDocument">アクティブなWordドキュメント</param>
        /// <returns>beta版モードを有効にする場合はtrue、しない場合はfalse（デフォルト: false）</returns>
        /// <remarks>
        /// プロパティ名: "isBetaMode"
        /// 有効な値: "true", "false", "1", "0", "yes", "no" (大文字小文字区別なし)
        /// beta版モードでは、詳細ログとCSV出力が有効になります
        /// </remarks>
        public static bool GetBetaModeSetting(Document activeDocument)
        {
            return GetBooleanProperty(activeDocument, PropertyIsBetaMode, DefaultIsBetaMode);
        }

        /// <summary>
        /// カスタムドキュメントプロパティから画像スケール設定（OutputScaleMultiplier）を取得
        /// </summary>
        /// <param name="activeDocument">アクティブなWordドキュメント</param>
        /// <returns>スケール倍率（デフォルト: 1.4f）</returns>
        /// <remarks>
        /// プロパティ名: "OutputScaleMultiplier"
        /// 有効な値: 数値文字列（例: "1.0", "1.4", "2.0"）
        /// 有効範囲: 0.1〜10.0
        /// </remarks>
        public static float GetOutputScaleMultiplierSetting(Document activeDocument)
        {
            return GetFloatProperty(activeDocument, PropertyOutputScaleMultiplier, DefaultOutputScaleMultiplier, MinScaleMultiplier, MaxScaleMultiplier);
        }

        /// <summary>
        /// カスタムドキュメントプロパティから表内画像スケール設定（TableImageScaleMultiplier）を取得
        /// </summary>
        /// <param name="activeDocument">アクティブなWordドキュメント</param>
        /// <returns>スケール倍率（デフォルト: 1.2f）</returns>
        /// <remarks>
        /// プロパティ名: "TableImageScaleMultiplier"
        /// 有効な値: 数値文字列（例: "1.0", "1.2"）
        /// 有効範囲: 0.1〜10.0
        /// </remarks>
        public static float GetTableImageScaleMultiplierSetting(Document activeDocument)
        {
            return GetFloatProperty(activeDocument, PropertyTableImageScaleMultiplier, DefaultTableImageScaleMultiplier, MinScaleMultiplier, MaxScaleMultiplier);
        }

        /// <summary>
        /// カスタムドキュメントプロパティからコラム内画像スケール設定（ColumnImageScaleMultiplier）を取得
        /// </summary>
        /// <param name="activeDocument">アクティブなWordドキュメント</param>
        /// <returns>スケール倍率（デフォルト: 1.2f）</returns>
        /// <remarks>
        /// プロパティ名: "ColumnImageScaleMultiplier"
        /// 有効な値: 数値文字列（例: "1.0", "1.2"）
        /// 有効範囲: 0.1〜10.0
        /// </remarks>
        public static float GetColumnImageScaleMultiplierSetting(Document activeDocument)
        {
            return GetFloatProperty(activeDocument, PropertyColumnImageScaleMultiplier, DefaultColumnImageScaleMultiplier, MinScaleMultiplier, MaxScaleMultiplier);
        }

        /// <summary>
        /// カスタムドキュメントプロパティからブール値の設定を取得する汎用メソッド
        /// </summary>
        /// <param name="activeDocument">アクティブなWordドキュメント</param>
        /// <param name="propertyName">プロパティ名</param>
        /// <param name="defaultValue">デフォルト値</param>
        /// <returns>取得したブール値（取得失敗時はデフォルト値）</returns>
        private static bool GetBooleanProperty(Document activeDocument, string propertyName, bool defaultValue)
        {
            if (activeDocument == null)
            {
                System.Diagnostics.Debug.WriteLine($"{propertyName}設定の取得失敗: ドキュメントがnullです");
                return defaultValue;
            }

            Microsoft.Office.Core.DocumentProperties properties = null;
            try
            {
                properties = (Microsoft.Office.Core.DocumentProperties)activeDocument.CustomDocumentProperties;
                var property = properties.Cast<Microsoft.Office.Core.DocumentProperty>()
                    .FirstOrDefault(x => x.Name == propertyName);

                if (property != null)
                {
                    string value = property.Value?.ToString()?.ToLower();

                    if (!string.IsNullOrEmpty(value))
                    {
                        // "true", "1", "yes" の場合はtrue
                        if (TrueValues.Contains(value))
                        {
                            return true;
                        }
                        // "false", "0", "no" の場合はfalse
                        if (FalseValues.Contains(value))
                        {
                            return false;
                        }

                        // 有効な値ではない場合はログに記録
                        System.Diagnostics.Debug.WriteLine(
                            $"{propertyName}設定の値が無効です: '{value}' (有効な値: true/false/1/0/yes/no)");
                    }
                }
            }
            catch (Exception ex)
            {
                // プロパティ取得に失敗した場合はデバッグログに記録
                System.Diagnostics.Debug.WriteLine($"{propertyName}設定の取得に失敗: {ex.Message}");
            }
            finally
            {
                // COMオブジェクトを適切に解放
                if (properties != null)
                {
                    Marshal.ReleaseComObject(properties);
                }
            }

            // プロパティが存在しない場合、またはエラーが発生した場合はデフォルト値を返す
            return defaultValue;
        }

        /// <summary>
        /// カスタムドキュメントプロパティから浮動小数点数の設定を取得する汎用メソッド
        /// </summary>
        /// <param name="activeDocument">アクティブなWordドキュメント</param>
        /// <param name="propertyName">プロパティ名</param>
        /// <param name="defaultValue">デフォルト値</param>
        /// <param name="minValue">最小値</param>
        /// <param name="maxValue">最大値</param>
        /// <returns>取得した浮動小数点数（取得失敗時または範囲外の場合はデフォルト値）</returns>
        private static float GetFloatProperty(Document activeDocument, string propertyName, float defaultValue, float minValue, float maxValue)
        {
            if (activeDocument == null)
            {
                System.Diagnostics.Debug.WriteLine($"{propertyName}設定の取得失敗: ドキュメントがnullです");
                return defaultValue;
            }

            Microsoft.Office.Core.DocumentProperties properties = null;
            try
            {
                properties = (Microsoft.Office.Core.DocumentProperties)activeDocument.CustomDocumentProperties;
                var property = properties.Cast<Microsoft.Office.Core.DocumentProperty>()
                    .FirstOrDefault(x => x.Name == propertyName);

                if (property != null)
                {
                    string value = property.Value?.ToString();
                    if (!string.IsNullOrEmpty(value) && float.TryParse(value, out float result))
                    {
                        // 妥当な範囲内かチェック
                        if (result >= minValue && result <= maxValue)
                        {
                            return result;
                        }

                        System.Diagnostics.Debug.WriteLine(
                            $"{propertyName}設定の値が範囲外です: {result} (有効範囲: {minValue}〜{maxValue})");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"{propertyName}設定の値が数値として解析できません: '{value}'");
                    }
                }
            }
            catch (Exception ex)
            {
                // プロパティ取得に失敗した場合はデバッグログに記録
                System.Diagnostics.Debug.WriteLine($"{propertyName}設定の取得に失敗: {ex.Message}");
            }
            finally
            {
                // COMオブジェクトを適切に解放
                if (properties != null)
                {
                    Marshal.ReleaseComObject(properties);
                }
            }

            // プロパティが存在しない場合、またはエラーが発生した場合はデフォルト値を返す
            return defaultValue;
        }
    }
}
