// ApplicationSettings.cs

using System;

namespace WordAddIn1
{
    /// <summary>
    /// Word全体で共有する高画質画像抽出機能に関連する設定を管理するヘルパークラス
    /// </summary>
    public static class ApplicationSettings
    {
        // スケール倍率の有効範囲
        private const float MinScaleMultiplier = 0.1f;
        private const float MaxScaleMultiplier = 10.0f;

        // 出力画像サイズの有効範囲
        private const int MinOutputSize = 100;
        private const int MaxOutputSize = 4096;

        // デフォルト値の定義（Settings.settingsと一致させる）
        private const bool DefaultExtractHighQualityImages = true;
        private const bool DefaultIsBetaMode = false;
        private const float DefaultOutputScaleMultiplier = 1.5f;
        private const float DefaultTableImageScaleMultiplier = 1.5f;
        private const float DefaultColumnImageScaleMultiplier = 1.5f;
        private const int DefaultMaxOutputWidth = 1024;
        private const int DefaultMaxOutputHeight = 1024;
        private const bool DefaultShowSettingsButton = true;

        /// <summary>
        /// デフォルト値を取得するための構造体
        /// </summary>
        public struct DefaultValues
        {
            public bool ExtractHighQualityImages { get; }
            public bool IsBetaMode { get; }
            public float OutputScaleMultiplier { get; }
            public float TableImageScaleMultiplier { get; }
            public float ColumnImageScaleMultiplier { get; }
            public int MaxOutputWidth { get; }
            public int MaxOutputHeight { get; }
            public bool ShowSettingsButton { get; }

            internal DefaultValues(bool dummy)
            {
                ExtractHighQualityImages = DefaultExtractHighQualityImages;
                IsBetaMode = DefaultIsBetaMode;
                OutputScaleMultiplier = DefaultOutputScaleMultiplier;
                TableImageScaleMultiplier = DefaultTableImageScaleMultiplier;
                ColumnImageScaleMultiplier = DefaultColumnImageScaleMultiplier;
                MaxOutputWidth = DefaultMaxOutputWidth;
                MaxOutputHeight = DefaultMaxOutputHeight;
                ShowSettingsButton = DefaultShowSettingsButton;
            }
        }

        /// <summary>
        /// すべてのデフォルト値を取得
        /// </summary>
        /// <returns>デフォルト値を格納した構造体</returns>
        public static DefaultValues GetDefaultValues()
        {
            return new DefaultValues(true);
        }

        /// <summary>
        /// すべての設定をデフォルト値にリセット
        /// </summary>
        public static void ResetAllToDefaults()
        {
            SetExtractHighQualityImagesSetting(DefaultExtractHighQualityImages);
            SetBetaModeSetting(DefaultIsBetaMode);
            SetOutputScaleMultiplierSetting(DefaultOutputScaleMultiplier);
            SetTableImageScaleMultiplierSetting(DefaultTableImageScaleMultiplier);
            SetColumnImageScaleMultiplierSetting(DefaultColumnImageScaleMultiplier);
            SetMaxOutputWidthSetting(DefaultMaxOutputWidth);
            SetMaxOutputHeightSetting(DefaultMaxOutputHeight);
            SetShowSettingsButtonSetting(DefaultShowSettingsButton);
        }

        /// <summary>
        /// 高画質画像抽出機能の設定を取得
        /// </summary>
        /// <returns>高画質画像抽出を実行する場合はtrue、しない場合はfalse（デフォルト: true）</returns>
        public static bool GetExtractHighQualityImagesSetting()
        {
            try
            {
                return Properties.Settings.Default.extractHighQualityImages;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"extractHighQualityImages設定の取得に失敗: {ex.Message}");
                return DefaultExtractHighQualityImages;
            }
        }

        /// <summary>
        /// 高画質画像抽出機能の設定を保存
        /// </summary>
        /// <param name="value">設定値</param>
        public static void SetExtractHighQualityImagesSetting(bool value)
        {
            try
            {
                Properties.Settings.Default.extractHighQualityImages = value;
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"extractHighQualityImages設定の保存に失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// beta版モードの設定を取得
        /// </summary>
        /// <returns>beta版モードを有効にする場合はtrue、しない場合はfalse（デフォルト: false）</returns>
        /// <remarks>
        /// beta版モードでは、詳細ログとCSV出力が有効になります
        /// </remarks>
        public static bool GetBetaModeSetting()
        {
            try
            {
                return Properties.Settings.Default.isBetaMode;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"isBetaMode設定の取得に失敗: {ex.Message}");
                return DefaultIsBetaMode;
            }
        }

        /// <summary>
        /// beta版モードの設定を保存
        /// </summary>
        /// <param name="value">設定値</param>
        public static void SetBetaModeSetting(bool value)
        {
            try
            {
                Properties.Settings.Default.isBetaMode = value;
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"isBetaMode設定の保存に失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 画像スケール設定（OutputScaleMultiplier）を取得
        /// </summary>
        /// <returns>スケール倍率（デフォルト: 1.4f）</returns>
        /// <remarks>
        /// 有効範囲: 0.1?10.0
        /// </remarks>
        public static float GetOutputScaleMultiplierSetting()
        {
            return GetFloatSetting(
                () => Properties.Settings.Default.outputScaleMultiplier,
                DefaultOutputScaleMultiplier,
                "outputScaleMultiplier"
            );
        }

        /// <summary>
        /// 画像スケール設定（OutputScaleMultiplier）を保存
        /// </summary>
        /// <param name="value">設定値</param>
        /// <returns>保存に成功した場合はtrue、範囲外の値の場合はfalse</returns>
        public static bool SetOutputScaleMultiplierSetting(float value)
        {
            return SetFloatSetting(
                value,
                v => Properties.Settings.Default.outputScaleMultiplier = v,
                "outputScaleMultiplier"
            );
        }

        /// <summary>
        /// 表内画像スケール設定（TableImageScaleMultiplier）を取得
        /// </summary>
        /// <returns>スケール倍率（デフォルト: 1.2f）</returns>
        /// <remarks>
        /// 有効範囲: 0.1?10.0
        /// </remarks>
        public static float GetTableImageScaleMultiplierSetting()
        {
            return GetFloatSetting(
                () => Properties.Settings.Default.tableImageScaleMultiplier,
                DefaultTableImageScaleMultiplier,
                "tableImageScaleMultiplier"
            );
        }

        /// <summary>
        /// 表内画像スケール設定（TableImageScaleMultiplier）を保存
        /// </summary>
        /// <param name="value">設定値</param>
        /// <returns>保存に成功した場合はtrue、範囲外の値の場合はfalse</returns>
        public static bool SetTableImageScaleMultiplierSetting(float value)
        {
            return SetFloatSetting(
                value,
                v => Properties.Settings.Default.tableImageScaleMultiplier = v,
                "tableImageScaleMultiplier"
            );
        }

        /// <summary>
        /// コラム内画像スケール設定（ColumnImageScaleMultiplier）を取得
        /// </summary>
        /// <returns>スケール倍率（デフォルト: 1.2f）</returns>
        /// <remarks>
        /// 有効範囲: 0.1?10.0
        /// </remarks>
        public static float GetColumnImageScaleMultiplierSetting()
        {
            return GetFloatSetting(
                () => Properties.Settings.Default.columnImageScaleMultiplier,
                DefaultColumnImageScaleMultiplier,
                "columnImageScaleMultiplier"
            );
        }

        /// <summary>
        /// コラム内画像スケール設定（ColumnImageScaleMultiplier）を保存
        /// </summary>
        /// <param name="value">設定値</param>
        /// <returns>保存に成功した場合はtrue、範囲外の値の場合はfalse</returns>
        public static bool SetColumnImageScaleMultiplierSetting(float value)
        {
            return SetFloatSetting(
                value,
                v => Properties.Settings.Default.columnImageScaleMultiplier = v,
                "columnImageScaleMultiplier"
            );
        }

        /// <summary>
        /// 出力画像の最大幅設定を取得
        /// </summary>
        /// <returns>最大幅（デフォルト: 1024）</returns>
        /// <remarks>
        /// 有効範囲: 100?4096
        /// </remarks>
        public static int GetMaxOutputWidthSetting()
        {
            return GetIntSetting(
                () => Properties.Settings.Default.maxOutputWidth,
                DefaultMaxOutputWidth,
                "maxOutputWidth"
            );
        }

        /// <summary>
        /// 出力画像の最大幅設定を保存
        /// </summary>
        /// <param name="value">設定値</param>
        /// <returns>保存に成功した場合はtrue、範囲外の値の場合はfalse</returns>
        public static bool SetMaxOutputWidthSetting(int value)
        {
            return SetIntSetting(
                value,
                v => Properties.Settings.Default.maxOutputWidth = v,
                "maxOutputWidth"
            );
        }

        /// <summary>
        /// 出力画像の最大高さ設定を取得
        /// </summary>
        /// <returns>最大高さ（デフォルト: 1024）</returns>
        /// <remarks>
        /// 有効範囲: 100?4096
        /// </remarks>
        public static int GetMaxOutputHeightSetting()
        {
            return GetIntSetting(
                () => Properties.Settings.Default.maxOutputHeight,
                DefaultMaxOutputHeight,
                "maxOutputHeight"
            );
        }

        /// <summary>
        /// 出力画像の最大高さ設定を保存
        /// </summary>
        /// <param name="value">設定値</param>
        /// <returns>保存に成功した場合はtrue、範囲外の値の場合はfalse</returns>
        public static bool SetMaxOutputHeightSetting(int value)
        {
            return SetIntSetting(
                value,
                v => Properties.Settings.Default.maxOutputHeight = v,
                "maxOutputHeight"
            );
        }

        /// <summary>
        /// 設定ボタンの表示/非表示設定を取得
        /// </summary>
        /// <returns>設定ボタンを表示する場合はtrue、表示しない場合はfalse（デフォルト: true）</returns>
        /// <remarks>
        /// この設定により、リボンの「画像出力設定」ボタンの表示/非表示を制御できます。
        /// 今後のバージョンによっては設定ボタンを使わない場合に、非表示にできます。
        /// </remarks>
        public static bool GetShowSettingsButtonSetting()
        {
            try
            {
                return Properties.Settings.Default.showSettingsButton;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"showSettingsButton設定の取得に失敗: {ex.Message}");
                return DefaultShowSettingsButton;
            }
        }

        /// <summary>
        /// 設定ボタンの表示/非表示設定を保存
        /// </summary>
        /// <param name="value">設定値</param>
        public static void SetShowSettingsButtonSetting(bool value)
        {
            try
            {
                Properties.Settings.Default.showSettingsButton = value;
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"showSettingsButton設定の保存に失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 浮動小数点数の設定を取得する汎用メソッド
        /// </summary>
        private static float GetFloatSetting(Func<float> getter, float defaultValue, string settingName)
        {
            try
            {
                float value = getter();
                
                // 妥当な範囲内かチェック
                if (value >= MinScaleMultiplier && value <= MaxScaleMultiplier)
                {
                    return value;
                }

                System.Diagnostics.Debug.WriteLine(
                    $"{settingName}設定の値が範囲外です: {value} (有効範囲: {MinScaleMultiplier}?{MaxScaleMultiplier})");
                return defaultValue;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"{settingName}設定の取得に失敗: {ex.Message}");
                return defaultValue;
            }
        }

        /// <summary>
        /// 浮動小数点数の設定を保存する汎用メソッド
        /// </summary>
        private static bool SetFloatSetting(float value, Action<float> setter, string settingName)
        {
            try
            {
                // 妥当な範囲内かチェック
                if (value < MinScaleMultiplier || value > MaxScaleMultiplier)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"{settingName}設定の値が範囲外です: {value} (有効範囲: {MinScaleMultiplier}?{MaxScaleMultiplier})");
                    return false;
                }

                setter(value);
                Properties.Settings.Default.Save();
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"{settingName}設定の保存に失敗: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 整数値の設定を取得する汎用メソッド
        /// </summary>
        private static int GetIntSetting(Func<int> getter, int defaultValue, string settingName)
        {
            try
            {
                int value = getter();
                
                // 妥当な範囲内かチェック
                if (value >= MinOutputSize && value <= MaxOutputSize)
                {
                    return value;
                }

                System.Diagnostics.Debug.WriteLine(
                    $"{settingName}設定の値が範囲外です: {value} (有効範囲: {MinOutputSize}?{MaxOutputSize})");
                return defaultValue;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"{settingName}設定の取得に失敗: {ex.Message}");
                return defaultValue;
            }
        }

        /// <summary>
        /// 整数値の設定を保存する汎用メソッド
        /// </summary>
        private static bool SetIntSetting(int value, Action<int> setter, string settingName)
        {
            try
            {
                // 妥当な範囲内かチェック
                if (value < MinOutputSize || value > MaxOutputSize)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"{settingName}設定の値が範囲外です: {value} (有効範囲: {MinOutputSize}?{MaxOutputSize})");
                    return false;
                }

                setter(value);
                Properties.Settings.Default.Save();
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"{settingName}設定の保存に失敗: {ex.Message}");
                return false;
            }
        }
    }
}
