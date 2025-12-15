// Utils.Progress.cs

using System;
using System.Diagnostics;

namespace MJS_fileJoin
{
    internal partial class Utils
    {
        /// <summary>
        /// プログレスバーとラベルを統一的に管理するヘルパークラス
        /// </summary>
        public class ProgressScope : IDisposable
        {
            private readonly MainForm _form;
            private readonly string _labelText;
            private readonly int _maximum;
            private readonly Stopwatch _stopwatch;
            private bool _disposed;

            /// <summary>
            /// プログレスバーとラベルを初期化します
            /// </summary>
            /// <param name="form">MainFormのインスタンス</param>
            /// <param name="labelText">表示するラベルテキスト</param>
            /// <param name="maximum">プログレスバーの最大値</param>
            public ProgressScope(MainForm form, string labelText, int maximum)
            {
                _form = form ?? throw new ArgumentNullException(nameof(form));
                _labelText = labelText;
                _maximum = maximum;
                _disposed = false;
                _stopwatch = Stopwatch.StartNew();

                // 初期化
                _form.label10.Text = _labelText;
                _form.progressBar1.Maximum = _maximum;
                _form.progressBar1.Value = 0;

                Trace.WriteLine($"[Progress] 開始: {_labelText} (最大値: {_maximum})");
            }

            /// <summary>
            /// プログレスバーの値を設定します
            /// </summary>
            /// <param name="value">設定する値</param>
            public void SetValue(int value)
            {
                if (_disposed)
                    return;

                if (value < 0)
                    value = 0;
                if (value > _maximum)
                    value = _maximum;

                _form.progressBar1.Value = value;
            }

            /// <summary>
            /// プログレスバーの値を増加させます
            /// </summary>
            /// <param name="increment">増加量（省略時は1）</param>
            public void Increment(int increment = 1)
            {
                if (_disposed)
                    return;

                int newValue = _form.progressBar1.Value + increment;
                SetValue(newValue);
            }

            /// <summary>
            /// プログレスバーを最大値まで進めて完了します
            /// </summary>
            public void Complete()
            {
                if (_disposed)
                    return;

                _form.progressBar1.Value = _maximum;
                _stopwatch.Stop();
                Trace.WriteLine($"[Progress] 完了: {_labelText} (処理時間: {_stopwatch.ElapsedMilliseconds}ms)");
            }

            /// <summary>
            /// リソースを解放し、プログレスバーをリセットします
            /// </summary>
            public void Dispose()
            {
                if (_disposed)
                    return;

                _disposed = true;

                // プログレスバーを最大値まで進める
                if (_form.progressBar1.Value < _maximum)
                {
                    _form.progressBar1.Value = _maximum;
                }

                _stopwatch.Stop();
                Trace.WriteLine($"[Progress] 終了: {_labelText} (処理時間: {_stopwatch.ElapsedMilliseconds}ms)");

                // プログレスバーとラベルをリセット
                _form.progressBar1.Value = 0;
                _form.label10.Text = "";
            }
        }

        /// <summary>
        /// プログレスバーとラベルを初期化して処理を開始します
        /// </summary>
        /// <param name="form">MainFormのインスタンス</param>
        /// <param name="labelText">表示するラベルテキスト</param>
        /// <param name="maximum">プログレスバーの最大値</param>
        /// <returns>ProgressScopeのインスタンス</returns>
        public static ProgressScope BeginProgress(MainForm form, string labelText, int maximum)
        {
            return new ProgressScope(form, labelText, maximum);
        }
    }
}
