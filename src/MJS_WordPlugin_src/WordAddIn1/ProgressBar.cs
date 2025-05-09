using System;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class ProgressBar : Form
    {
        public static ProgressBar mInstance = null;

        public ProgressBar()
        {
            InitializeComponent();
        }

        public static new void Show()
        {
            // BeginInvokeで実処理を別スレッド実行
            Action showProc = new Action(ShowProcess);
            IAsyncResult async = showProc.BeginInvoke(null, null);

            // そのままメインスレッドは処理が流れるため、別スレッドでインスタンスが生成されるまで待つ
            while (true)
            {
                if (mInstance != null)
                {
                    break;
                }
            }
            return;
        }

        private static void ShowProcess()
        {
            if (mInstance == null || mInstance.IsDisposed == true)
            {
                mInstance = new ProgressBar();
            }

           // Showだと処理が流れてスレッドが終了してしまうので、ShowDialogで表示して
           // 別スレッド側は処理を待つ
           ((Form)mInstance).ShowDialog();
        }

        // クローズ処理
        public static new void Close()
        {
            if (mInstance == null || mInstance.IsDisposed == true) return;

            if (mInstance.InvokeRequired == true)
            {
                mInstance.Invoke(new Action(Close));
            }
            else
            {
                ((Form)mInstance).Dispose();
                ((Form)mInstance).Close();
            }
        }

        // プログレスバー最大値の設定
        public static void SetProgressBar(int iMax)
        {
            if (mInstance == null || mInstance.IsDisposed == true) return;

            if (mInstance.InvokeRequired)
            {
                mInstance.Invoke(new Action<int>(SetProgressBar), new object[]{iMax});
            }
            else
            {
                mInstance.progressBar1.Maximum = iMax;
                mInstance.label4.Text = Convert.ToString(iMax);
            }
        }


        // プログレスバー現在値の設定
        public static void SetProgressBarValue(int iVal)
        {
            if (mInstance == null || mInstance.IsDisposed == true) return;

            if (mInstance.InvokeRequired)
            {
                try
                {
                    mInstance.Invoke(new Action<int>(SetProgressBarValue), new object[] { iVal });
                }
                catch
                {
                    return;
                }
            }
            else
            {
                mInstance.progressBar1.Value = iVal;
                mInstance.label2.Text = Convert.ToString(iVal);
            }
        }

        // 経過時間
        public static void ProgressTime(System.TimeSpan ts)
        {
            if (mInstance == null || mInstance.IsDisposed == true) return;

            if (mInstance.InvokeRequired)
            {
                try
                {
                    mInstance.Invoke(new Action<System.TimeSpan>(ProgressTime), new object[] { ts });
                }
                catch
                {
                    return;
                }
            }
            else
            {
                mInstance.label7.Text = ts.ToString("mm\\:ss");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProgressBar.Close();
        }
    }
}
