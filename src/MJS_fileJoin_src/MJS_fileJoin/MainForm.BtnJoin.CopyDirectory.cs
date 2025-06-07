using System.IO;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        private void CopyDirectoryRecursive(string sourceDir, string destinationDir)
        {
            Directory.CreateDirectory(destinationDir);

            foreach (string file in Directory.GetFiles(sourceDir))
            {
                string destinationFile = Path.Combine(destinationDir, Path.GetFileName(file));
                if (Path.GetFileName(file).ToLower().Contains("image"))
                {

                }
                else if (Path.GetFileName(file).ToLower().Contains(".html"))
                {

                }
                else
                {
                    File.Copy(file, destinationFile, true);
                }
            }

            foreach (string subDir in Directory.GetDirectories(sourceDir))
            {
                string destSubDir = Path.Combine(destinationDir, Path.GetFileName(subDir));
                CopyDirectoryRecursive(subDir, destSubDir);
            }
        }

        public static void CopyDirectoryWithOverwriteOption(string sourceDir, string destinationDir, bool overwrite)
        {
            // コピー先のディレクトリがなければ作成する
            if (!Directory.Exists(destinationDir))
            {
                Directory.CreateDirectory(destinationDir);
                File.SetAttributes(destinationDir, File.GetAttributes(sourceDir));
                overwrite = true;
            }

            // コピー元のディレクトリにあるすべてのファイルをコピーする
            if (overwrite)
            {
                foreach (string copyFrom in Directory.GetFiles(sourceDir))
                {
                    string copyTo = Path.Combine(destinationDir, Path.GetFileName(copyFrom));
                    File.Copy(copyFrom, copyTo, true);
                }
            }
            else
            {
                foreach (string copyFrom in Directory.GetFiles(sourceDir))
                {
                    string copyTo = Path.Combine(destinationDir, Path.GetFileName(copyFrom));
                    if (!File.Exists(copyTo))
                    {
                        File.Copy(copyFrom, copyTo, false);
                    }
                }
            }

            // コピー元のディレクトリをすべてコピーする (再帰)
            foreach (string copyFrom in Directory.GetDirectories(sourceDir))
            {
                string copyTo = Path.Combine(destinationDir, Path.GetFileName(copyFrom));
                CopyDirectoryWithOverwriteOption(copyFrom, copyTo, overwrite);
            }
        }
    }
}
