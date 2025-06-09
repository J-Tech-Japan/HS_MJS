using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        /*
        1. アクティブドキュメントと同じ階層に webhelpフォルダがあるか確認
        2. webhelpフォルダの中にある .htmlファイルをすべて調べる
        3. AppData/Local/Temp にある参照先の画像をすべて pict フォルダにコピー
        4. imgタグの src 属性を新しい参照先に書き換える
        */

        private void CopyImagesFromAppDataLocalTemp(string activeDocumentPath)
        {
            // アクティブドキュメントと同じ階層にwebhelpフォルダがあるか確認
            var docDir = Path.GetDirectoryName(activeDocumentPath);
            var webhelpDir = Path.Combine(docDir, "webhelp");
            if (!Directory.Exists(webhelpDir)) return;

            // pictフォルダのパスを決定
            var imgFromTempDir = Path.Combine(webhelpDir, "pict");
            if (!Directory.Exists(imgFromTempDir))
            {
                Directory.CreateDirectory(imgFromTempDir);
            }

            // webhelpフォルダ内の.htmlファイルをすべて取得
            var htmlFiles = Directory.GetFiles(webhelpDir, "*.html", SearchOption.TopDirectoryOnly);

            // imgタグのsrc属性にAppData/Local/Tempを含むものを抽出
            var imgTagRegex = new Regex("<img([^>]+)src=[\"']([^\"']+AppData/Local/Temp[^\"']+)[\"']([^>]*)>", RegexOptions.IgnoreCase);

            foreach (var htmlFile in htmlFiles)
            {
                var htmlContent = File.ReadAllText(htmlFile);
                var matches = imgTagRegex.Matches(htmlContent);
                bool changed = false;

                // 画像コピーとsrc書き換え
                string replaced = imgTagRegex.Replace(htmlContent, match =>
                {
                    var src = match.Groups[2].Value;
                    string filePath = src;

                    // file:/// 形式の場合はローカルパスに変換
                    if (filePath.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
                    {
                        filePath = filePath.Substring("file:///".Length);
                        filePath = filePath.Replace('/', '\\');
                    }

                    // デコード（スペースや日本語などのエンコード対応）
                    filePath = Uri.UnescapeDataString(filePath);

                    string fileName = Path.GetFileName(filePath);
                    var destPath = Path.Combine(imgFromTempDir, fileName);

                    // 画像をpictにコピー
                    try
                    {
                        if (File.Exists(filePath))
                        {
                            File.Copy(filePath, destPath, true);
                            File.Delete(filePath); // コピー後に削除
                        }
                    }
                    catch (Exception)
                    {
                    }

                    // imgタグのsrcを書き換え
                    changed = true;
                    string attr1 = match.Groups[1].Value;
                    string attr2 = match.Groups[3].Value;
                    string newSrc = $"pict/{fileName}";
                    return $"<img{attr1}src=\"{newSrc}\"{attr2}>";
                });

                if (changed)
                {
                    File.WriteAllText(htmlFile, replaced, Encoding.UTF8);
                }
            }
        }
    }
}
