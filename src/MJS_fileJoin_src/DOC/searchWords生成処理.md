# HTML結合におけるsearchWords生成処理

このドキュメントでは、MJS_fileJoinツールのHTML結合機能において、検索用データ（`searchWords`）がどのように生成されるかの処理フローを説明します。

## 概要

`searchWords`は、結合されたHTMLファイル群に対して全文検索機能を提供するためのXMLデータです。各HTMLファイルの内容を解析し、検索可能な形式に変換して`search.js`ファイルに埋め込まれます。

## 処理フロー

### 1. 初期化（MainForm.BtnJoin.cs）

```csharp
XmlDocument searchWords = new System.Xml.XmlDocument();
searchWords.LoadXml("<div class='search'></div>");
```

HTML結合処理の開始時に、空の`<div class='search'></div>`としてXMLドキュメントが初期化されます。

### 2. HTMLファイル処理とデータ蓄積

各結合元フォルダのHTMLファイルを順次処理する際、以下の流れでデータが蓄積されます。

#### 2.1 HTMLファイルの処理（MainForm.BtnJoin.ProcessHtmlFiles.cs）

`ProcessHtmlFiles`メソッドが、各結合元フォルダのHTMLファイルを処理します：

```csharp
private void ProcessHtmlFiles(
    string htmlDir,
    string outputDir,
    int picCount,
    List<string> lsfiles,
    XmlNode objTocRoot,
    XmlDocument objToc,
    XmlDocument searchWords,
    List<string> errorList)
{
    foreach (DataRow selRow in bookInfo[htmlDir].Select("Column1 = true"))
    {
        // HTMLファイルのコピーと画像処理
        string selHtml = CopyHtmlAndImages(htmlFile, outFile, htmlDir, outputDir, picCount);
        
        // パンくずリスト・目次・検索データ生成
        selHtml = GenerateBreadcrumbsAndToc(selHtml, selRow, objTocRoot, objToc, searchWords);
        
        // 相対リンク修正
        selHtml = FixRelativeLinks(selHtml, lsfiles);
        
        // 処理済みHTMLを保存
        using (var sw = new StreamWriter(outFile, false, Encoding.UTF8))
        {
            sw.Write(selHtml);
        }
    }
}
```

#### 2.2 検索データの生成（GenerateBreadcrumbsAndToc メソッド）

`GenerateBreadcrumbsAndToc`メソッド内で、各HTMLファイルから検索用データが抽出されます：

##### ステップ1: 検索用div要素の追加

```csharp
searchWords.DocumentElement.AppendChild(searchWords.CreateElement("div"));
((XmlElement)searchWords.DocumentElement.LastChild).SetAttribute("id", selRow["Column4"].ToString());
```

トピックIDを持つdiv要素が追加されます。

##### ステップ2: HTML本文からテキスト抽出

```csharp
string bodyStr = Regex.Replace(
    Regex.Replace(
        Regex.Replace(
            Regex.Replace(selHtml, "\r?\n", ""),              // 改行削除
            "^.+<body[^>]*>(.+?)</body>.*$", "$1", RegexOptions.Multiline),  // body部分抽出
        @"<div style=""text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;"">.+?</div>", ""),  // パンくずリスト削除
    "<.+?>", "");  // すべてのHTMLタグ削除
```

以下の処理が行われます：
- 改行文字の削除
- `<body>`タグ内のコンテンツ抽出
- パンくずリスト領域の除外
- すべてのHTMLタグの削除

##### ステップ3: エスケープ処理と表示テキスト作成

```csharp
string searchText = bodyStr.Replace("&", "&amp;").Replace("<", "&lt;");
string displayText = searchText;
if (searchText.Length >= 90)
{
    displayText = displayText.Substring(0, 90) + " ...";
}
```

- XMLエンティティのエスケープ処理
- 表示用テキストを90文字に切り詰め

##### ステップ4: 全角→半角変換と正規化

```csharp
string[] wide = { "０", "１", "２", ... };  // 全角文字配列
string[] narrow = { "0", "1", "2", ... };   // 半角文字配列

for (int p = 0; p < wide.Length; p++)
{
    searchText = Regex.Replace(searchText, wide[p], narrow[p]);
}
searchText = searchText.ToLower();
```

検索精度を向上させるため、以下の変換が行われます：
- 全角英数字→半角英数字
- 全角カタカナ→半角カタカナ
- 全角記号→半角記号
- 濁音・半濁音の正規化（例：ガ→ｶﾞ）
- すべて小文字に変換

##### ステップ5: XML構造の構築

```csharp
searchWords.DocumentElement.LastChild.InnerXml = 
    "<div class='search_breadcrumbs'>" + breadcrumb.Replace("&", "&amp;").Replace("<", "&lt;") + 
    "</div><div class='search_title'>" + title.Replace("&", "&amp;").Replace("<", "&lt;") +
    "</div><div class='displayText'>" + displayText +
    "</div><div class='search_word'>" + searchText + "</div>";
```

最終的な検索データのXML構造：

```xml
<div id="トピックID">
  <div class='search_breadcrumbs'>パンくずリスト（例：製品A > 機能B > 設定）</div>
  <div class='search_title'>タイトル（例：初期設定の方法）</div>
  <div class='displayText'>表示用テキスト（90文字まで）...</div>
  <div class='search_word'>正規化された検索用テキスト（半角・小文字）</div>
</div>
```

### 3. XMLのエスケープ処理（MainForm.BtnJoin.cs）

すべてのHTMLファイルの処理が完了後、searchWords XML全体に対してエスケープ処理が行われます：

```csharp
string processedSearchWordsXml = Regex.Replace(
    searchWords.OuterXml, 
    @"(?<=>)([^<]*?)""([^<]*?)(?=<)", 
    "$1&quot;$2", 
    RegexOptions.Singleline)
    .Replace("'", "&apos;");
```

以下の文字がエスケープされます：
- `"` → `&quot;`
- `'` → `&apos;`

これにより、JavaScriptの文字列リテラル内に安全に埋め込めるようになります。

### 4. search.jsファイルへの埋め込み（define.cs）

`GetSearchJsWithReplacedWords`メソッドが、処理済みXMLをJavaScriptファイルに埋め込みます：

```csharp
private static string GetSearchJsWithReplacedWords(List<string> htmlDirs, string newSearchWordsXml)
{
    // 結合元フォルダから最初に見つかったsearch.jsを使用
    foreach (string htmlDir in htmlDirs)
    {
        string searchJsPath = Path.Combine(htmlDir, "search.js");
        
        if (File.Exists(searchJsPath))
        {
            try
            {
                string searchJsContent = File.ReadAllText(searchJsPath, Encoding.UTF8);
                
                // var searchWords = $('...'); の部分を新しい内容で置き換え
                string pattern = @"var\s+searchWords\s*=\s*\$\('.*?'\);?";
                string replacement = $"var searchWords = $('{newSearchWordsXml}');";
                
                string result = Regex.Replace(searchJsContent, pattern, replacement, RegexOptions.Singleline);
                
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"search.js読み込みエラー ({searchJsPath}): {ex.Message}");
            }
        }
    }

    // どの結合元フォルダにもsearch.jsが見つからない場合は、最小限のsearch.jsを生成
    System.Diagnostics.Trace.WriteLine("search.jsが見つかりませんでした。最小限の内容を生成します。");
    return $"var searchWords = $('{newSearchWordsXml}');";
}
```

処理の流れ：
1. 結合元フォルダから既存の`search.js`を検索
2. 最初に見つかった`search.js`を読み込み
3. 正規表現で`var searchWords = $('...');`部分を検索
4. 新しいXMLデータで置き換え
5. `search.js`が見つからない場合は最小限の内容を生成

### 5. ファイル出力（MainForm.BtnJoin.cs）

最後に、生成された`search.js`ファイルを出力ディレクトリに保存します：

```csharp
string searchJsPath = Path.Combine(tbOutputDir.Text, exportDir, "search.js");
sw = new StreamWriter(searchJsPath, false, Encoding.UTF8);
sw.Write(searchJsContent);
sw.Close();
```

## データ構造の例

### 入力HTMLファイル（例）

```html
<!DOCTYPE html>
<html>
<head>
    <meta name="topic-breadcrumbs" content="">
    <title>初期設定の方法</title>
</head>
<body>
    <div style="text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;">
        製品A &gt; 機能B &gt; 設定
    </div>
    <h1>初期設定の方法</h1>
    <p>この機能を使用するには、まず初期設定を行う必要があります。</p>
</body>
</html>
```

### 生成されるsearchWordsデータ（例）

```xml
<div class='search'>
  <div id="topic001">
    <div class='search_breadcrumbs'>製品A &gt; 機能B &gt; 設定</div>
    <div class='search_title'>初期設定の方法</div>
    <div class='displayText'>初期設定の方法 この機能を使用するには、まず初期設定を行う必要があります。</div>
    <div class='search_word'>初期設定の方法 この機能を使用するには、まず初期設定を行う必要があります。</div>
  </div>
  <!-- 他のトピックも同様に続く -->
</div>
```

### 最終的なsearch.jsファイル（例）

```javascript
var searchWords = $('<div class=\'search\'><div id="topic001"><div class=\'search_breadcrumbs\'>製品A &gt; 機能B &gt; 設定</div><div class=\'search_title\'>初期設定の方法</div><div class=\'displayText\'>初期設定の方法 この機能を使用するには...</div><div class=\'search_word\'>初期設定の方法 この機能を使用するには...</div></div></div>');

// 検索機能の実装（既存のsearch.jsから継承）
function performSearch(keyword) {
    // 検索ロジック
}
```

## まとめ

searchWords生成処理は、以下の5つのステップで実行されます：

1. **初期化**: 空のXMLドキュメントを作成
2. **データ蓄積**: 各HTMLファイルから検索データを抽出・蓄積
   - HTML本文からテキスト抽出
   - 全角→半角変換と正規化
   - パンくずリスト、タイトル、本文をXML構造化
3. **エスケープ処理**: XML全体をJavaScript埋め込み用にエスケープ
4. **search.js埋め込み**: 既存のsearch.jsテンプレートにデータを統合
5. **ファイル出力**: 出力ディレクトリに保存

この仕組みにより、複数のHTMLファイルの内容が1つの`searchWords`変数に集約され、ブラウザ上でのクライアントサイド全文検索が可能になります。

## 関連ファイル

- `MJS_fileJoin\MainForm.BtnJoin.cs` - 結合処理のメインロジック
- `MJS_fileJoin\MainForm.BtnJoin.ProcessHtmlFiles.cs` - HTMLファイル処理とsearchWords生成
- `MJS_fileJoin\define.cs` - search.js生成とテンプレート統合
