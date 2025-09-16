# MJS Wordプラグインの概要
MJS Wordプラグインは、Word原稿の編集支援、およびHTML出力のためのツールです。

## アーキテクチャ
- VSTO（Visual Studio Tools for Office）ベースのWordアドイン
- .NET Framework 4.8 / C# 7.3
- COM Interopを使用したOffice連携
- XML/HTML処理による文書変換
- リボンUI による操作インターフェース

## プラグインの機能
Wordプラグインがインストールされている状態でWordを起動すると、メニューに「MJSワードプラグイン」タブが追加されます。
「MJSワードプラグイン」をクリックすると、左から順に「書誌情報出力」、「リンク設定」、「スタイルチェック」、「HTML出力」のリボンが表示されます。

### 書誌情報出力
- Word原稿の目次構成の情報である「書誌情報」を、テキストファイル（書誌情報ファイルと呼びます）として出力（更新）します。
- 書誌情報ファイルは、Word原稿（doc）と同じ階層にある「headerFile」フォルダ内に保存されます。

#### 書誌情報出力に関連するファイル
- BookInfoButton.cs: 書誌情報出力ボタンの処理。makeBookInfo()メソッドを呼び出してメイン処理を実行
- RibbonMJS.MakeBookInfo.cs: 書誌情報作成のメイン処理。文書内の見出し段落を解析してブックマークを生成し、新旧書誌情報の比較・更新を実行
- RibbonMJS.MakeBookInfo.Helper.cs: 書誌情報作成のヘルパー関数群。ファイル名チェック、ブックマーク操作、新旧書誌情報の比較処理
- RibbonMJS.CheckDocInfo.cs: 新旧書誌情報の比較処理。項目の追加・削除・変更・ID不一致・タイトル変更など8種類の変更パターンを検出し、比較結果リストを生成
- RibbonMJS.CheckSortInfo.cs: 書誌情報比較結果のソート処理。項番階層（4レベル）に基づく適切な並び順での比較結果表示機能
- RibbonMJS.MakeBookInfo.HeaderFile.cs: ヘッダーファイル（書誌情報ファイル）の読み込み・書き込み処理とファイルアクセス制御
- BookInfo.cs: 書誌情報のデフォルト値入力用ダイアログフォーム。2桁の数値入力と全角・半角変換機能
- HeadingInfo.cs: 見出し情報を格納するデータクラス（項番・タイトル・ID・マージ先情報）

### リンク設定
- Word原稿の編集補助機能として、他の項目へのリンク（参照）を設定します。
- 書誌情報ファイルから項目を読み込み、相対パス計算とURL生成を行ってハイパーリンクを作成します。

#### リンク設定に関連するファイル
- SetLink.cs: リンク設定ダイアログフォーム
- SetLinkButton.cs: リンク設定ボタンの処理。SetLinkダイアログフォームを表示してリンク設定機能を起動

### スタイルチェック
- HTML出力の前準備として、Word原稿内で規定以外のスタイルが使われていないかをチェックします。

#### スタイルチェックに関連するファイル
- StyleCheckButton.cs: スタイルチェックのメイン処理。テンプレートから許可されたMJSスタイルリストを取得し、ドキュメント全体の検証を実行
- StyleCheckButton.HandleProcess.cs: スタイルチェック完了後の結果処理（成功・失敗・停止時のメッセージ表示とボタン制御）
- StyleCheckButton.Paragraphs.cs: 各段落のスタイル検証処理。MJSスタイル適合性チェックと手順番号リセット用スタイルの整合性検証
- StyleCheckButton.NonInlineShape.cs: 図形・画像の配置チェック処理。行内配置以外のシェイプや描画キャンバスの配置エラーを検出

### HTML出力
- Word原稿の内容を、HTMLで出力します。
- HTMLファイルはすべてwebhelpフォルダに保存されます。
- Word原稿から取得した画像は、webhelpフォルダ内のpictフォルダに保存されます。

#### HTML出力に関連するファイル
- GenerateHTMLButton.cs: HTML出力のメイン処理。WordドキュメントからWebヘルプ形式のHTMLコンテンツを生成し、表紙画像の抽出から検索機能まで包括的な変換処理を実行
- GenerateHTMLButton.CopyDocumentToHtml.cs: WordドキュメントをHTML変換用に複製する処理。クリップボード経由でドキュメント全体をコピーし、新規ドキュメントに貼り付け
- GenerateHTMLButton.StyleProcessor.cs: Word文書から抽出したCSSスタイル定義の解析処理。mso-style-name属性を基に章分割クラスやスタイル名辞書を生成
- GenerateHTMLButton.ProcessHTML.cs: 一時的に保存されたHTMLファイルの読み込みと前処理。文字エンコーディング修正やHTML構造の正規化を実行
- GenerateHTMLButton.HtmlTemplate1.cs: 個別ページ用HTMLテンプレート生成。パンくずリスト、目次階層、検索機能を含む標準ページレイアウトの構築
- GenerateHTMLButton.HtmlCoverTemplate.cs: 表紙ページ用HTMLテンプレート生成。製品ロゴ、タイトル、商標情報を含む表紙レイアウトの構築
- GenerateHTMLButton.IdxHtmlTemplate.cs: インデックス（目次）ページ用HTMLテンプレート生成。全体のナビゲーション構造とフレーム設定を含むメインページの構築
- GenerateHTMLButton.CollectInfo.cs: Word文書から表紙・商標・バージョン情報の収集処理。特定スタイルの段落からタイトルや著作権情報を抽出
- GenerateHTMLButton.CollectMergeScript.cs: 見出し結合情報を書誌情報ファイル（headerFile）から収集し、HTML出力時のページマージ処理用辞書を生成
- GenerateHTMLButton.Helper.cs: HTML生成で共通利用するヘルパー関数群。パス処理、ファイル操作、表紙選択ダイアログ、例外処理など
- GenerateHTMLButton.CopyImagesFromAppDataLocalTemp.cs: AppData/Local/Tempフォルダから画像ファイルを検索し、webhelp/pictフォルダに適切なファイル名でコピー
- RibbonMJS.InnerNode.cs: HTML変換時のXMLノード内部処理。Word文書の各要素（表・図形・スタイル）をHTML要素に変換するメイン処理
- RibbonMJS.InnerNode.Helper.cs: InnerNode.csのヘルパー関数群。操作手順・Q&A・選択肢・箇条書き・表・コラムなど各種MJSスタイルの専用HTML変換処理


#### XML・HTML変換関連のファイル
- GenerateHTMLButton.XMLProcessDocument.cs: WordのHTML出力をXML形式に変換し、目次・本文構造の解析と分割処理を実行
- GenerateHTMLButton.XMLBuildTocBody.cs: XML形式の文書データから目次構造と本文ページを構築。章分割やページ階層の生成処理
- GenerateHTMLButton.XMLExportTocAsJsFiles.cs: 目次データをJavaScript形式で出力し、Webヘルプシステムのナビゲーション機能を生成

#### 検索機能関連のファイル
- GenerateHTMLButton.SearchIndex.cs: 検索対象となるHTMLページと検索用インデックスファイル（search.js）を生成し、検索語彙の抽出と索引化を実行
- GenerateHTMLButton.RemoveSearchBlock.cs: 指定されたタイトルのページから検索ブロックを削除し、検索対象外コンテンツの除外処理を実行

## その他のファイル

### 共通ユーティリティクラス（Utils）
- Utils.FileIO.cs: リスト型変数の内容をテキストファイルに書き込む汎用メソッドと、ファイル操作の共通処理を提供
- Utils.TextProcessing.cs: 全角文字から半角文字への変換機能。数字・英字・記号の文字種変換辞書とConvertWideToNarrowメソッドを提供
- Utils.RemoveSpanTagFromHtml.cs: HTMLファイルから不要なspanタグを削除する機能。HTMLファイル単体またはフォルダ内の一括処理に対応し、属性なしのシンプルなspanタグのみを対象として中身のテキストは保持
- Utils.ExtractImagesFromWord.cs: WordドキュメントからEnhMetaFileBitsを使用してインライン図形・フローティング図形・キャンバス図形を高品質で抽出する処理。抽出した画像に対応するマーカーをWord文書内に挿入し、後続のHTML処理で画像パスを正確に参照できるよう制御
- Utils.ExtractImagesFromWord.CheckStyle.cs: 画像抽出時のスタイル判定処理。MJS特定スタイル（画像、手順内、本文内、コラム内、表内、処理フロー等）の判定とスタイルベース強制抽出・スキップ制御、表紙セクション判定機能を提供
- Utils.ExtractImagesFromWord.Info.cs: 画像抽出結果の統計情報生成とテキストファイル出力機能。抽出した画像の種別・サイズ・数量に関する詳細レポートを生成
- Utils.ExtractImagesFromWord.InsertMarker.cs: 抽出した画像の位置にマーカーテキストを挿入する処理。インライン図形とフローティング図形それぞれに対応した[IMAGEMARKER:xxx]形式のマーカー挿入機能
- Utils.ProcessImageMarkers.cs: HTML出力後のwebhelpディレクトリ内で、[IMAGEMARKER:xxx]パターンを検索し対応する画像ファイルへのsrc属性を更新する処理。画像マーカーとHTMLのimg要素を関連付けて、抽出された画像への正確なリンクを生成
- Utils.RemoveAllImageMarkers.cs: Word文書から画像マーカーテキストを削除する処理。HTML出力完了後のクリーンアップ機能
- Utils.RemoveImageMarkersFromSearchJs.cs: 検索機能用JavaScript（search.js）ファイルから画像マーカーテキストを削除し、検索対象外コンテンツとして除外する処理

### 設定・初期化関連のファイル
- RibbonMJS.Config.cs: HTML出力用パス一覧の準備、各種定数・パターンの定義、検索条件設定などプラグイン全体の設定機能を提供

### UI・システム機能
- RibbonMJS.ClearClipboard.cs: クリップボードの安全なクリア処理。COMException対応のリトライ機能付きクリップボード操作
- RibbonMJS.Designer.cs: リボンUI（MJSワードプラグインタブ）のデザイナー自動生成コード。ボタン配置・イベントハンドラー設定・リソース管理


## Copilot リファクタリング指示書

### 一般的なガイドライン
- 指定がない限り、既存の機能を維持したままコードを変更してください。
- 変数名やメソッド名は、処理の内容が把握できるように記述してください。
- ネストを減らすため、早期リターン（early return）を心がけてください。
- マジックナンバーは避け、定数やenumを使用してください。
- 適切に例外処理を行い、必要に応じてエラーログを出力してください。
- 現状でログの記録が必要だと思った箇所には、適切なログ出力を追加してください。

### C#/.NET 固有
- 右辺から型が明確な場合のみ `var` を使用してください。
- `IDisposable` なオブジェクトは `using` 文で確実に破棄してください。
- コレクションやオブジェクト初期化子を積極的に利用してください。
- 文字列連結には可能な限り文字列補間（$"...") を使ってください。
- LINQなどの機能を活用し、可読性の高いコードを書いてください。

### アドイン／Interop 固有
- メモリリーク防止のため、COMオブジェクトは確実に解放してください。
- Interopのロジックはヘルパーメソッドやクラスにまとめてください。
