// Publish project specific data
(function() {
rh = window.rh;
model = rh.model;

rh.consts('DEFAULT_TOPIC', encodeURI("#初期画面.htm".substring(1)));
rh.consts('HOME_FILEPATH', encodeURI("index.html"));
rh.consts('START_FILEPATH', encodeURI('index.html'));
rh.consts('HELP_ID', 'CFF3E8CE-4750-496C-8FF0-9AAEB1C57FD3' || 'preview');
rh.consts('LNG_STOP_WORDS', ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "after", "all", "along", "already", "also", "am", "among", "an", "and", "another", "any", "are", "at", "be", "because", "been", "between", "but", "by", "can", "do", "does", "doesn", "done", "each", "either", "for", "from", "get", "has", "have", "here", "how", "i", "if", "in", "into", "is", "isn", "it", "like", "may", "maybe", "more", "must", "need", "non", "not", "of", "ok", "okay", "on", "or", "other", "rather", "re", "s", "same", "see", "so", "some", "such", "t", "than", "that", "the", "their", "them", "then", "there", "these", "they", "this", "those", "to", "too", "unless", "use", "used", "using", "ve", "want", "was", "way", "were", "what", "when", "when", "whenever", "where", "whether", "which", "will", "with", "within", "without", "yet", "you", "your"]);
rh.consts('LNG_SUBSTR_SEARCH', 0);

model.publish(rh.consts('KEY_DIR'), "ltr");
model.publish(rh.consts('KEY_LNG_NAME'), "ja_JP");
model.publish(rh.consts('KEY_LNG'), {"Reset":"リセット","SearchResultsPerScreen":"検索結果 / ページ","SyncToc":"SyncToc","HomeButton":"ホーム","WebSearchButton":"WebSearch","Welcome_header":"ヘルプセンターへようこそ","ApplyTip":"適用","HighlightSearchResults":"検索結果をハイライト","GlossaryFilterTerms":"用語を検索","WebSearch":"WebSearch","Show":"表示","Welcome_text":"お問い合わせの内容","EnableAndSearch":"すべての検索語句を含む結果を表示","ShowAll":"すべて表示","Next":">>","Print":"印刷","NoScriptErrorMsg":"このページを表示するには、ブラウザーで JavaScript サポートを有効にしてください。","PreviousLabel":"前へ","Hide":"非表示","Search":"検索","Contents":"目次","ShowHide":"表示 / 非表示","Canceled":"キャンセルされました","favoritesLabel":"お気に入り","EndOfResults":"検索結果の最後です。","Loading":"読み込み中...","SidebarToggleTip":"展開 / 折りたたむ","ContentFilterChanged":"コンテンツフィルターが変更されています、再検索してください","Logo":"ロゴ","Logo/Author":"Powered By","JS_alert_LoadXmlFailed":"エラー : xml ファイルの読み込みに失敗しました。","favoritesNameLabel":"名前","Copyright":"© Copyright 2017. All rights reserved.","SearchTitle":"検索","Searching":"検索中...","Disabled Next":">>","nofavoritesFound":"お気に入りとしてマークしたページがありません。","unsetAsFavorite":"お気に入りから削除","Cancel":"キャンセル","JS_alert_InitDatabaseFailed":"エラー : データベースの初期化に失敗しました。","FilterIntro":"フィルターを選択してください :","ResultsFoundText":"%2 に %1 個の結果が見つかりました","UnknownError":"不明なエラー","Seperate":"|","Index":"索引","setAsFavorite":"お気に入りとして設定","setAsFavorites":"お気に入りに追加","TopicsNotFound":"トピックが見つかりません。","SearchPageTitle":"検索結果","Glossary":"用語集","SearchButtonTitle":"検索","Filter":"フィルター","HideAll":"すべて非表示","TableOfContents":"目次","NextLabel":"次へ","Disabled Prev":"<<","Back":"戻る","SearchOptions":"検索オプション","OpenLinkInNewTab":"新規タブで開く","Prev":"<<","ShowTopicInContext":"このマニュアルの目次を表示するには、ここをクリックしてください。","FavoriteBoxTitle":"お気に入り","ToTopTip":"トップへ移動","NavTip":"メニュー","IeCompatibilityErrorMsg":"このページは、Internet Explorer 8 以前のバージョンでは表示できません.","IndexFilterKewords":"キーワードを検索","JS_alert_InvalidExpression_1":"入力した内容は有効な式ではありません。"});

model.publish(rh.consts('KEY_HEADER_DEFAULT_TITLE_COLOR'), "#ffffff");
model.publish(rh.consts('KEY_HEADER_DEFAULT_BACKGROUND_COLOR'), "#025172");
model.publish(rh.consts('KEY_LAYOUT_DEFAULT_FONT_FAMILY'), "\"Trebuchet MS\", Arial, sans-serif");

model.publish(rh.consts('KEY_HEADER_TITLE'), "");
model.publish(rh.consts('KEY_HEADER_TITLE_COLOR'), "#ffffff");
model.publish(rh.consts('KEY_HEADER_BACKGROUND_COLOR'), "#509de6");
model.publish(rh.consts('KEY_HEADER_LOGO_PATH'), "");
model.publish(rh.consts('KEY_LAYOUT_FONT_FAMILY'), "\"Trebuchet MS\", Arial, sans-serif");
model.publish(rh.consts('KEY_HEADER_HTML'), "<div class='topic-header'>\
  <div class='logo' onClick='rh._.redirectToLayout()'>\
    <img src='#{logo}' />\
  </div>\
  <div class='nav'>\
    <div class='title' title='#{title}'>\
      <span onClick='rh._.redirectToLayout()'>#{title}</span>\
    </div>\
    <div class='gotohome' title='#{tooltip}' onClick='rh._.redirectToLayout()'>\
      <span>#{label}</span>\
    </div></div>\
  </div>\
<div class='topic-header-shadow'></div>\
");
model.publish(rh.consts('KEY_HEADER_CSS'), ".topic-header { background-color: #{background-color}; color: #{color}; width: calc(100%); height: 3em; position: fixed; left: 0; top: 0; font-family: #{font-family}; display: table; box-sizing: border-box; }\
.topic-header-shadow { height: 3em; width: 100%; }\
.logo { cursor: pointer; padding: 0.2em; height: calc(100% - 0.4em); text-align: center; display: table-cell; vertical-align: middle; }\
.logo img { max-height: 100%; display: block; }\
.nav { width: 100%; display: table-cell; }\
.title { width: 40%; height: 100%; float: left; line-height: 3em; cursor: pointer; }\
.gotohome { width: 60%; float: left; text-align: right; height: 100%; line-height: 3em; cursor: pointer; }\
.title span, .gotohome span { padding: 0em 1em; white-space: nowrap; text-overflow: ellipsis; overflow: hidden; display: block; }");

})();