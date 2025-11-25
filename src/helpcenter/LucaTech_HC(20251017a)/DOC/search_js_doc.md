```js
var searchWords = $('省略');

var wide = Array("省略");

var narrow = Array("省略");

var hilight = Array("(?:０|0)","(?:１|1)","(?:２|2)","(?:３|3)","(?:４|4)","(?:５|5)","(?:６|6)","(?:７|7)","(?:８|8)","(?:９|9)","(?:Ａ|A|ａ|a)","(?:Ｂ|B|ｂ|b)","(?:Ｃ|C|ｃ|c)","(?:Ｄ|D|ｄ|d)","(?:Ｅ|E|ｅ|e)","(?:Ｆ|F|ｆ|f)","(?:Ｇ|G|ｇ|g)","(?:Ｈ|H|ｈ|h)","(?:Ｉ|I|ｉ|i)","(?:Ｊ|J|ｊ|j)","(?:Ｋ|K|ｋ|k)","(?:Ｌ|L|ｌ|l)","(?:Ｍ|M|ｍ|m)","(?:Ｎ|N|ｎ|n)","(?:Ｏ|O|ｏ|o)","(?:Ｐ|P|ｐ|p)","(?:Ｑ|Q|ｑ|q)","(?:Ｒ|R|ｒ|r)","(?:Ｓ|S|ｓ|s)","(?:Ｔ|T|ｔ|t)","(?:Ｕ|U|ｕ|u)","(?:Ｖ|V|ｖ|v)","(?:Ｗ|W|ｗ|w)","(?:Ｘ|X|ｘ|x)","(?:Ｙ|Y|ｙ|y)","(?:Ｚ|Z|ｚ|z)","(?:ガ|ｶﾞ)","(?:ギ|ｷﾞ)","(?:グ|ｸﾞ)","(?:ゲ|ｹﾞ)","(?:ゴ|ｺﾞ)","(?:ザ|ｻﾞ)","(?:ジ|ｼﾞ)","(?:ズ|ｽﾞ)","(?:ゼ|ｾﾞ)","(?:ゾ|ｿﾞ)","(?:ダ|ﾀﾞ)","(?:ヂ|ﾁﾞ)","(?:ヅ|ﾂﾞ)","(?:デ|ﾃﾞ)","(?:ド|ﾄﾞ)","(?:バ|ﾊﾞ)","(?:ビ|ﾋﾞ)","(?:ブ|ﾌﾞ)","(?:ベ|ﾍﾞ)","(?:ボ|ﾎﾞ)","(?:パ|ﾊﾟ)","(?:ピ|ﾋﾟ)","(?:プ|ﾌﾟ)","(?:ペ|ﾍﾟ)","(?:ポ|ﾎﾟ)","(?:。|｡)","(?:「|｢)","(?:」|｣)","(?:、|､)","(?:ヲ|ｦ)","(?:ァ|ｧ)","(?:ィ|ｨ)","(?:ゥ|ｩ)","(?:ェ|ｪ)","(?:ォ|ｫ)","(?:ャ|ｬ)","(?:ュ|ｭ)","(?:ョ|ｮ)","(?:ッ|ｯ)","(?:ー|ｰ)","(?:ア|ｱ)","(?:イ|ｲ)","(?:ウ|ｳ)","(?:エ|ｴ)","(?:オ|ｵ)","(?:カ|ｶ)","(?:キ|ｷ)","(?:ク|ｸ)","(?:ケ|ｹ)","(?:コ|ｺ)","(?:サ|ｻ)","(?:シ|ｼ)","(?:ス|ｽ)","(?:セ|ｾ)","(?:ソ|ｿ)","(?:タ|ﾀ)","(?:チ|ﾁ)","(?:ツ|ﾂ)","(?:テ|ﾃ)","(?:ト|ﾄ)","(?:ナ|ﾅ)","(?:ニ|ﾆ)","(?:ヌ|ﾇ)","(?:ネ|ﾈ)","(?:ノ|ﾉ)","(?:ハ|ﾊ)","(?:ヒ|ﾋ)","(?:フ|ﾌ)","(?:ヘ|ﾍ)","(?:ホ|ﾎ)","(?:マ|ﾏ)","(?:ミ|ﾐ)","(?:ム|ﾑ)","(?:メ|ﾒ)","(?:モ|ﾓ)","(?:ヤ|ﾔ)","(?:ユ|ﾕ)","(?:ヨ|ﾖ)","(?:ラ|ﾗ)","(?:リ|ﾘ)","(?:ル|ﾙ)","(?:レ|ﾚ)","(?:ロ|ﾛ)","(?:ワ|ﾜ)","(?:ン|ﾝ)");

// MutationObserver機能のグローバル変数
var currentSearchValue = ""; // 現在の検索キーワード
var mutationObserver = null; // MutationObserverのインスタンス
var debounceTimer = null; // DOM変更用のデバウンスタイマー

/**
 * 正規表現の特殊文字をエスケープする関数
 * @param {string} val - エスケープする文字列
 * @returns {string} エスケープされた文字列
 */
function selectorEscape(val){
  // 正規表現で特別な意味を持つ文字をエスケープ文字（\）でエスケープ
  return val.replace(/[-\/\\^$*+?.()|[\]{}\!]/g, '\\$&');
}

/**
 * iframeコンテンツにハイライトを適用する関数
 * @param {string} searchValue - 検索キーワード
 */
// iframeコンテンツにハイライトを適用
function applyHighlight(searchValue) {
  // 全角スペースを半角スペースに変換し、前後の空白を削除
  var searchWordTmp = searchValue.split("　").join(" ").trim();

  // 連続する空白を単一の空白に変換
  searchWordTmp = searchWordTmp.split("  ").join(" ");
  
  // 全角文字を半角文字に変換（wideからnarrowへのマッピング）
  for(var i = 0; i < wide.length; i++) {
    searchWordTmp = searchWordTmp.replace(wide[i], narrow[i]);
  }
  
  // スペースで区切って単語の配列に分割
  var searchWord = searchWordTmp.split(" ");
  
  // 各単語をエスケープして正規表現で安全に使える形に変換
  for(var i = 0; i < searchWord.length; i++) {
    // HTMLエンティティをエスケープ
    searchWord[i] = selectorEscape(searchWord[i].replace(">", "&gt;").replace("<", "&lt;"));
  }
  
  // 検索単語をOR演算子で結合（例: "word1|word2|word3"）
  var hilightWord = searchWord.join("|");
  
  // hilightWordをより柔軟な検索パターンに変換
  // 全角・半角の両方にマッチするパターンに置換
  for(var i = 0; i < hilight.length; i++) {
    var reg = new RegExp(hilight[i], "gm");
    hilightWord = hilightWord.replace(reg, hilight[i]);
  }

  // 各種正規表現パターンを定義（HTMLタグの外側の文字のみを対象）
  var reg = new RegExp("("+hilightWord+")(?=[^<>]*<)", "gm");  // メインのハイライト用正規表現
  var regnbsp = new RegExp("&nbsp;(?=[^<>]*<)", "gm");         // &nbsp;を全角スペースに変換
  var reggt = new RegExp("&gt;(?=[^<>]*<)", "gm");             // &gt;を>に変換
  var reglt = new RegExp("&lt;(?=[^<>]*<)", "gm");             // &lt;を<に変換
  var regquot = new RegExp("&quot;(?=[^<>]*<)", "gm");         // &quot;を"に変換
  var regamp = new RegExp("&amp;(?=[^<>]*<)", "gm");           // &amp;を&に変換
  
  // iframeの内容を取得し、各種HTMLエンティティを変換してからハイライトを適用
  $("iframe.topic").contents().find("body").html(
    $("iframe.topic").contents().find("body").html()
      .replace(regnbsp, "　")        // &nbsp;を全角スペースに
      .replace(reggt, ">")           // &gt;を>に
      .replace(reglt, "<")           // &lt;を<に
      .replace(regquot, '"')         // &quot;を"に
      .replace(regamp, "&")          // &amp;を&に
      .replace(reg, "<font class='keyword' style='color:rgb(0, 0, 0); background-color:rgb(252, 255, 0);'>$1</font>")  // キーワードをハイライト
  );
}

/**
 * 以前に適用されたキーワードハイライトを削除する関数
 * ハイライト用のfontタグを除去し、テキストのみを残す
 */
// キーワードのハイライトを削除
function removeHighlight() {
  // iframeコンテンツ内の.keywordクラス要素を検索
  $("iframe.topic").contents().find(".keyword").each(function() {
    // keywordタグの子要素（テキストノード）を親要素に移動
    for(var i = 0; i < $(this)[0].childNodes.length; i++) {
      this.parentNode.insertBefore($(this)[0].childNodes[i], this);
    }
    // 空になったkeywordタグを削除
    $(this).remove();
  });
}

/**
 * iframeコンテンツのDOM変更を監視するMutationObserverを設定する関数
 * 動的に変更されるコンテンツに対しても検索ハイライトを維持する
 */
// iframeコンテンツ用のMutationObserverを設定
function setupMutationObserver() {
  // 既存のオブザーバーがある場合は切断
  disconnectMutationObserver();
  
  try {
    // iframe要素を取得
    var $iframe = $("iframe.topic");
    if ($iframe.length === 0) return;  // iframeが見つからない場合は終了
    
    // iframeのdocumentオブジェクトを取得
    var iframeDocument = $iframe[0].contentDocument || $iframe[0].contentWindow.document;
    if (!iframeDocument || !iframeDocument.body) return;  // documentまたはbodyが取得できない場合は終了
    
    // MutationObserverのコールバック関数を定義
    mutationObserver = new MutationObserver(function(mutations) {
      // 検索値が空の場合は何もしない
      if (!currentSearchValue || currentSearchValue.trim() === "") {
        return;
      }
      
      var shouldReHighlight = false;  // 再ハイライトが必要かどうかのフラグ
      
      // 各DOM変更をチェック
      for (var i = 0; i < mutations.length; i++) {
        var mutation = mutations[i];
        
        // 子要素の追加・削除の場合
        if (mutation.type === 'childList') {
          var addedKeywords = false;  // キーワード要素が追加されたかのフラグ
          
          // 追加されたノードをチェック
          for (var j = 0; j < mutation.addedNodes.length; j++) {
            var node = mutation.addedNodes[j];
            // 要素ノードの場合
            if (node.nodeType === Node.ELEMENT_NODE) {
              // 追加された要素がkeywordクラスを持つ場合
              if (node.classList && node.classList.contains('keyword')) {
                addedKeywords = true;
                break;
              }
              // 追加された要素の子にkeywordクラスがある場合
              if (node.querySelector && node.querySelector('.keyword')) {
                addedKeywords = true;
                break;
              }
            }
          }
          
          // キーワード要素が1つだけ追加された場合は無視
          // （ハイライト処理自体による変更なので、再ハイライトは不要）
          if (addedKeywords && mutation.addedNodes.length === 1) {
            continue;
          }
          
          // 要素の追加や削除があった場合は再ハイライトが必要
          if (mutation.addedNodes.length > 0 || mutation.removedNodes.length > 0) {
            shouldReHighlight = true;
            break;
          }
        } else if (mutation.type === 'characterData') {
          // テキストデータが変更された場合も再ハイライトが必要
          shouldReHighlight = true;
          break;
        }
      }
      
      // 再ハイライトが必要な場合は遅延実行
      if (shouldReHighlight) {
        debouncedReHighlight();
      }
    });
    
    // MutationObserverを開始（監視対象と設定を指定）
    mutationObserver.observe(iframeDocument.body, {
      childList: true,     // 子要素の追加・削除を監視
      subtree: true,       // 孫要素以下の変更も監視
      characterData: true, // テキストノードの変更を監視
      attributes: false    // 属性の変更は監視しない
    });
    
    console.debug("MutationObserver setup for iframe content");
  } catch (error) {
    console.warn("Failed to setup MutationObserver:", error);
  }
}

/**
 * DOM変更の頻繁な発生に対応するためのデバウンス機能付き再ハイライト関数
 * 短時間に複数回の変更が発生した場合、最後の変更から一定時間後に実行
 */
// デバウンスされた再ハイライト処理
function debouncedReHighlight() {
  // 既存のタイマーがある場合はクリア
  if (debounceTimer) {
    clearTimeout(debounceTimer);
  }
  
  // 500ms後に再ハイライト処理を実行
  debounceTimer = setTimeout(function() {
    reHighlightAfterDomChange();
  }, 500);
}

/**
 * DOM変更後に実行される再ハイライト処理
 * 既存のハイライトを削除してから新しいハイライトを適用
 */
// DOM変更後の再ハイライト処理
function reHighlightAfterDomChange() {
  // 検索値が空の場合は何もしない
  if (!currentSearchValue || currentSearchValue.trim() === "") {
    return;
  }
  
  try {
    // 一時的にMutationObserverを停止（無限ループ防止）
    disconnectMutationObserver();
    
    // 既存のハイライトを削除
    removeHighlight();
    
    // 新しいハイライトを適用
    applyHighlight(currentSearchValue);
    
    // 少し遅れてMutationObserverを再開
    setTimeout(function() {
      setupMutationObserver();
    }, 100);
    
    console.debug("Re-highlighted search terms after DOM change");
  } catch (error) {
    console.warn("Failed to re-highlight after DOM change:", error);
    // エラーが発生してもMutationObserverは再開
    setupMutationObserver();
  }
}

/**
 * MutationObserverとタイマーを停止・クリアする関数
 * メモリリークを防ぐためのクリーンアップ処理
 */
// MutationObserverを切断
function disconnectMutationObserver() {
  // MutationObserverが存在する場合は停止して削除
  if (mutationObserver) {
    mutationObserver.disconnect();
    mutationObserver = null;
  }
  
  // デバウンスタイマーが存在する場合はクリア
  if (debounceTimer) {
    clearTimeout(debounceTimer);
    debounceTimer = null;
  }
}

/**
 * jQuery DOM Ready 関数
 * ページが読み込まれた後に実行される初期化処理
 */
$(function(){
  // 目次（TOC）のブック項目クリック時の処理
  $(document).on("click", "ul.toc li.book", function() {
    // リンクが無効でない場合
    if($(this).children("a[href='#'],a[href='javascript:void 0;']").length == 0)
    {
      // 各リンクに対してページ遷移処理を実行
      $(this).children("a").each(function(){
        location.href=location.href.replace(location.hash,"")+"#t="+$(this).attr("href");
      });
    }
  });
  
  // 検索フィールドの既存イベントを削除（重複回避）
  $(".wSearchField").each(function() {
    $(this).off();
  });
  
  /**
   * 検索フィールドでのキー入力時の処理
   * 検索実行とハイライト処理を行う
   */
  $(document).on("keyup", ".wSearchField", function(){
    // 検索フィールドが空の場合
    if($(this).val() == "")
    {
      // 検索結果をクリア
      $(".wSearchResultItemsBlock").html("");
      $(".wSearchResultsEnd").addClass("rh-hide");
      $(".wSearchResultsEnd").attr("hidden", "");
      
      // 検索ヘルプメッセージを表示
      $("#searchMsg").html("2つ以上の語句を入力して検索する場合は、スペース（空白）で区切ります。");
      
      // ハイライトとオブザーバーをクリア
      removeHighlight();
      currentSearchValue = ""; // 現在の検索値をクリア
      disconnectMutationObserver(); // クリア時にオブザーバーを切断
    }
    else
    {
      // 検索処理開始
      $("#searchMsg").html("");
      currentSearchValue = $(this).val(); // 現在の検索値を保存
      
      // 検索語の正規化処理
      var searchWordTmp = $(this).val().replace(/(.*?)(?:　| )+(.*?)/g, "$1 $2").trim().toLowerCase();
      
      // 全角文字を半角に変換
      for(i = 0; i < wide.length; i ++)
      {
        searchWordTmp = searchWordTmp.split(wide[i]).join(narrow[i]);
      }
      
      // スペースで分割して検索語の配列を作成
      var searchWord = searchWordTmp.split(" ");
      
      // jQueryの:contains()セレクタ用のクエリ文字列を構築
      var searchQuery = "";
      for(i = 0; i < searchWord.length; i ++)
      {
        searchQuery += ":contains(" + searchWord[i] + ")";
      }
      
      // 検索データから該当する項目を検索
      var findItems = searchWords.find(".search_word"+searchQuery);
      
      // 検索結果が見つかった場合
      if(findItems.length != 0)
      {
        // 検索結果エリアを表示
        $(".wSearchResultsEnd").removeClass("rh-hide");
        $(".wSearchResultsEnd").removeAttr("hidden");
        $(".wSearchResultItemsBlock").html("");
        
        // 検索結果を順次追加
        findItems.each(function() {
          var displayText = $(this).parent().find(".displayText").text();
          // 検索結果アイテムのHTMLを構築して追加
          $(".wSearchResultItemsBlock").append($("<div class='wSearchResultItem'><a class='nolink' href='./"+$(this).parent().attr("id")+".html'><div class='wSearchResultTitle'>"+$(this).parent().find(".search_title").html()+"</div></a><div class='wSearchContent'><span class='wSearchContext'>"+displayText+"</span></div></div>"));
        });
        
        // 既存のハイライトを削除してから新しいハイライトを適用
        removeHighlight();
        applyHighlight(currentSearchValue);
      }
      else
      {
        // 検索結果が見つからなかった場合
        removeHighlight();
        $(".wSearchResultsEnd").addClass("rh-hide");
        $(".wSearchResultsEnd").attr("hidden", "");
        $(".wSearchResultItemsBlock").html("");
        
        // "見つからない"メッセージを表示
        displayText = "検索条件に一致するトピックはありません。";
        $(".wSearchResultItemsBlock").append($("<div class='wSearchResultItem'><div class='wSearchContent'><span class='wSearchContext'>"+displayText+"</span></div></div>"));
      }
    }
  });
  
  /**
   * iframeのロード時の処理
   * ページ読み込み完了後にハイライトとMutationObserverを設定
   */
  $("iframe.topic").on("load", function(){
    // 検索入力エリアが表示されており、検索フィールドに値が入っている場合
    if($(".search-input", document).is(":not(.rh-hide)") && ($(".wSearchField", document).val() != ""))
    {
      var searchValue = $(".wSearchField", document).val();
      currentSearchValue = searchValue; // 現在の検索値を保存
      applyHighlight(searchValue);      // ハイライトを適用
    }
    
    // iframeコンテンツの変更用にMutationObserverを設定
    setupMutationObserver();
  });
});
```