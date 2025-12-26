var wide = Array("０","１","２","３","４","５","６","７","８","９","ａ","ｂ","ｃ","ｄ","ｅ","ｆ","ｇ","ｈ","ｉ","ｊ","ｋ","ｌ","ｍ","ｎ","ｏ","ｐ","ｑ","ｒ","ｓ","ｔ","ｕ","ｖ","ｗ","ｘ","ｙ","ｚ","ガ","ギ","グ","ゲ","ゴ","ザ","ジ","ズ","ゼ","ゾ","ダ","ヂ","ヅ","デ","ド","バ","ビ","ブ","ベ","ボ","パ","ピ","プ","ペ","ポ","。","「","」","、","ヲ","ァ","ィ","ゥ","ェ","ォ","ャ","ュ","ョ","ッ","ー","ア","イ","ウ","エ","オ","カ","キ","ク","ケ","コ","サ","シ","ス","セ","ソ","タ","チ","ツ","テ","ト","ナ","ニ","ヌ","ネ","ノ","ハ","ヒ","フ","ヘ","ホ","マ","ミ","ム","メ","モ","ヤ","ユ","ヨ","ラ","リ","ル","レ","ロ","ワ","ン");
var narrow = Array("0","1","2","3","4","5","6","7","8","9","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","ｶﾞ","ｷﾞ","ｸﾞ","ｹﾞ","ｺﾞ","ｻﾞ","ｼﾞ","ｽﾞ","ｾﾞ","ｿﾞ","ﾀﾞ","ﾁﾞ","ﾂﾞ","ﾃﾞ","ﾄﾞ","ﾊﾞ","ﾋﾞ","ﾌﾞ","ﾍﾞ","ﾎﾞ","ﾊﾟ","ﾋﾟ","ﾌﾟ","ﾍﾟ","ﾎﾟ","｡","｢","｣","､","ｦ","ｧ","ｨ","ｩ","ｪ","ｫ","ｬ","ｭ","ｮ","ｯ","ｰ","ｱ","ｲ","ｳ","ｴ","ｵ","ｶ","ｷ","ｸ","ｹ","ｺ","ｻ","ｼ","ｽ","ｾ","ｿ","ﾀ","ﾁ","ﾂ","ﾃ","ﾄ","ﾅ","ﾆ","ﾇ","ﾈ","ﾉ","ﾊ","ﾋ","ﾌ","ﾍ","ﾎ","ﾏ","ﾐ","ﾑ","ﾒ","ﾓ","ﾔ","ﾕ","ﾖ","ﾗ","ﾘ","ﾙ","ﾚ","ﾛ","ﾜ","ﾝ");
var highlight = Array("(?:０|0)","(?:１|1)","(?:２|2)","(?:３|3)","(?:４|4)","(?:５|5)","(?:６|6)","(?:７|7)","(?:８|8)","(?:９|9)","(?:ａ|a)","(?:ｂ|b)","(?:ｃ|c)","(?:ｄ|d)","(?:ｅ|e)","(?:ｆ|f)","(?:ｇ|g)","(?:ｈ|h)","(?:ｉ|i)","(?:ｊ|j)","(?:ｋ|k)","(?:ｌ|l)","(?:ｍ|m)","(?:ｎ|n)","(?:ｏ|o)","(?:ｐ|p)","(?:ｑ|q)","(?:ｒ|r)","(?:ｓ|s)","(?:ｔ|t)","(?:ｕ|u)","(?:ｖ|v)","(?:ｗ|w)","(?:ｘ|x)","(?:ｙ|y)","(?:ｚ|z)","(?:ガ|ｶﾞ)","(?:ギ|ｷﾞ)","(?:グ|ｸﾞ)","(?:ゲ|ｹﾞ)","(?:ゴ|ｺﾞ)","(?:ザ|ｻﾞ)","(?:ジ|ｼﾞ)","(?:ズ|ｽﾞ)","(?:ゼ|ｾﾞ)","(?:ゾ|ｿﾞ)","(?:ダ|ﾀﾞ)","(?:ヂ|ﾁﾞ)","(?:ヅ|ﾂﾞ)","(?:デ|ﾃﾞ)","(?:ド|ﾄﾞ)","(?:バ|ﾊﾞ)","(?:ビ|ﾋﾞ)","(?:ブ|ﾌﾞ)","(?:ベ|ﾍﾞ)","(?:ボ|ﾎﾞ)","(?:パ|ﾊﾟ)","(?:ピ|ﾋﾟ)","(?:プ|ﾌﾟ)","(?:ペ|ﾍﾟ)","(?:ポ|ﾎﾟ)","(?:。|｡)","(?:「|｢)","(?:」|｣)","(?:、|､)","(?:ヲ|ｦ)","(?:ァ|ｧ)","(?:ィ|ｨ)","(?:ゥ|ｩ)","(?:ェ|ｪ)","(?:ォ|ｫ)","(?:ャ|ｬ)","(?:ュ|ｭ)","(?:ョ|ｮ)","(?:ッ|ｯ)","(?:ー|ｰ)","(?:ア|ｱ)","(?:イ|ｲ)","(?:ウ|ｳ)","(?:エ|ｴ)","(?:オ|ｵ)","(?:カ|ｶ)","(?:キ|ｷ)","(?:ク|ｸ)","(?:ケ|ｹ)","(?:コ|ｺ)","(?:サ|ｻ)","(?:シ|ｼ)","(?:ス|ｽ)","(?:セ|ｾ)","(?:ソ|ｿ)","(?:タ|ﾀ)","(?:チ|ﾁ)","(?:ツ|ﾂ)","(?:テ|ﾃ)","(?:ト|ﾄ)","(?:ナ|ﾅ)","(?:ニ|ﾆ)","(?:ヌ|ﾇ)","(?:ネ|ﾈ)","(?:ノ|ﾉ)","(?:ハ|ﾊ)","(?:ヒ|ﾋ)","(?:フ|ﾌ)","(?:ヘ|ﾍ)","(?:ホ|ﾎ)","(?:マ|ﾏ)","(?:ミ|ﾐ)","(?:ム|ﾑ)","(?:メ|ﾒ)","(?:モ|ﾓ)","(?:ヤ|ﾔ)","(?:ユ|ﾕ)","(?:ヨ|ﾖ)","(?:ラ|ﾗ)","(?:リ|ﾘ)","(?:ル|ﾙ)","(?:レ|ﾚ)","(?:ロ|ﾛ)","(?:ワ|ﾜ)","(?:ン|ﾝ)");

// MutationObserver機能用のグローバル変数
var currentSearchValue = ""; // 現在の検索キーワード
var mutationObserver = null; // MutationObserverインスタンス
var debounceTimer = null; // DOM変更用のデバウンスタイマー

// jQueryセレクタキャッシュ
var $cachedElements = {
  iframe: null,
  searchField: null,
  searchResultItemsBlock: null,
  searchResultsEnd: null,
  searchMsg: null,
  searchInput: null
};

// 文字変換マップ（効率化のため初期化時に正規表現を構築）
var characterMappings = (function() {
  // 全角→半角の変換マップを作成
  var wideToNarrowMap = {};
  var wideToHighlightMap = {};
  
  for (var i = 0; i < wide.length; i++) {
    wideToNarrowMap[wide[i]] = narrow[i];
    wideToHighlightMap[wide[i]] = highlight[i];
  }
  
  // 全角文字の正規表現パターンを作成（エスケープ処理を含む）
  var wideCharsPattern = wide.map(function(char) {
    return char.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }).join('|');
  
  var wideToNarrowRegex = new RegExp(wideCharsPattern, 'g');
  
  // ハイライト用の正規表現パターンマップを作成
  var highlightRegexMap = {};
  for (var i = 0; i < highlight.length; i++) {
    highlightRegexMap[highlight[i]] = new RegExp(highlight[i], "gm");
  }
  
  return {
    wideToNarrowMap: wideToNarrowMap,
    wideToHighlightMap: wideToHighlightMap,
    wideToNarrowRegex: wideToNarrowRegex,
    highlightRegexMap: highlightRegexMap,
    // 効率的な変換関数
    convertWideToNarrow: function(text) {
      return text.replace(wideToNarrowRegex, function(match) {
        return wideToNarrowMap[match] || match;
      });
    },
    // ハイライトパターンの変換
    applyHighlightPattern: function(text) {
      var result = text;
      for (var pattern in highlightRegexMap) {
        if (highlightRegexMap.hasOwnProperty(pattern)) {
          result = result.replace(highlightRegexMap[pattern], pattern);
        }
      }
      return result;
    }
  };
})();

// キャッシュを初期化
function initializeCachedElements() {
  $cachedElements.iframe = $("iframe.topic");
  $cachedElements.searchField = $(".wSearchField");
  $cachedElements.searchResultItemsBlock = $(".wSearchResultItemsBlock");
  $cachedElements.searchResultsEnd = $(".wSearchResultsEnd");
  $cachedElements.searchMsg = $("#searchMsg");
  $cachedElements.searchInput = $(".search-input", document);
}

// キャッシュされた要素を取得（存在チェック付き）
function getCachedElement(key) {
  if (!$cachedElements[key] || $cachedElements[key].length === 0) {
    // キャッシュが無効な場合は再取得
    switch(key) {
      case 'iframe':
        $cachedElements.iframe = $("iframe.topic");
        break;
      case 'searchField':
        $cachedElements.searchField = $(".wSearchField");
        break;
      case 'searchResultItemsBlock':
        $cachedElements.searchResultItemsBlock = $(".wSearchResultItemsBlock");
        break;
      case 'searchResultsEnd':
        $cachedElements.searchResultsEnd = $(".wSearchResultsEnd");
        break;
      case 'searchMsg':
        $cachedElements.searchMsg = $("#searchMsg");
        break;
      case 'searchInput':
        $cachedElements.searchInput = $(".search-input", document);
        break;
    }
  }
  return $cachedElements[key];
}

function selectorEscape(val){
  return val.replace(/[-\/\\^$*+?.()|[\]{}\!]/g, '\\$&');
}

// 文字列を正規化（全角→半角カナ変換、小文字化）
function normalizeForSearch(text) {
  var normalized = text.toLowerCase();
  // 効率的な一括変換
  return characterMappings.convertWideToNarrow(normalized);
}

// 検索語を正規化してエスケープ
function prepareSearchWords(searchValue) {
  var searchWordTmp = searchValue.split("　").join(" ").trim();
  searchWordTmp = searchWordTmp.split("  ").join(" ");
  searchWordTmp = searchWordTmp.toLowerCase();
  // 効率的な一括変換
  searchWordTmp = characterMappings.convertWideToNarrow(searchWordTmp);
  
  var searchWord = searchWordTmp.split(" ");
  for(var i = 0; i < searchWord.length; i++) {
    searchWord[i] = selectorEscape(searchWord[i].replace(">", "&gt;").replace("<", "&lt;"));
  }
  return searchWord;
}

// ハイライト用の正規表現パターンを生成
function createHighlightPattern(searchWords) {
  var highlightWord = searchWords.join("|");
  // 効率的なパターン変換
  return characterMappings.applyHighlightPattern(highlightWord);
}

// HTMLエンティティを復元
function decodeHtmlEntities(html) {
  var regnbsp = new RegExp("&nbsp;(?=[^<>]*<)", "gm");
  var reggt = new RegExp("&gt;(?=[^<>]*<)", "gm");
  var reglt = new RegExp("&lt;(?=[^<>]*<)", "gm");
  var regquot = new RegExp("&quot;(?=[^<>]*<)", "gm");
  var regamp = new RegExp("&amp;(?=[^<>]*<)", "gm");
  
  return html
    .replace(regnbsp, "　")
    .replace(reggt, ">")
    .replace(reglt, "<")
    .replace(regquot, '"')
    .replace(regamp, "&");
}

// iframeのbody要素を取得
function getIframeBody() {
  var $iframe = getCachedElement('iframe');
  if ($iframe.length === 0) return null;
  return $iframe.contents().find("body");
}

// iframeコンテンツにハイライトを適用
function applyHighlight(searchValue) {
  var $body = getIframeBody();
  if (!$body) return;
  
  var searchWords = prepareSearchWords(searchValue);
  var highlightPattern = createHighlightPattern(searchWords);
  
  var reg = new RegExp("("+highlightPattern+")(?=[^<>]*<)", "gmi");
  var html = $body.html();
  var decodedHtml = decodeHtmlEntities(html);
  var highlightedHtml = decodedHtml.replace(reg, "<font class='keyword' style='color:rgb(0, 0, 0); background-color:rgb(252, 255, 0);'>$1</font>");
  
  $body.html(highlightedHtml);
}

// キーワードのハイライトを削除
function removeHighlight() {
  var $body = getIframeBody();
  if (!$body) return;
  
  $body.find(".keyword").each(function() {
    for(var i = 0; i < $(this)[0].childNodes.length; i++) {
      this.parentNode.insertBefore($(this)[0].childNodes[i], this);
    }
    $(this).remove();
  });
}

// iframeコンテンツ用のMutationObserverをセットアップ
function setupMutationObserver() {
  disconnectMutationObserver();
  
  try {
    var $iframe = getCachedElement('iframe');
    if ($iframe.length === 0) return;
    
    var iframeDocument = $iframe[0].contentDocument || $iframe[0].contentWindow.document;
    if (!iframeDocument || !iframeDocument.body) return;
    
    mutationObserver = new MutationObserver(function(mutations) {
      if (!currentSearchValue || currentSearchValue.trim() === "") {
        return;
      }
      
      var shouldReHighlight = false;
      for (var i = 0; i < mutations.length; i++) {
        var mutation = mutations[i];
        if (mutation.type === 'childList') {
          var addedKeywords = false;
          for (var j = 0; j < mutation.addedNodes.length; j++) {
            var node = mutation.addedNodes[j];
            if (node.nodeType === Node.ELEMENT_NODE) {
              if (node.classList && node.classList.contains('keyword')) {
                addedKeywords = true;
                break;
              }
              if (node.querySelector && node.querySelector('.keyword')) {
                addedKeywords = true;
                break;
              }
            }
          }
          
          if (addedKeywords && mutation.addedNodes.length === 1) {
            continue;
          }
          
          if (mutation.addedNodes.length > 0 || mutation.removedNodes.length > 0) {
            shouldReHighlight = true;
            break;
          }
        } else if (mutation.type === 'characterData') {
          shouldReHighlight = true;
          break;
        }
      }
      
      if (shouldReHighlight) {
        debouncedReHighlight();
      }
    });
    
    mutationObserver.observe(iframeDocument.body, {
      childList: true,
      subtree: true,
      characterData: true,
      attributes: false
    });
    
    console.debug("iframeコンテンツ用のMutationObserverをセットアップしました");
  } catch (error) {
    console.warn("MutationObserverのセットアップに失敗しました:", error);
  }
}

// デバウンスされた再ハイライト処理
function debouncedReHighlight() {
  if (debounceTimer) {
    clearTimeout(debounceTimer);
  }
  
  debounceTimer = setTimeout(function() {
    reHighlightAfterDomChange();
  }, 500);
}

// DOM変更後の再ハイライト処理
function reHighlightAfterDomChange() {
  if (!currentSearchValue || currentSearchValue.trim() === "") {
    return;
  }
  
  try {
    disconnectMutationObserver();
    removeHighlight();
    applyHighlight(currentSearchValue);
    
    setTimeout(function() {
      setupMutationObserver();
    }, 100);
    
    console.debug("DOM変更後に検索語を再ハイライトしました");
  } catch (error) {
    console.warn("DOM変更後の再ハイライトに失敗しました:", error);
    setupMutationObserver();
  }
}

// MutationObserverを切断
function disconnectMutationObserver() {
  if (mutationObserver) {
    mutationObserver.disconnect();
    mutationObserver = null;
  }
  
  if (debounceTimer) {
    clearTimeout(debounceTimer);
    debounceTimer = null;
  }
}

// カスタムの:contains()セレクタ（正規化された検索用）
$.expr[':'].containsNormalized = function(elem, index, match) {
  var normalizedElemText = normalizeForSearch($(elem).text());
  var normalizedSearchText = normalizeForSearch(match[3]);
  return normalizedElemText.indexOf(normalizedSearchText) >= 0;
};

$(function(){
  // 要素のキャッシュを初期化
  initializeCachedElements();
  
  $(document).on("click", "ul.toc li.book", function() {
    if($(this).children("a[href='#'],a[href='javascript:void 0;']").length == 0)
    {
      $(this).children("a").each(function(){
        location.href=location.href.replace(location.hash,"")+"#t="+$(this).attr("href");
      });
    }
  });
  
  getCachedElement('searchField').each(function() {
    $(this).off();
  });
  
  $(document).on("keyup", ".wSearchField", function(){
    var $searchResultItemsBlock = getCachedElement('searchResultItemsBlock');
    var $searchResultsEnd = getCachedElement('searchResultsEnd');
    var $searchMsg = getCachedElement('searchMsg');
    
    if($(this).val() == "")
    {
      $searchResultItemsBlock.html("");
      $searchResultsEnd.addClass("rh-hide");
      $searchResultsEnd.attr("hidden", "");
      $searchMsg.html("2つ以上の語句を入力して検索する場合は、スペース（空白）で区切ります。");
      removeHighlight();
      currentSearchValue = ""; // 現在の検索値をクリア
      disconnectMutationObserver(); // クリア時にObserverを切断
    }
    else
    {
      $searchMsg.html("");
      currentSearchValue = $(this).val(); // 現在の検索値を保存
      var searchWordTmp = $(this).val().replace(/(.*?)(?:　| )+(.*?)/g, "$1 $2").trim();

      // 正規化（全角→半角カナ、小文字化）
      searchWordTmp = normalizeForSearch(searchWordTmp);
      var searchWord = searchWordTmp.split(" ");
      var searchQuery = "";
      for(i = 0; i < searchWord.length; i ++)
      {
        searchQuery += ":containsNormalized(" + searchWord[i] + ")";
      }
      
      var findItems = searchWords.find(".search_word"+searchQuery);
      if(findItems.length != 0)
      {
        $searchResultsEnd.removeClass("rh-hide");
        $searchResultsEnd.removeAttr("hidden");
        $searchResultItemsBlock.html("");
        findItems.each(function() {
          var displayText = $(this).parent().find(".displayText").text();
          $searchResultItemsBlock.append($("<div class='wSearchResultItem'><a class='nolink' href='./"+$(this).parent().attr("id")+".html'><div class='wSearchResultTitle'>"+$(this).parent().find(".search_title").html()+"</div></a><div class='wSearchContent'><span class='wSearchContext'>"+displayText+"</span></div></div>"));
        });
        removeHighlight();
        applyHighlight(currentSearchValue);
      }
      else
      {
        removeHighlight();
        $searchResultsEnd.addClass("rh-hide");
        $searchResultsEnd.attr("hidden", "");
        $searchResultItemsBlock.html("");
        displayText = "検索条件に一致するトピックはありません。";
        $searchResultItemsBlock.append($("<div class='wSearchResultItem'><div class='wSearchContent'><span class='wSearchContext'>"+displayText+"</span></div></div>"));
      }
    }
  });
  
  getCachedElement('iframe').on("load", function(){
    var $searchInput = getCachedElement('searchInput');
    var $searchField = getCachedElement('searchField');
    
    if($searchInput.is(":not(.rh-hide)") && ($searchField.val() != ""))
    {
      var searchValue = $searchField.val();
      currentSearchValue = searchValue; // 現在の検索値を保存
      applyHighlight(searchValue);
    }
    
    // iframeコンテンツ変更用のMutationObserverをセットアップ
    setupMutationObserver();
  });
});
